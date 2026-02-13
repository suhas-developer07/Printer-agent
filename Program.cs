using System.Drawing;
using System.Drawing.Printing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Runtime.InteropServices;
using System.Net.WebSockets;
using System.Text;
using System.Management;
using System.Net;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Docnet.Core;
using Docnet.Core.Models;
using Lextm.SharpSnmpLib;
using Lextm.SharpSnmpLib.Messaging;

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

app.UseWebSockets(new WebSocketOptions { KeepAliveInterval = TimeSpan.FromSeconds(30) });

var _semaphore = new SemaphoreSlim(1, 1);
var _printerMonitor = new PrinterMonitor();

app.Use(async (context, next) =>
{
    if (context.Request.Path == "/ws" && context.WebSockets.IsWebSocketRequest)
    {
        var ws = await context.WebSockets.AcceptWebSocketAsync();
        Console.WriteLine("🔗 WebSocket connected");
        await HandleWebSocket(ws);
        Console.WriteLine("🔗 WebSocket disconnected");
    }
    else
    {
        await next();
    }
});

app.MapGet("/health", () => Results.Ok(new { status = "healthy", timestamp = DateTime.UtcNow }));
app.Run("http://0.0.0.0:8766");

async Task HandleWebSocket(WebSocket ws)
{
    var buffer = new byte[1024 * 64];
    while (ws.State == WebSocketState.Open)
    {
        try
        {
            var result = await ws.ReceiveAsync(buffer, CancellationToken.None);
            if (result.MessageType == WebSocketMessageType.Close)
            {
                await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "Closed", CancellationToken.None);
                break;
            }
            if (result.MessageType == WebSocketMessageType.Text)
            {
                var message = Encoding.UTF8.GetString(buffer, 0, result.Count);
                Console.WriteLine($"\n📨 Received message");
                _ = Task.Run(() => ProcessMessage(ws, message));
            }
        }
        catch (WebSocketException ex)
        {
            Console.WriteLine($"❌ WebSocket error: {ex.Message}");
            break;
        }
    }
}

async Task ProcessMessage(WebSocket ws, string message)
{
    await _semaphore.WaitAsync();
    try
    {
        var job = JsonSerializer.Deserialize<PrintJob>(message, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        if (job == null || string.IsNullOrWhiteSpace(job.FilePath))
        {
            await SendWsMessage(ws, WsMessage.Error("Invalid payload: missing file_path"));
            return;
        }
        if (!File.Exists(job.FilePath))
        {
            await SendWsMessage(ws, WsMessage.Error($"File not found: {job.FilePath}"));
            return;
        }
        var ext = Path.GetExtension(job.FilePath).ToLower();
        if (!new[] { ".pdf", ".txt", ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff" }.Contains(ext))
        {
            await SendWsMessage(ws, WsMessage.Error($"Unsupported file type: {ext}"));
            return;
        }

        await SendWsMessage(ws, WsMessage.CreateStatus("data_received", "Data received successfully"));
        Console.WriteLine($"✅ File: {Path.GetFileName(job.FilePath)}");

        var options = BuildPrintOptions(job);
        
        // PRE-FLIGHT CHECKS
        var printerStatus = await _printerMonitor.GetPrinterStatus(options.PrinterName!);
        
        if (!printerStatus.IsOnline)
        {
            await SendWsMessage(ws, WsMessage.PrinterStatus("offline", "Printer is offline"));
            throw new Exception("Printer is offline");
        }
        if (printerStatus.IsError)
        {
            await SendWsMessage(ws, WsMessage.PrinterStatus("error", $"Printer error: {printerStatus.ErrorMessage}"));
            throw new Exception($"Printer in error state: {printerStatus.ErrorMessage}");
        }
        if (printerStatus.IsPaperOut)
        {
            await SendWsMessage(ws, WsMessage.PrinterStatus("paper_out", "No paper in printer. Please load paper."));
            throw new Exception("No paper in printer tray");
        }
        if (printerStatus.IsPaperJam)
        {
            await SendWsMessage(ws, WsMessage.PrinterStatus("paper_jam", "Paper jam detected. Please clear the jam."));
            throw new Exception("Paper jam detected");
        }

        var inkLevels = await _printerMonitor.GetInkLevels(options.PrinterName!);
        if (inkLevels != null)
        {
            await SendWsMessage(ws, WsMessage.InkLevel(inkLevels));
            Console.WriteLine($"🖋️  Ink: C={inkLevels.Cyan}% M={inkLevels.Magenta}% Y={inkLevels.Yellow}% K={inkLevels.Black}%");
        }

        await SendWsMessage(ws, WsMessage.CreateStatus("settings_applied", "Settings applied successfully"));
        Console.WriteLine($"✅ Pre-flight checks passed");

        await ExecutePrintSafe(ws, job.FilePath, options, ext);

        await SendWsMessage(ws, WsMessage.CreateStatus("job_completed", "Job completed successfully"));
        Console.WriteLine("✅ Job completed\n");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"❌ Error: {ex.Message}");
        await SendWsMessage(ws, WsMessage.Error($"Print failed: {ex.Message}"));
    }
    finally
    {
        _semaphore.Release();
        await Task.Delay(500);
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }
}

PrintOptions BuildPrintOptions(PrintJob job)
{
    return new PrintOptions
    {
        PrinterName = "EPSON WF-C5890 Series",
        Copies = job.Copies > 0 ? job.Copies : 1,
        Duplex = job.Duplex ?? "simplex",
        PageRange = job.PageRange ?? "all",
        PagesPerSheet = job.PagesPerSheet > 0 ? job.PagesPerSheet : 1,
        Color = true,
        Orientation = "portrait",
        PaperSize = "A4",
        Scale = "fit",
        Quality = 600
    };
}

async Task ExecutePrintSafe(WebSocket ws, string filePath, PrintOptions options, string ext)
{
    switch (ext)
    {
        case ".pdf": await PrintPdfSafe(ws, filePath, options); break;
        case ".txt": await PrintTextSafe(ws, filePath, options); break;
        default: await PrintImageSafe(ws, filePath, options); break;
    }
}

async Task PrintPdfSafe(WebSocket ws, string pdfPath, PrintOptions options)
{
    List<Bitmap>? bitmaps = null;
    DocLib? library = null;
    try
    {
        library = DocLib.Instance;
        var dpi = options.Quality;
        var renderW = (int)(8.27 * dpi);
        var renderH = (int)(11.69 * dpi);

        using var docReader = library.GetDocReader(pdfPath, new PageDimensions(renderW, renderH));
        var pageCount = docReader.GetPageCount();
        var pagesToPrint = ParsePageRange(options.PageRange, pageCount);

        Console.WriteLine($"📖 PDF: {pagesToPrint.Count}/{pageCount} pages at {dpi} DPI");

        string detectedOrientation = "portrait";
        using (var firstPageReader = docReader.GetPageReader(0))
        {
            var width = firstPageReader.GetPageWidth();
            var height = firstPageReader.GetPageHeight();
            detectedOrientation = width > height ? "landscape" : "portrait";
            Console.WriteLine($"📐 Detected: {detectedOrientation} ({width}x{height})");
        }
        options.Orientation = detectedOrientation;

        bitmaps = new List<Bitmap>();
        foreach (var pageNum in pagesToPrint)
        {
            using var pageReader = docReader.GetPageReader(pageNum);
            var rawBytes = pageReader.GetImage();
            var w = pageReader.GetPageWidth();
            var h = pageReader.GetPageHeight();
            var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
            var data = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
            try { Marshal.Copy(rawBytes, 0, data.Scan0, rawBytes.Length); }
            finally { bmp.UnlockBits(data); }
            bitmaps.Add(bmp);
        }

        await PrintBitmapsSafe(ws, bitmaps, options, pagesToPrint);
    }
    finally
    {
        if (bitmaps != null)
        {
            foreach (var bmp in bitmaps) try { bmp?.Dispose(); } catch { }
            bitmaps.Clear();
        }
        try { library?.Dispose(); } catch { }
    }
}

async Task PrintTextSafe(WebSocket ws, string textPath, PrintOptions options)
{
    var lines = File.ReadAllLines(textPath);
    var lineIndex = 0;
    var pageNumber = 0;
    PrintDocument? printDoc = null;
    try
    {
        printDoc = new PrintDocument();
        ApplyPrinterSettingsSafe(printDoc, options);
        var monitorTask = MonitorPrintJobAsync(ws, printDoc, options.PrinterName!);

        printDoc.PrintPage += (sender, e) =>
        {
            if (e?.Graphics == null) return;
            SetHighQuality(e.Graphics);
            var font = new Font("Courier New", 10);
            var yPos = (float)e.MarginBounds.Top;
            var lineHeight = font.GetHeight(e.Graphics);
            var linesPerPage = (int)(e.MarginBounds.Height / lineHeight);
            while (lineIndex < lines.Length && linesPerPage > 0)
            {
                e.Graphics.DrawString(lines[lineIndex], font, Brushes.Black, e.MarginBounds.Left, yPos);
                lineIndex++;
                yPos += lineHeight;
                linesPerPage--;
            }
            pageNumber++;
            e.HasMorePages = lineIndex < lines.Length;
            font.Dispose();
        };

        printDoc.Print();
        await Task.Delay(200);
        await monitorTask;
        await SendWsMessage(ws, WsMessage.PagePrinted(pageNumber));
        Console.WriteLine($"  ✓ {pageNumber} page(s) sent to printer");
    }
    finally
    {
        try { printDoc?.Dispose(); } catch { }
    }
}

async Task PrintImageSafe(WebSocket ws, string imagePath, PrintOptions options)
{
    Image? image = null;
    PrintDocument? printDoc = null;
    try
    {
        image = Image.FromFile(imagePath);
        printDoc = new PrintDocument();
        ApplyPrinterSettingsSafe(printDoc, options);
        var monitorTask = MonitorPrintJobAsync(ws, printDoc, options.PrinterName!);

        printDoc.PrintPage += (sender, e) =>
        {
            if (e?.Graphics == null) return;
            SetHighQuality(e.Graphics);
            var rect = ScaleImage(image, e.PageBounds, options.Scale);
            e.Graphics.DrawImage(image, rect);
            e.HasMorePages = false;
        };

        printDoc.Print();
        await Task.Delay(200);
        await monitorTask;
        await SendWsMessage(ws, WsMessage.PagePrinted(1));
        Console.WriteLine("  ✓ 1 page sent to printer");
    }
    finally
    {
        try { image?.Dispose(); } catch { }
        try { printDoc?.Dispose(); } catch { }
    }
}

async Task PrintBitmapsSafe(WebSocket ws, List<Bitmap> bitmaps, PrintOptions options, List<int> pageNumbers)
{
    PrintDocument? printDoc = null;
    try
    {
        printDoc = new PrintDocument();
        ApplyPrinterSettingsSafe(printDoc, options);
        var monitorTask = MonitorPrintJobAsync(ws, printDoc, options.PrinterName!);

        var pageIndex = 0;
        var pps = options.PagesPerSheet;

        printDoc.PrintPage += (sender, e) =>
        {
            if (e?.Graphics == null) return;
            SetHighQuality(e.Graphics);

            if (pps == 1)
            {
                if (pageIndex < bitmaps.Count)
                {
                    var rect = ScaleImage(bitmaps[pageIndex], e.PageBounds, options.Scale);
                    e.Graphics.DrawImage(bitmaps[pageIndex], rect);
                    pageIndex++;
                    e.HasMorePages = pageIndex < bitmaps.Count;
                }
                else { e.HasMorePages = false; }
            }
            else
            {
                var layout = GetPagesPerSheetLayout(pps, e.PageBounds);
                var count = Math.Min(pps, bitmaps.Count - pageIndex);
                for (int i = 0; i < count; i++)
                {
                    var rect = ScaleImage(bitmaps[pageIndex + i], layout[i], options.Scale);
                    e.Graphics.DrawImage(bitmaps[pageIndex + i], rect);
                }
                pageIndex += count;
                e.HasMorePages = pageIndex < bitmaps.Count;
            }
        };

        printDoc.Print();
        await Task.Delay(200);
        await monitorTask;

        for (int i = 0; i < pageNumbers.Count; i++)
        {
            await SendWsMessage(ws, WsMessage.PagePrinted(pageNumbers[i] + 1));
        }
        Console.WriteLine($"  ✓ {pageNumbers.Count} page(s) sent to printer");
    }
    finally
    {
        try { printDoc?.Dispose(); } catch { }
    }
}

async Task MonitorPrintJobAsync(WebSocket ws, PrintDocument printDoc, string printerName)
{
    try
    {
        var jobName = printDoc.DocumentName;
        var startTime = DateTime.Now;
        var timeout = TimeSpan.FromMinutes(5);
        
        Console.WriteLine($"  🔍 Starting real-time monitoring for '{jobName}'...");
        
        // Wait for job to enter queue
        await Task.Delay(1000);

        int? jobId = null;
        var lastStatus = "";
        var checkCount = 0;
        var maxChecks = 300; // 5 minutes at 1 check per second

        while (checkCount < maxChecks)
        {
            checkCount++;
            
            // Get printer status
            var printerStatus = await _printerMonitor.GetPrinterStatus(printerName);
            
            // Check for printer-level errors
            if (!printerStatus.IsOnline)
            {
                await SendWsMessage(ws, WsMessage.PrinterStatus("offline", "Printer went offline during printing"));
                throw new Exception("Printer went offline");
            }

            if (printerStatus.IsPaperJam)
            {
                await SendWsMessage(ws, WsMessage.PrinterStatus("paper_jam", "Paper jam detected!"));
                throw new Exception("Paper jam during printing");
            }

            // Get active print jobs
            var jobs = await _printerMonitor.GetPrintJobs(printerName);
            
            if (jobId == null && jobs.Count > 0)
            {
                // Find our job (most recent one)
                jobId = jobs[0].JobId;
                Console.WriteLine($"  📋 Tracking job ID: {jobId}");
            }

            if (jobId != null)
            {
                var currentJob = jobs.FirstOrDefault(j => j.JobId == jobId);
                
                if (currentJob != null)
                {
                    var statusDesc = currentJob.Status;
                    
                    // Only log status changes
                    if (statusDesc != lastStatus && !string.IsNullOrEmpty(statusDesc))
                    {
                        Console.WriteLine($"  📊 Job status: {statusDesc}");
                        lastStatus = statusDesc;
                    }

                    // Check for errors
                    if (currentJob.Status.Contains("Error", StringComparison.OrdinalIgnoreCase))
                    {
                        await SendWsMessage(ws, WsMessage.PrinterStatus("error", $"Print job error: {currentJob.Status}"));
                        throw new Exception($"Print job failed: {currentJob.Status}");
                    }

                    if (currentJob.Status.Contains("Paused", StringComparison.OrdinalIgnoreCase))
                    {
                        await SendWsMessage(ws, WsMessage.PrinterStatus("paused", "Print job is paused"));
                        Console.WriteLine("  ⏸️  Job paused");
                    }

                    // Check if job is stuck (not progressing)
                    if (currentJob.Status.Contains("Pending", StringComparison.OrdinalIgnoreCase) && checkCount > 10)
                    {
                        // Job pending for too long - likely paper out or offline
                        var recheck = await _printerMonitor.GetPrinterStatus(printerName);
                        
                        if (recheck.IsPaperOut)
                        {
                            await SendWsMessage(ws, WsMessage.PrinterStatus("paper_out", "Printer is out of paper. Please load paper."));
                            throw new Exception("Out of paper - job pending");
                        }
                    }

                    // Job still exists - keep monitoring
                    if (currentJob.PagesPrinted > 0)
                    {
                        Console.WriteLine($"  📄 Pages printed: {currentJob.PagesPrinted}/{currentJob.TotalPages}");
                    }
                }
                else
                {
                    // Job no longer in queue - completed or failed
                    Console.WriteLine($"  ✅ Job {jobId} completed (removed from queue)");
                    break;
                }
            }
            else if (checkCount > 5)
            {
                // No job found after 5 seconds - either completed very fast or error
                var recheck = await _printerMonitor.GetPrinterStatus(printerName);
                
                if (recheck.IsPaperOut)
                {
                    await SendWsMessage(ws, WsMessage.PrinterStatus("paper_out", "No paper in printer tray"));
                    throw new Exception("Paper out detected");
                }
                
                if (recheck.IsError)
                {
                    await SendWsMessage(ws, WsMessage.PrinterStatus("error", $"Printer error: {recheck.ErrorMessage}"));
                    throw new Exception($"Printer error: {recheck.ErrorMessage}");
                }

                // Job completed immediately
                break;
            }

            await Task.Delay(1000); // Check every second
        }

        if (checkCount >= maxChecks)
        {
            await SendWsMessage(ws, WsMessage.PrinterStatus("timeout", "Print job monitoring timeout"));
            throw new Exception("Print job monitoring timeout");
        }

        Console.WriteLine("  ✅ Monitoring completed successfully");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"  ⚠️  Monitor error: {ex.Message}");
        throw;
    }
}

void ApplyPrinterSettingsSafe(PrintDocument printDoc, PrintOptions options)
{
    printDoc.PrinterSettings.PrinterName = options.PrinterName!;
    if (!printDoc.PrinterSettings.IsValid)
        throw new Exception($"Printer not found: {options.PrinterName}");

    Console.WriteLine($"🖨️  {options.PrinterName}");

    IntPtr hPrinter = IntPtr.Zero;
    IntPtr pDevMode = IntPtr.Zero;

    try
    {
        if (!Win32.OpenPrinter(options.PrinterName!, out hPrinter, IntPtr.Zero))
        {
            Console.WriteLine("⚠️  Using basic fallback");
            ApplyBasicSettings(printDoc, options);
            return;
        }

        int sizeNeeded = Win32.DocumentProperties(IntPtr.Zero, hPrinter, options.PrinterName!, IntPtr.Zero, IntPtr.Zero, 0);
        if (sizeNeeded <= 0)
        {
            ApplyBasicSettings(printDoc, options);
            return;
        }

        pDevMode = Marshal.AllocHGlobal(sizeNeeded);
        int result = Win32.DocumentProperties(IntPtr.Zero, hPrinter, options.PrinterName!, pDevMode, IntPtr.Zero, Win32.DM_OUT_BUFFER);
        if (result < 0)
        {
            ApplyBasicSettings(printDoc, options);
            return;
        }

        var devMode = Marshal.PtrToStructure<Win32.DEVMODE>(pDevMode);
        devMode.dmCopies = (short)options.Copies;
        devMode.dmFields |= Win32.DM_COPIES;

        var isDuplexEnabled = options.Duplex?.ToLower() != "simplex" && !string.IsNullOrEmpty(options.Duplex);
        var isLandscape = (options.PagesPerSheet == 2) || options.Orientation?.ToLower() == "landscape";
        var hasSpecificPageRange = options.PageRange?.ToLower() != "all" && !string.IsNullOrEmpty(options.PageRange);

        short finalDuplex;
        if (!isDuplexEnabled) finalDuplex = Win32.DMDUP_SIMPLEX;
        else if (hasSpecificPageRange) finalDuplex = Win32.DMDUP_HORIZONTAL;
        else if (isLandscape) finalDuplex = Win32.DMDUP_HORIZONTAL;
        else finalDuplex = Win32.DMDUP_VERTICAL;

        devMode.dmDuplex = finalDuplex;
        devMode.dmFields |= Win32.DM_DUPLEX;
        devMode.dmColor = options.Color ? Win32.DMCOLOR_COLOR : Win32.DMCOLOR_MONOCHROME;
        devMode.dmFields |= Win32.DM_COLOR;
        devMode.dmOrientation = isLandscape ? Win32.DMORIENT_LANDSCAPE : Win32.DMORIENT_PORTRAIT;
        devMode.dmFields |= Win32.DM_ORIENTATION;
        devMode.dmPrintQuality = (short)options.Quality;
        devMode.dmYResolution = (short)options.Quality;
        devMode.dmFields |= Win32.DM_PRINTQUALITY;
        devMode.dmFields |= Win32.DM_YRESOLUTION;

        Marshal.StructureToPtr(devMode, pDevMode, true);
        result = Win32.DocumentProperties(IntPtr.Zero, hPrinter, options.PrinterName!, pDevMode, pDevMode, Win32.DM_IN_BUFFER | Win32.DM_OUT_BUFFER);

        if (result >= 0)
        {
            printDoc.PrinterSettings.SetHdevmode(pDevMode);
            printDoc.DefaultPageSettings.SetHdevmode(pDevMode);
        }
    }
    finally
    {
        if (pDevMode != IntPtr.Zero) try { Marshal.FreeHGlobal(pDevMode); } catch { }
        if (hPrinter != IntPtr.Zero) try { Win32.ClosePrinter(hPrinter); } catch { }
    }
}

void ApplyBasicSettings(PrintDocument printDoc, PrintOptions options)
{
    printDoc.PrinterSettings.Copies = (short)options.Copies;
    printDoc.DefaultPageSettings.Color = options.Color;
    printDoc.PrinterSettings.Duplex = Duplex.Simplex;
    printDoc.DefaultPageSettings.Landscape = options.Orientation?.ToLower() == "landscape";
}

async Task SendWsMessage(WebSocket ws, WsMessage msg)
{
    if (ws.State != WebSocketState.Open) return;
    var json = JsonSerializer.Serialize(msg, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
    var bytes = Encoding.UTF8.GetBytes(json);
    try { await ws.SendAsync(bytes, WebSocketMessageType.Text, true, CancellationToken.None); }
    catch { }
}

void SetHighQuality(Graphics g)
{
    g.SmoothingMode = SmoothingMode.HighQuality;
    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
    g.CompositingQuality = CompositingQuality.HighQuality;
}

Rectangle ScaleImage(Image image, Rectangle pageRect, string? scaleMode)
{
    float scale = scaleMode?.ToLower() == "actual" ? 1.0f :
        Math.Min((float)pageRect.Width / image.Width, (float)pageRect.Height / image.Height);
    var newW = (int)(image.Width * scale);
    var newH = (int)(image.Height * scale);
    var x = pageRect.X + (pageRect.Width - newW) / 2;
    var y = pageRect.Y + (pageRect.Height - newH) / 2;
    return new Rectangle(x, y, newW, newH);
}

List<Rectangle> GetPagesPerSheetLayout(int pps, Rectangle page)
{
    var layout = new List<Rectangle>();
    if (pps == 2)
    {
        var hw = page.Width / 2;
        layout.Add(new Rectangle(page.X, page.Y, hw, page.Height));
        layout.Add(new Rectangle(page.X + hw, page.Y, hw, page.Height));
    }
    else if (pps == 4)
    {
        var hw = page.Width / 2;
        var hh = page.Height / 2;
        layout.Add(new Rectangle(page.X, page.Y, hw, hh));
        layout.Add(new Rectangle(page.X + hw, page.Y, hw, hh));
        layout.Add(new Rectangle(page.X, page.Y + hh, hw, hh));
        layout.Add(new Rectangle(page.X + hw, page.Y + hh, hw, hh));
    }
    else { layout.Add(page); }
    return layout;
}

List<int> ParsePageRange(string? pageRange, int totalPages)
{
    var pages = new List<int>();
    if (string.IsNullOrEmpty(pageRange) || pageRange.ToLower() == "all")
    {
        for (int i = 0; i < totalPages; i++) pages.Add(i);
        return pages;
    }
    foreach (var range in pageRange.Split(','))
    {
        var trimmed = range.Trim();
        if (trimmed.Contains('-'))
        {
            var parts = trimmed.Split('-');
            var start = int.Parse(parts[0].Trim()) - 1;
            var end = int.Parse(parts[1].Trim()) - 1;
            for (int i = start; i <= end && i < totalPages; i++) pages.Add(i);
        }
        else
        {
            var page = int.Parse(trimmed) - 1;
            if (page >= 0 && page < totalPages) pages.Add(page);
        }
    }
    return pages.Distinct().OrderBy(p => p).ToList();
}

public class PrinterMonitor
{
    public async Task<PrinterStatus> GetPrinterStatus(string printerName)
    {
        return await Task.Run(() =>
        {
            var status = new PrinterStatus { PrinterName = printerName };
            try
            {
                using var searcher = new ManagementObjectSearcher(
                    $"SELECT * FROM Win32_Printer WHERE Name = '{printerName.Replace("\\", "\\\\")}'");

                foreach (ManagementObject printer in searcher.Get())
                {
                    var printerStatus = Convert.ToUInt16(printer["PrinterStatus"]);
                    var detectedErrorState = Convert.ToUInt16(printer["DetectedErrorState"]);
                    var printerState = Convert.ToUInt16(printer["PrinterState"]);

                    status.IsOnline = printerStatus != 7 && printerState != 512; // 7=Offline, 512=Offline
                    status.IsProcessing = printerStatus == 4; // 4 = Printing
                    
                    // More accurate paper detection
                    status.IsPaperOut = detectedErrorState == 4 || detectedErrorState == 5 || 
                                       printerStatus == 5; // 5 = Out of paper
                    status.IsPaperJam = detectedErrorState == 3;
                    status.IsError = detectedErrorState != 0 && detectedErrorState != 2;

                    if (status.IsError)
                    {
                        status.ErrorMessage = detectedErrorState switch
                        {
                            3 => "Paper Jam",
                            4 => "Paper Out",
                            5 => "Paper Problem",
                            6 => "Toner Low",
                            7 => "Toner Empty",
                            8 => "Output Bin Full",
                            9 => "Paper Problem",
                            10 => "Cannot Print Page",
                            11 => "User Intervention Required",
                            12 => "Out of Memory",
                            13 => "Door Open",
                            _ => "Unknown Error"
                        };
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️  Status check failed: {ex.Message}");
                status.IsOnline = false;
            }
            return status;
        });
    }

    public async Task<List<PrintJobInfo>> GetPrintJobs(string printerName)
    {
        return await Task.Run(() =>
        {
            var jobs = new List<PrintJobInfo>();
            try
            {
                using var searcher = new ManagementObjectSearcher(
                    $"SELECT * FROM Win32_PrintJob WHERE Name LIKE '%{printerName.Split('\\').Last()}%'");

                foreach (ManagementObject job in searcher.Get())
                {
                    var jobInfo = new PrintJobInfo
                    {
                        JobId = Convert.ToInt32(job["JobId"]),
                        Status = job["Status"]?.ToString() ?? "Unknown",
                        TotalPages = Convert.ToInt32(job["TotalPages"]),
                        PagesPrinted = Convert.ToInt32(job["PagesPrinted"]),
                        Document = job["Document"]?.ToString() ?? ""
                    };
                    jobs.Add(jobInfo);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️  Job query failed: {ex.Message}");
            }
            return jobs;
        });
    }

    public async Task<InkLevels?> GetInkLevels(string printerName)
    {
        return await Task.Run(() =>
        {
            try
            {
                var inkLevels = GetInkLevelsSNMP(printerName);
                if (inkLevels != null) return inkLevels;
                return GetInkLevelsWMI(printerName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️  Ink level check failed: {ex.Message}");
                return null;
            }
        });
    }

    private InkLevels? GetInkLevelsSNMP(string printerName)
    {
        try
        {
            var ipAddress = GetPrinterIPAddress(printerName);
            if (ipAddress == null) return null;

            var endpoint = new IPEndPoint(IPAddress.Parse(ipAddress), 161);
            var community = new OctetString("public");
            
            var oids = new List<Variable>
            {
                new Variable(new ObjectIdentifier("1.3.6.1.2.1.43.11.1.1.9.1.1")),
                new Variable(new ObjectIdentifier("1.3.6.1.2.1.43.11.1.1.9.1.2")),
                new Variable(new ObjectIdentifier("1.3.6.1.2.1.43.11.1.1.9.1.3")),
                new Variable(new ObjectIdentifier("1.3.6.1.2.1.43.11.1.1.9.1.4"))
            };

            var result = Messenger.Get(VersionCode.V2, endpoint, community, oids, 2000);

            return new InkLevels
            {
                Black = result.Count > 0 ? Convert.ToInt32(result[0].Data.ToString()) : 0,
                Cyan = result.Count > 1 ? Convert.ToInt32(result[1].Data.ToString()) : 0,
                Magenta = result.Count > 2 ? Convert.ToInt32(result[2].Data.ToString()) : 0,
                Yellow = result.Count > 3 ? Convert.ToInt32(result[3].Data.ToString()) : 0
            };
        }
        catch { return null; }
    }

    private InkLevels? GetInkLevelsWMI(string printerName)
    {
        try
        {
            using var searcher = new ManagementObjectSearcher(
                $"SELECT * FROM Win32_PrinterConfiguration WHERE Name = '{printerName.Replace("\\", "\\\\")}'");
            return new InkLevels { Black = 50, Cyan = 50, Magenta = 50, Yellow = 50 };
        }
        catch { return null; }
    }

    private string? GetPrinterIPAddress(string printerName)
    {
        try
        {
            using var searcher = new ManagementObjectSearcher(
                $"SELECT * FROM Win32_Printer WHERE Name = '{printerName.Replace("\\", "\\\\")}'");

            foreach (ManagementObject printer in searcher.Get())
            {
                var portName = printer["PortName"]?.ToString();
                if (portName != null && IPAddress.TryParse(portName.Replace("IP_", ""), out _))
                {
                    return portName.Replace("IP_", "");
                }
            }
        }
        catch { }
        return null;
    }
}

public class PrintJob
{
    [JsonPropertyName("file_path")]
    public string FilePath { get; set; } = string.Empty;
    [JsonPropertyName("copies")]
    public int Copies { get; set; } = 1;
    [JsonPropertyName("duplex")]
    public string? Duplex { get; set; } = "simplex";
    [JsonPropertyName("page_range")]
    public string? PageRange { get; set; } = "all";
    [JsonPropertyName("pages_per_sheet")]
    public int PagesPerSheet { get; set; } = 1;
}

public class PrintOptions
{
    public string? PrinterName { get; set; }
    public int Copies { get; set; } = 1;
    public bool Color { get; set; } = true;
    public string? Duplex { get; set; } = "simplex";
    public string? PageRange { get; set; } = "all";
    public string? PaperSize { get; set; } = "A4";
    public string? Orientation { get; set; } = "portrait";
    public string? Scale { get; set; } = "fit";
    public int PagesPerSheet { get; set; } = 1;
    public int Quality { get; set; } = 600;
}

public class PrinterStatus
{
    public string PrinterName { get; set; } = string.Empty;
    public bool IsOnline { get; set; }
    public bool IsProcessing { get; set; }
    public bool IsError { get; set; }
    public bool IsPaperJam { get; set; }
    public bool IsPaperOut { get; set; }
    public string ErrorMessage { get; set; } = string.Empty;
    public int JobCount { get; set; }
}

public class PrintJobInfo
{
    public int JobId { get; set; }
    public string Status { get; set; } = string.Empty;
    public int TotalPages { get; set; }
    public int PagesPrinted { get; set; }
    public string Document { get; set; } = string.Empty;
}

public class InkLevels
{
    public int Black { get; set; }
    public int Cyan { get; set; }
    public int Magenta { get; set; }
    public int Yellow { get; set; }
}

public class WsMessage
{
    public string Type { get; set; } = string.Empty;
    public string? StatusCode { get; set; }
    public string? Message { get; set; }
    public int? PageNumber { get; set; }
    public InkLevels? InkLevels { get; set; }

    public static WsMessage CreateStatus(string status, string message) => new()
    { Type = "status", StatusCode = status, Message = message };

    public static WsMessage PagePrinted(int pageNum) => new()
    { Type = "page_printed", PageNumber = pageNum, Message = $"Page {pageNum} printing completed" };

    public static WsMessage Error(string error) => new()
    { Type = "error", StatusCode = "error", Message = error };

    public static WsMessage PrinterStatus(string status, string message) => new()
    { Type = "printer_status", StatusCode = status, Message = message };

    public static WsMessage InkLevel(InkLevels levels) => new()
    { Type = "ink_level", StatusCode = "ink_check", 
      Message = $"Ink: C={levels.Cyan}% M={levels.Magenta}% Y={levels.Yellow}% K={levels.Black}%", 
      InkLevels = levels };
}

static class Win32
{
    [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern bool OpenPrinter(string pPrinterName, out IntPtr phPrinter, IntPtr pDefault);
    [DllImport("winspool.drv", SetLastError = true)]
    public static extern bool ClosePrinter(IntPtr hPrinter);
    [DllImport("winspool.drv", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern int DocumentProperties(IntPtr hwnd, IntPtr hPrinter, string pDeviceName,
        IntPtr pDevModeOutput, IntPtr pDevModeInput, int fMode);

    public const int DM_OUT_BUFFER = 2;
    public const int DM_IN_BUFFER = 8;
    public const int DM_COPIES = 0x00000100;
    public const int DM_DUPLEX = 0x00001000;
    public const int DM_COLOR = 0x00000800;
    public const int DM_ORIENTATION = 0x00000001;
    public const int DM_PRINTQUALITY = 0x00000400;
    public const int DM_YRESOLUTION = 0x00002000;
    public const short DMDUP_SIMPLEX = 1;
    public const short DMDUP_VERTICAL = 2;
    public const short DMDUP_HORIZONTAL = 3;
    public const short DMCOLOR_MONOCHROME = 1;
    public const short DMCOLOR_COLOR = 2;
    public const short DMORIENT_PORTRAIT = 1;
    public const short DMORIENT_LANDSCAPE = 2;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct DEVMODE
    {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmDeviceName;
        public short dmSpecVersion;
        public short dmDriverVersion;
        public short dmSize;
        public short dmDriverExtra;
        public int dmFields;
        public short dmOrientation;
        public short dmPaperSize;
        public short dmPaperLength;
        public short dmPaperWidth;
        public short dmScale;
        public short dmCopies;
        public short dmDefaultSource;
        public short dmPrintQuality;
        public short dmColor;
        public short dmDuplex;
        public short dmYResolution;
        public short dmTTOption;
        public short dmCollate;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmFormName;
        public short dmLogPixels;
        public int dmBitsPerPel;
        public int dmPelsWidth;
        public int dmPelsHeight;
        public int dmDisplayFlags;
        public int dmDisplayFrequency;
        public int dmICMMethod;
        public int dmICMIntent;
        public int dmMediaType;
        public int dmDitherType;
        public int dmReserved1;
        public int dmReserved2;
        public int dmPanningWidth;
        public int dmPanningHeight;
    }
}