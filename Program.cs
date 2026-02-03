using System.Drawing;
using System.Drawing.Printing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Runtime.InteropServices;
using System.Net.WebSockets;
using System.Text;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Docnet.Core;
using Docnet.Core.Models;

// ============== APP SETUP ==============
var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

app.UseWebSockets(new WebSocketOptions { KeepAliveInterval = TimeSpan.FromSeconds(30) });

var _semaphore = new SemaphoreSlim(1, 1); // Ensure one print job at a time

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
app.Run("http://0.0.0.0:8765");

// ============== WEBSOCKET HANDLER ==============
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

// ============== MESSAGE PROCESSOR ==============
async Task ProcessMessage(WebSocket ws, string message)
{
    // Acquire semaphore to prevent concurrent printing
    await _semaphore.WaitAsync();
    
    try
    {
        var job = JsonSerializer.Deserialize<PrintJob>(message, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        });

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
        
        await SendWsMessage(ws, WsMessage.CreateStatus("settings_applied", "Settings applied successfully"));
        Console.WriteLine($"✅ Settings prepared");

        // Execute print with proper cleanup
        await ExecutePrintSafe(ws, job.FilePath, options, ext);

        await SendWsMessage(ws, WsMessage.CreateStatus("job_completed", "Job completed successfully"));
        Console.WriteLine("✅ Job completed\n");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"❌ Error: {ex.Message}");
        Console.WriteLine($"Stack: {ex.StackTrace}");
        await SendWsMessage(ws, WsMessage.Error($"Print failed: {ex.Message}"));
    }
    finally
    {
        // CRITICAL: Always release the semaphore
        _semaphore.Release();
        
        // Small delay to let spooler fully process
        await Task.Delay(500);
        
        // Force garbage collection to release COM objects
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }
}

// ============== BUILD OPTIONS ==============
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

// ============== SAFE PRINT EXECUTION ==============
async Task ExecutePrintSafe(WebSocket ws, string filePath, PrintOptions options, string ext)
{
    switch (ext)
    {
        case ".pdf":
            await PrintPdfSafe(ws, filePath, options);
            break;
        case ".txt":
            await PrintTextSafe(ws, filePath, options);
            break;
        default:
            await PrintImageSafe(ws, filePath, options);
            break;
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
        // CRITICAL: Dispose in correct order
        if (bitmaps != null)
        {
            foreach (var bmp in bitmaps)
            {
                try { bmp?.Dispose(); } catch { }
            }
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
    var completedPages = new List<int>();
    PrintDocument? printDoc = null;

    try
    {
        printDoc = new PrintDocument();
        
        // Apply settings BEFORE attaching events
        ApplyPrinterSettingsSafe(printDoc, options);

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
            completedPages.Add(pageNumber);
            e.HasMorePages = lineIndex < lines.Length;
            
            font.Dispose();
        };

        printDoc.Print();
        
        // Wait for spooler to accept job
        await Task.Delay(200);

        foreach (var p in completedPages)
        {
            await SendWsMessage(ws, WsMessage.PagePrinted(p));
            Console.WriteLine($"  ✓ Page {p}");
        }
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

        await SendWsMessage(ws, WsMessage.PagePrinted(1));
        Console.WriteLine("  ✓ Page 1");
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
        
        // Apply settings FIRST
        ApplyPrinterSettingsSafe(printDoc, options);

        var pageIndex = 0;
        var pps = options.PagesPerSheet;
        var completedPages = new List<int>();

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
                    completedPages.Add(pageNumbers[pageIndex] + 1);
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
                    completedPages.Add(pageNumbers[pageIndex + i] + 1);
                }

                pageIndex += count;
                e.HasMorePages = pageIndex < bitmaps.Count;
            }
        };

        printDoc.Print();
        
        // Wait for spooler
        await Task.Delay(200);

        foreach (var p in completedPages)
        {
            await SendWsMessage(ws, WsMessage.PagePrinted(p));
            Console.WriteLine($"  ✓ Page {p}");
        }
    }
    finally
    {
        try { printDoc?.Dispose(); } catch { }
    }
}

// ============== PRINTER SETTINGS (SAFE) ==============
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
            Console.WriteLine("⚠️  DEVMODE size failed, using fallback");
            ApplyBasicSettings(printDoc, options);
            return;
        }

        pDevMode = Marshal.AllocHGlobal(sizeNeeded);

        int result = Win32.DocumentProperties(IntPtr.Zero, hPrinter, options.PrinterName!, pDevMode, IntPtr.Zero, Win32.DM_OUT_BUFFER);
        if (result < 0)
        {
            Console.WriteLine("⚠️  DM_OUT_BUFFER failed, using fallback");
            ApplyBasicSettings(printDoc, options);
            return;
        }

        var devMode = Marshal.PtrToStructure<Win32.DEVMODE>(pDevMode);

        devMode.dmCopies = (short)options.Copies;
        devMode.dmFields |= Win32.DM_COPIES;

        devMode.dmDuplex = options.Duplex?.ToLower() switch
        {
            "vertical" or "double" => Win32.DMDUP_VERTICAL,
            "horizontal" => Win32.DMDUP_HORIZONTAL,
            _ => Win32.DMDUP_SIMPLEX
        };
        devMode.dmFields |= Win32.DM_DUPLEX;

        devMode.dmColor = options.Color ? Win32.DMCOLOR_COLOR : Win32.DMCOLOR_MONOCHROME;
        devMode.dmFields |= Win32.DM_COLOR;

        devMode.dmOrientation = options.Orientation?.ToLower() == "landscape"
            ? Win32.DMORIENT_LANDSCAPE : Win32.DMORIENT_PORTRAIT;
        devMode.dmFields |= Win32.DM_ORIENTATION;

        devMode.dmPrintQuality = (short)options.Quality;
        devMode.dmYResolution = (short)options.Quality;
        devMode.dmFields |= Win32.DM_PRINTQUALITY;
        devMode.dmFields |= Win32.DM_YRESOLUTION;

        Marshal.StructureToPtr(devMode, pDevMode, true);

        result = Win32.DocumentProperties(IntPtr.Zero, hPrinter, options.PrinterName!,
            pDevMode, pDevMode, Win32.DM_IN_BUFFER | Win32.DM_OUT_BUFFER);

        if (result >= 0)
        {
            printDoc.PrinterSettings.SetHdevmode(pDevMode);
            printDoc.DefaultPageSettings.SetHdevmode(pDevMode);
            Console.WriteLine($"  ✓ Copies: {options.Copies}, Duplex: {options.Duplex}, Quality: {options.Quality} DPI");
        }
        else
        {
            Console.WriteLine($"⚠️  DM_IN_BUFFER returned {result}, using fallback");
            ApplyBasicSettings(printDoc, options);
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"⚠️  Win32 error: {ex.Message}, using fallback");
        ApplyBasicSettings(printDoc, options);
    }
    finally
    {
        // CRITICAL: Always free resources
        if (pDevMode != IntPtr.Zero)
        {
            try { Marshal.FreeHGlobal(pDevMode); } catch { }
        }
        if (hPrinter != IntPtr.Zero)
        {
            try { Win32.ClosePrinter(hPrinter); } catch { }
        }
    }
}

void ApplyBasicSettings(PrintDocument printDoc, PrintOptions options)
{
    printDoc.PrinterSettings.Copies = (short)options.Copies;
    printDoc.DefaultPageSettings.Color = options.Color;
    printDoc.PrinterSettings.Duplex = options.Duplex?.ToLower() switch
    {
        "vertical" or "double" => Duplex.Vertical,
        "horizontal" => Duplex.Horizontal,
        _ => Duplex.Simplex
    };
    printDoc.DefaultPageSettings.Landscape = options.Orientation?.ToLower() == "landscape";
    Console.WriteLine("  ✓ Basic settings applied");
}

// ============== WEBSOCKET HELPERS ==============
async Task SendWsMessage(WebSocket ws, WsMessage msg)
{
    if (ws.State != WebSocketState.Open) return;

    var json = JsonSerializer.Serialize(msg, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
    var bytes = Encoding.UTF8.GetBytes(json);

    try
    {
        await ws.SendAsync(bytes, WebSocketMessageType.Text, true, CancellationToken.None);
    }
    catch { }
}

// ============== RENDERING HELPERS ==============
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

// ============== MODELS ==============
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

public class WsMessage
{
    public string Type { get; set; } = string.Empty;
    public string? StatusCode { get; set; }
    public string? Message { get; set; }
    public int? PageNumber { get; set; }

    public static WsMessage CreateStatus(string status, string message) => new()
    {
        Type = "status",
        StatusCode = status,
        Message = message
    };

    public static WsMessage PagePrinted(int pageNum) => new()
    {
        Type = "page_printed",
        PageNumber = pageNum,
        Message = $"Page {pageNum} printing completed"
    };

    public static WsMessage Error(string error) => new()
    {
        Type = "error",
        StatusCode = "error",
        Message = error
    };
}

// ============== WIN32 API ==============
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