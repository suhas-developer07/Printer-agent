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

// MUST be before any middleware that uses WebSockets
app.UseWebSockets(new WebSocketOptions { KeepAliveInterval = TimeSpan.FromSeconds(30) });

// Track active websocket
WebSocket? _activeSocket = null;
var _socketLock = new object();

// ============== WEBSOCKET ENDPOINT ==============
app.Use(async (context, next) =>
{
    if (context.Request.Path == "/ws" && context.WebSockets.IsWebSocketRequest)
    {
        var ws = await context.WebSockets.AcceptWebSocketAsync();
        lock (_socketLock) { _activeSocket = ws; }
        Console.WriteLine("🔗 WebSocket connected");

        await HandleWebSocket(ws);

        lock (_socketLock)
        {
            if (_activeSocket == ws) _activeSocket = null;
        }
        Console.WriteLine("🔗 WebSocket disconnected");
    }
    else
    {
        await next();
    }
});

// Health check (HTTP fallback)
app.MapGet("/health", () => Results.Ok(new { status = "healthy", timestamp = DateTime.UtcNow }));

app.Run("http://0.0.0.0:8765");

// ============== WEBSOCKET HANDLER ==============
async Task HandleWebSocket(WebSocket ws)
{
    var buffer = new byte[1024 * 64]; // 64KB buffer

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
                Console.WriteLine($"📨 Received: {message}");

                // Fire and forget so we can keep listening
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
    PrintJob? job = null;
    try
    {
        job = JsonSerializer.Deserialize<PrintJob>(message, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        });

        if (job == null || string.IsNullOrWhiteSpace(job.FilePath))
        {
            await SendWsMessage(ws, WsMessage.Error("Invalid payload: missing or empty file_path"));
            return;
        }

        // Validate file exists
        if (!File.Exists(job.FilePath))
        {
            await SendWsMessage(ws, WsMessage.Error($"File not found: {job.FilePath}"));
            return;
        }

        // Validate file type
        var ext = Path.GetExtension(job.FilePath).ToLower();
        if (!new[] { ".pdf", ".txt", ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff" }.Contains(ext))
        {
            await SendWsMessage(ws, WsMessage.Error($"Unsupported file type: {ext}"));
            return;
        }

        // ACK
        await SendWsMessage(ws, WsMessage.CreateStatus("data_received", "Data received successfully"));
        Console.WriteLine("✅ Data received and validated");

        // Build full print options with defaults
        var options = BuildPrintOptions(job);

        // Apply settings
        using var printDoc = new PrintDocument();
        ApplyPrinterSettings(printDoc, options);
        await SendWsMessage(ws, WsMessage.CreateStatus("settings_applied", "Settings applied successfully"));
        Console.WriteLine("✅ Settings applied");

        // Execute print
        await ExecutePrint(ws, job.FilePath, options, ext);

        // Success
        await SendWsMessage(ws, WsMessage.CreateStatus("job_completed", "Job completed successfully"));
        Console.WriteLine("✅ Job completed\n");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"❌ Error: {ex.Message}");
        await SendWsMessage(ws, WsMessage.Error(ex.Message));
    }
}

// ============== BUILD OPTIONS WITH DEFAULTS ==============
PrintOptions BuildPrintOptions(PrintJob job)
{
    return new PrintOptions
    {
        PrinterName = "EPSON WF-C5890 Series",
        Copies = job.Copies > 0 ? job.Copies : 1,
        Duplex = job.Duplex ?? "simplex",
        PageRange = job.PageRange ?? "all",
        PagesPerSheet = job.PagesPerSheet > 0 ? job.PagesPerSheet : 1,
        // Defaults applied automatically
        Color = true,
        Orientation = "portrait",
        PaperSize = "A4",
        Scale = "fit",
        Quality = 600
    };
}

// ============== PRINT EXECUTION ==============
async Task ExecutePrint(WebSocket ws, string filePath, PrintOptions options, string ext)
{
    switch (ext)
    {
        case ".pdf":
            await PrintPdf(ws, filePath, options);
            break;
        case ".txt":
            await PrintText(ws, filePath, options);
            break;
        default: // images
            await PrintImage(ws, filePath, options);
            break;
    }
}

async Task PrintPdf(WebSocket ws, string pdfPath, PrintOptions options)
{
    var bitmaps = new List<Bitmap>();
    var pagesToPrint = new List<int>();

    try
    {
        using var library = DocLib.Instance;
        var dpi = options.Quality;

        // Render at correct DPI based on A4 dimensions (8.27 x 11.69 inches)
        var renderW = (int)(8.27 * dpi);
        var renderH = (int)(11.69 * dpi);

        using var docReader = library.GetDocReader(pdfPath, new PageDimensions(renderW, renderH));
        var pageCount = docReader.GetPageCount();
        pagesToPrint = ParsePageRange(options.PageRange, pageCount);

        Console.WriteLine($"📖 PDF: {pageCount} total, printing {pagesToPrint.Count} pages at {dpi} DPI");

        foreach (var pageNum in pagesToPrint)
        {
            using var pageReader = docReader.GetPageReader(pageNum);
            var rawBytes = pageReader.GetImage();
            var w = pageReader.GetPageWidth();
            var h = pageReader.GetPageHeight();

            var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb);
            var data = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.WriteOnly, bmp.PixelFormat);
            try { Marshal.Copy(rawBytes, 0, data.Scan0, rawBytes.Length); }
            finally { bmp.UnlockBits(data); }

            bitmaps.Add(bmp);
        }

        // Print bitmaps with page-by-page WS notifications
        await PrintBitmaps(ws, bitmaps, options, pagesToPrint);
    }
    finally
    {
        foreach (var bmp in bitmaps) bmp.Dispose();
    }
}

async Task PrintText(WebSocket ws, string textPath, PrintOptions options)
{
    var lines = File.ReadAllLines(textPath);
    var lineIndex = 0;
    var pageNumber = 0;

    using var printDoc = new PrintDocument();
    ApplyPrinterSettings(printDoc, options);

    // We need sync event handler but async WS sends
    // Use a list to collect page completions, then send after
    var completedPages = new List<int>();

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
    };

    printDoc.Print();

    // Send page notifications after print completes
    foreach (var p in completedPages)
    {
        await SendWsMessage(ws, WsMessage.PagePrinted(p));
        Console.WriteLine($"  ✓ Page {p} printed");
    }
}

async Task PrintImage(WebSocket ws, string imagePath, PrintOptions options)
{
    using var image = Image.FromFile(imagePath);
    using var printDoc = new PrintDocument();

    ApplyPrinterSettings(printDoc, options);

    printDoc.PrintPage += (sender, e) =>
    {
        if (e?.Graphics == null) return;
        SetHighQuality(e.Graphics);
        var rect = ScaleImage(image, e.PageBounds, options.Scale);
        e.Graphics.DrawImage(image, rect);
        e.HasMorePages = false;
    };

    printDoc.Print();
    await SendWsMessage(ws, WsMessage.PagePrinted(1));
    Console.WriteLine("  ✓ Page 1 printed");
}

async Task PrintBitmaps(WebSocket ws, List<Bitmap> bitmaps, PrintOptions options, List<int> pageNumbers)
{
    using var printDoc = new PrintDocument();
    ApplyPrinterSettings(printDoc, options);

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
            // Multi page per sheet
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

    // Send per-page notifications after print
    foreach (var p in completedPages)
    {
        await SendWsMessage(ws, WsMessage.PagePrinted(p));
        Console.WriteLine($"  ✓ Page {p} printed");
    }
}

// ============== WIN32 PRINTER SETTINGS ==============
void ApplyPrinterSettings(PrintDocument printDoc, PrintOptions options)
{
    printDoc.PrinterSettings.PrinterName = options.PrinterName!;

    if (!printDoc.PrinterSettings.IsValid)
        throw new Exception($"Printer not found: {options.PrinterName}");

    Console.WriteLine($"\n--- Applying Settings to: {options.PrinterName} ---");

    IntPtr hPrinter = IntPtr.Zero;
    if (!Win32.OpenPrinter(options.PrinterName!, out hPrinter, IntPtr.Zero))
    {
        Console.WriteLine("⚠️  Win32 OpenPrinter failed, using basic fallback");
        ApplyBasicSettings(printDoc, options);
        return;
    }

    try
    {
        int sizeNeeded = Win32.DocumentProperties(IntPtr.Zero, hPrinter, options.PrinterName!, IntPtr.Zero, IntPtr.Zero, 0);
        IntPtr pDevMode = Marshal.AllocHGlobal(sizeNeeded);

        try
        {
            Win32.DocumentProperties(IntPtr.Zero, hPrinter, options.PrinterName!, pDevMode, IntPtr.Zero, Win32.DM_OUT_BUFFER);
            var devMode = Marshal.PtrToStructure<Win32.DEVMODE>(pDevMode);

            // Copies
            devMode.dmCopies = (short)options.Copies;
            devMode.dmFields |= Win32.DM_COPIES;

            // Duplex
            devMode.dmDuplex = options.Duplex?.ToLower() switch
            {
                "vertical" or "double" => Win32.DMDUP_VERTICAL,
                "horizontal" => Win32.DMDUP_HORIZONTAL,
                _ => Win32.DMDUP_SIMPLEX
            };
            devMode.dmFields |= Win32.DM_DUPLEX;

            // Color
            devMode.dmColor = options.Color ? Win32.DMCOLOR_COLOR : Win32.DMCOLOR_MONOCHROME;
            devMode.dmFields |= Win32.DM_COLOR;

            // Orientation
            devMode.dmOrientation = options.Orientation?.ToLower() == "landscape"
                ? Win32.DMORIENT_LANDSCAPE : Win32.DMORIENT_PORTRAIT;
            devMode.dmFields |= Win32.DM_ORIENTATION;

            // DPI / Quality
            devMode.dmPrintQuality = (short)options.Quality;
            devMode.dmYResolution = (short)options.Quality;
            devMode.dmFields |= Win32.DM_PRINTQUALITY;
            devMode.dmFields |= Win32.DM_YRESOLUTION;

            Marshal.StructureToPtr(devMode, pDevMode, true);

            int result = Win32.DocumentProperties(IntPtr.Zero, hPrinter, options.PrinterName!,
                pDevMode, pDevMode, Win32.DM_IN_BUFFER | Win32.DM_OUT_BUFFER);

            if (result >= 0)
            {
                printDoc.PrinterSettings.SetHdevmode(pDevMode);
                printDoc.DefaultPageSettings.SetHdevmode(pDevMode);
                Console.WriteLine($"  ✓ Copies: {options.Copies}");
                Console.WriteLine($"  ✓ Duplex: {options.Duplex}");
                Console.WriteLine($"  ✓ Color: {options.Color}");
                Console.WriteLine($"  ✓ Quality: {options.Quality} DPI");
                Console.WriteLine($"  ✓ Pages/Sheet: {options.PagesPerSheet}");
                Console.WriteLine("  ✓ Win32 settings applied");
            }
            else
            {
                Console.WriteLine($"⚠️  DocumentProperties returned {result}, using basic fallback");
                ApplyBasicSettings(printDoc, options);
            }
        }
        finally { Marshal.FreeHGlobal(pDevMode); }
    }
    finally { Win32.ClosePrinter(hPrinter); }

    Console.WriteLine("--- Settings Applied ---\n");
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
    catch (WebSocketException ex)
    {
        Console.WriteLine($"❌ Failed to send WS message: {ex.Message}");
    }
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
    float scale;
    if (scaleMode?.ToLower() == "actual")
        scale = 1.0f;
    else
        scale = Math.Min((float)pageRect.Width / image.Width, (float)pageRect.Height / image.Height);

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