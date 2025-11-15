using System.Drawing;
using System.Drawing.Printing;
using System.Drawing.Imaging;
using System.Text.Json;
using System.Runtime.InteropServices;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Docnet.Core;
using Docnet.Core.Models;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddCors();
var app = builder.Build();

app.UseCors(policy => policy.AllowAnyOrigin().AllowAnyMethod().AllowAnyHeader());

app.MapGet("/health", () => Results.Ok(new { status = "healthy", timestamp = DateTime.UtcNow }));

app.MapGet("/printers", () =>
{
    var printers = PrinterSettings.InstalledPrinters.Cast<string>().ToList();
    return Results.Ok(new { printers });
});

app.MapPost("/print", async (HttpContext context) =>
{
    try
    {
        var form = await context.Request.ReadFormAsync();
        
        var file = form.Files["file"];
        if (file == null || file.Length == 0)
            return Results.BadRequest(new { error = "No file provided" });

        var optionsJson = form["options"].ToString();
        Console.WriteLine($"Received options JSON: {optionsJson}");
        
        if (string.IsNullOrEmpty(optionsJson))
            return Results.BadRequest(new { error = "No options provided" });

        var options = JsonSerializer.Deserialize<PrintOptions>(optionsJson, new JsonSerializerOptions 
        { 
            PropertyNameCaseInsensitive = true 
        });
        
        if (options == null)
            return Results.BadRequest(new { error = "Invalid print options" });
        
        Console.WriteLine($"Parsed options - Printer: {options.PrinterName}, Copies: {options.Copies}, Duplex: {options.Duplex}, Color: {options.Color}");

        var tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + Path.GetExtension(file.FileName));
        using (var stream = new FileStream(tempPath, FileMode.Create))
        {
            await file.CopyToAsync(stream);
        }

        Console.WriteLine($"File saved: {tempPath} ({new FileInfo(tempPath).Length} bytes)");

        var jobId = await PrintFileAsync(tempPath, options, file.FileName);

        _ = Task.Run(async () =>
        {
            await Task.Delay(30000);
            try { File.Delete(tempPath); } catch { }
        });

        return Results.Ok(new { success = true, jobId, message = "Print job submitted successfully" });
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error: {ex.Message}\n{ex.StackTrace}");
        return Results.Problem(detail: ex.Message, statusCode: 500);
    }
});

app.Run("http://localhost:8765");

static async Task<string> PrintFileAsync(string filePath, PrintOptions options, string fileName)
{
    return await Task.Run(() =>
    {
        var extension = Path.GetExtension(filePath).ToLower();
        var jobId = Guid.NewGuid().ToString();

        Console.WriteLine($"\n=== PRINT JOB START ===");
        Console.WriteLine($"File: {fileName} ({extension})");
        Console.WriteLine($"Printer: {options.PrinterName ?? "Default"}");
        Console.WriteLine($"Copies: {options.Copies}");
        Console.WriteLine($"Duplex: {options.Duplex}");
        Console.WriteLine($"Color: {options.Color}");
        Console.WriteLine($"Pages: {options.PageRange}");

        switch (extension)
        {
            case ".pdf":
                PrintPdf(filePath, options);
                break;
            case ".txt":
                PrintTextFile(filePath, options);
                break;
            case ".jpg":
            case ".jpeg":
            case ".png":
            case ".bmp":
            case ".gif":
            case ".tif":
            case ".tiff":
                PrintImage(filePath, options);
                break;
            default:
                throw new NotSupportedException($"Unsupported file type: {extension}");
        }

        Console.WriteLine($"=== PRINT JOB COMPLETE ===\n");
        return jobId;
    });
}

static void PrintPdf(string pdfPath, PrintOptions options)
{
    var bitmaps = new List<Bitmap>();
    
    try
    {
        using var library = DocLib.Instance;
        using var docReader = library.GetDocReader(pdfPath, new PageDimensions(2400, 2400));
        
        var pageCount = docReader.GetPageCount();
        var pagesToPrint = GetPagesToPrint(options.PageRange ?? "all", pageCount);
        
        Console.WriteLine($"Rendering {pagesToPrint.Count} of {pageCount} pages...");
        
        foreach (var pageNumber in pagesToPrint)
        {
            using var pageReader = docReader.GetPageReader(pageNumber);
            var rawBytes = pageReader.GetImage();
            var width = pageReader.GetPageWidth();
            var height = pageReader.GetPageHeight();
            
            var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            var bitmapData = bitmap.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.WriteOnly, bitmap.PixelFormat);
            
            try
            {
                Marshal.Copy(rawBytes, 0, bitmapData.Scan0, rawBytes.Length);
            }
            finally
            {
                bitmap.UnlockBits(bitmapData);
            }
            
            bitmaps.Add(bitmap);
        }
        
        // Print with enforced settings
        PrintBitmapsWithSettings(bitmaps, options, pagesToPrint);
    }
    finally
    {
        foreach (var bitmap in bitmaps)
        {
            bitmap.Dispose();
        }
    }
}

static void PrintTextFile(string textPath, PrintOptions options)
{
    var lines = File.ReadAllLines(textPath);
    var lineIndex = 0;

    using var printDoc = new PrintDocument();
    
    // Apply settings using Win32 API
    ApplyPrinterSettings(printDoc, options);

    printDoc.PrintPage += (sender, e) =>
    {
        if (e?.Graphics == null) return;
        
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

        e.HasMorePages = lineIndex < lines.Length;
    };

    printDoc.Print();
}

static void PrintImage(string imagePath, PrintOptions options)
{
    using var image = Image.FromFile(imagePath);
    using var printDoc = new PrintDocument();
    
    ApplyPrinterSettings(printDoc, options);

    printDoc.PrintPage += (sender, e) =>
    {
        if (e?.Graphics == null) return;
        var imgRect = GetScaledImageRectangle(image, e.PageBounds);
        e.Graphics.DrawImage(image, imgRect);
        e.HasMorePages = false;
    };

    printDoc.Print();
}

static void PrintBitmapsWithSettings(List<Bitmap> bitmaps, PrintOptions options, List<int> pageNumbers)
{
    using var printDoc = new PrintDocument();
    
    // CRITICAL: Apply settings using Win32 API BEFORE printing
    ApplyPrinterSettings(printDoc, options);
    
    var pageIndex = 0;
    
    printDoc.PrintPage += (sender, e) =>
    {
        if (e?.Graphics == null) return;
        
        if (pageIndex < bitmaps.Count)
        {
            var bitmap = bitmaps[pageIndex];
            var imgRect = GetScaledImageRectangle(bitmap, e.PageBounds);
            e.Graphics.DrawImage(bitmap, imgRect);
            
            Console.WriteLine($"Printed page {pageNumbers[pageIndex] + 1}");
            pageIndex++;
            e.HasMorePages = pageIndex < bitmaps.Count;
        }
        else
        {
            e.HasMorePages = false;
        }
    };
    
    printDoc.Print();
}

static void ApplyPrinterSettings(PrintDocument printDoc, PrintOptions options)
{
    var printerName = string.IsNullOrEmpty(options.PrinterName) 
        ? new PrinterSettings().PrinterName 
        : options.PrinterName;
    
    printDoc.PrinterSettings.PrinterName = printerName;
    
    if (!printDoc.PrinterSettings.IsValid)
    {
        throw new Exception($"Printer not found: {printerName}");
    }

    Console.WriteLine($"\n--- Applying Settings to: {printerName} ---");

    // Get printer handle
    IntPtr hPrinter = IntPtr.Zero;
    if (!Win32.OpenPrinter(printerName, out hPrinter, IntPtr.Zero))
    {
        Console.WriteLine("Warning: Could not open printer for advanced settings");
        // Fallback to basic settings
        ApplyBasicSettings(printDoc, options);
        return;
    }

    try
    {
        // Get current DEVMODE
        int sizeNeeded = Win32.DocumentProperties(IntPtr.Zero, hPrinter, printerName, IntPtr.Zero, IntPtr.Zero, 0);
        IntPtr pDevMode = Marshal.AllocHGlobal(sizeNeeded);

        try
        {
            Win32.DocumentProperties(IntPtr.Zero, hPrinter, printerName, pDevMode, IntPtr.Zero, Win32.DM_OUT_BUFFER);
            
            var devMode = Marshal.PtrToStructure<Win32.DEVMODE>(pDevMode);

            // Set copies
            devMode.dmCopies = (short)options.Copies;
            devMode.dmFields |= Win32.DM_COPIES;
            Console.WriteLine($"✓ Copies: {devMode.dmCopies}");

            // Set duplex
            var duplexMode = (options.Duplex ?? "simplex").ToLower();
            devMode.dmDuplex = duplexMode switch
            {
                "vertical" => Win32.DMDUP_VERTICAL,
                "horizontal" => Win32.DMDUP_HORIZONTAL,
                _ => Win32.DMDUP_SIMPLEX
            };
            devMode.dmFields |= Win32.DM_DUPLEX;
            Console.WriteLine($"✓ Duplex: {duplexMode} ({devMode.dmDuplex})");

            // Set color
            devMode.dmColor = options.Color ? Win32.DMCOLOR_COLOR : Win32.DMCOLOR_MONOCHROME;
            devMode.dmFields |= Win32.DM_COLOR;
            Console.WriteLine($"✓ Color: {(options.Color ? "Color" : "Monochrome")}");

            // Set orientation
            var orientation = (options.Orientation ?? "portrait").ToLower();
            devMode.dmOrientation = orientation == "landscape" ? Win32.DMORIENT_LANDSCAPE : Win32.DMORIENT_PORTRAIT;
            devMode.dmFields |= Win32.DM_ORIENTATION;
            Console.WriteLine($"✓ Orientation: {orientation}");

            // Write back DEVMODE
            Marshal.StructureToPtr(devMode, pDevMode, true);
            
            // Apply settings
            int result = Win32.DocumentProperties(IntPtr.Zero, hPrinter, printerName, pDevMode, pDevMode, Win32.DM_IN_BUFFER | Win32.DM_OUT_BUFFER);
            
            if (result >= 0)
            {
                // Set to PrintDocument
                printDoc.PrinterSettings.SetHdevmode(pDevMode);
                printDoc.DefaultPageSettings.SetHdevmode(pDevMode);
                Console.WriteLine("✓ Settings applied successfully via Win32 API");
            }
            else
            {
                Console.WriteLine($"Warning: DocumentProperties returned {result}");
            }
        }
        finally
        {
            Marshal.FreeHGlobal(pDevMode);
        }
    }
    finally
    {
        Win32.ClosePrinter(hPrinter);
    }
    
    Console.WriteLine("--- Settings Applied ---\n");
}

static void ApplyBasicSettings(PrintDocument printDoc, PrintOptions options)
{
    printDoc.PrinterSettings.Copies = (short)options.Copies;
    printDoc.DefaultPageSettings.Color = options.Color;
    
    var duplexMode = (options.Duplex ?? "simplex").ToLower();
    printDoc.PrinterSettings.Duplex = duplexMode switch
    {
        "vertical" => Duplex.Vertical,
        "horizontal" => Duplex.Horizontal,
        _ => Duplex.Simplex
    };
    
    printDoc.DefaultPageSettings.Landscape = (options.Orientation ?? "portrait").ToLower() == "landscape";
}

static Rectangle GetScaledImageRectangle(Image image, Rectangle pageRect)
{
    var scale = Math.Min((float)pageRect.Width / image.Width, (float)pageRect.Height / image.Height);
    var newWidth = (int)(image.Width * scale);
    var newHeight = (int)(image.Height * scale);
    var x = pageRect.X + (pageRect.Width - newWidth) / 2;
    var y = pageRect.Y + (pageRect.Height - newHeight) / 2;
    return new Rectangle(x, y, newWidth, newHeight);
}

static List<int> GetPagesToPrint(string pageRange, int totalPages)
{
    var pages = new List<int>();
    if (pageRange.ToLower() == "all")
    {
        for (int i = 0; i < totalPages; i++) pages.Add(i);
        return pages;
    }

    foreach (var range in pageRange.Split(','))
    {
        if (range.Contains('-'))
        {
            var parts = range.Split('-');
            var start = int.Parse(parts[0].Trim()) - 1;
            var end = int.Parse(parts[1].Trim()) - 1;
            for (int i = start; i <= end && i < totalPages; i++) pages.Add(i);
        }
        else
        {
            var page = int.Parse(range.Trim()) - 1;
            if (page < totalPages) pages.Add(page);
        }
    }
    return pages.Distinct().OrderBy(p => p).ToList();
}

class PrintOptions
{
    public string? PrinterName { get; set; }
    public int Copies { get; set; } = 1;
    public bool Color { get; set; } = false;
    public string? Duplex { get; set; } = "simplex";
    public string? PageRange { get; set; } = "all";
    public string? PaperSize { get; set; } = "A4";
    public string? Orientation { get; set; } = "portrait";
}

// Win32 API Declarations
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