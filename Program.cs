using System.Drawing;
using System.Drawing.Printing;
using System.Drawing.Imaging;
using System.Text.Json;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Docnet.Core;
using Docnet.Core.Models;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddCors();
var app = builder.Build();

app.UseCors(policy => policy.AllowAnyOrigin().AllowAnyMethod().AllowAnyHeader());

// Health check endpoint
app.MapGet("/health", () => Results.Ok(new { status = "healthy", timestamp = DateTime.UtcNow }));

// Get available printers
app.MapGet("/printers", () =>
{
    var printers = PrinterSettings.InstalledPrinters.Cast<string>().ToList();
    return Results.Ok(new { printers });
});

// Print job endpoint
app.MapPost("/print", async (HttpContext context) =>
{
    try
    {
        var form = await context.Request.ReadFormAsync();
        
        var file = form.Files["file"];
        if (file == null || file.Length == 0)
            return Results.BadRequest(new { error = "No file provided" });

        var optionsJson = form["options"].ToString();
        if (string.IsNullOrEmpty(optionsJson))
            return Results.BadRequest(new { error = "No options provided" });

        var options = JsonSerializer.Deserialize<PrintOptions>(optionsJson);
        if (options == null)
            return Results.BadRequest(new { error = "Invalid print options" });

        // Save temp file
        var tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + Path.GetExtension(file.FileName));
        using (var stream = new FileStream(tempPath, FileMode.Create))
        {
            await file.CopyToAsync(stream);
        }

        Console.WriteLine($"File saved to: {tempPath}");
        Console.WriteLine($"File size: {new FileInfo(tempPath).Length} bytes");

        // Print the file
        var jobId = await PrintFileAsync(tempPath, options, file.FileName);

        // Cleanup temp file after a delay
        _ = Task.Run(async () =>
        {
            await Task.Delay(30000);
            try { File.Delete(tempPath); } catch { }
        });

        return Results.Ok(new { success = true, jobId, message = "Print job submitted successfully" });
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error: {ex.Message}");
        Console.WriteLine($"Stack: {ex.StackTrace}");
        return Results.Problem(detail: ex.Message, statusCode: 500);
    }
});

app.Run("http://localhost:8765");

// Helper methods
static async Task<string> PrintFileAsync(string filePath, PrintOptions options, string fileName)
{
    return await Task.Run(() =>
    {
        var extension = Path.GetExtension(filePath).ToLower();
        var jobId = Guid.NewGuid().ToString();

        Console.WriteLine($"Processing file: {fileName} (Type: {extension})");
        Console.WriteLine($"Printer: {options.PrinterName ?? "Default"}");

        try
        {
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
                    throw new NotSupportedException($"File type {extension} is not supported. Supported types: PDF, TXT, JPG, PNG, BMP, GIF, TIF");
            }

            Console.WriteLine($"Print job submitted successfully. Job ID: {jobId}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Print error: {ex.Message}");
            throw;
        }

        return jobId;
    });
}

static void PrintPdf(string pdfPath, PrintOptions options)
{
    Console.WriteLine("Loading PDF document...");
    
    var bitmaps = new List<Bitmap>();
    
    try
    {
        using var library = DocLib.Instance;
        using var docReader = library.GetDocReader(pdfPath, new PageDimensions(1920, 1920));
        
        var pageCount = docReader.GetPageCount();
        Console.WriteLine($"PDF has {pageCount} pages");
        
        var pagesToPrint = GetPagesToPrint(options.PageRange ?? "all", pageCount);
        Console.WriteLine($"Will print {pagesToPrint.Count} pages: {string.Join(", ", pagesToPrint.Select(p => p + 1))}");
        
        // Render each page to bitmap
        foreach (var pageNumber in pagesToPrint)
        {
            Console.WriteLine($"Rendering page {pageNumber + 1}...");
            using var pageReader = docReader.GetPageReader(pageNumber);
            var rawBytes = pageReader.GetImage();
            var width = pageReader.GetPageWidth();
            var height = pageReader.GetPageHeight();
            
            var bitmap = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            var bitmapData = bitmap.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.WriteOnly, bitmap.PixelFormat);
            
            try
            {
                System.Runtime.InteropServices.Marshal.Copy(rawBytes, 0, bitmapData.Scan0, rawBytes.Length);
            }
            finally
            {
                bitmap.UnlockBits(bitmapData);
            }
            
            bitmaps.Add(bitmap);
        }
        
        // Now print all bitmaps
        using var printDoc = new PrintDocument();
        ConfigurePrintSettings(printDoc, options);
        
        var pageIndex = 0;
        
        printDoc.PrintPage += (sender, e) =>
        {
            if (e == null || e.Graphics == null) return;
            
            if (pageIndex < bitmaps.Count)
            {
                var bitmap = bitmaps[pageIndex];
                var pageRect = e.PageBounds;
                var imgRect = GetScaledImageRectangle(bitmap, pageRect);
                
                e.Graphics.DrawImage(bitmap, imgRect);
                
                Console.WriteLine($"Printing page {pagesToPrint[pageIndex] + 1}...");
                pageIndex++;
                e.HasMorePages = pageIndex < bitmaps.Count;
            }
            else
            {
                e.HasMorePages = false;
            }
        };
        
        printDoc.Print();
        Console.WriteLine("PDF print job completed");
    }
    finally
    {
        // Cleanup bitmaps
        foreach (var bitmap in bitmaps)
        {
            bitmap.Dispose();
        }
    }
}

static void PrintTextFile(string textPath, PrintOptions options)
{
    Console.WriteLine("Printing text file...");
    var lines = File.ReadAllLines(textPath);
    var lineIndex = 0;

    using var printDoc = new PrintDocument();
    ConfigurePrintSettings(printDoc, options);

    printDoc.PrintPage += (sender, e) =>
    {
        if (e == null || e.Graphics == null) return;
        
        var font = new Font("Courier New", 10);
        var brush = Brushes.Black;
        var yPos = (float)e.MarginBounds.Top;
        var lineHeight = font.GetHeight(e.Graphics);
        var linesPerPage = (int)(e.MarginBounds.Height / lineHeight);

        while (lineIndex < lines.Length && linesPerPage > 0)
        {
            e.Graphics.DrawString(lines[lineIndex], font, brush, e.MarginBounds.Left, yPos);
            lineIndex++;
            yPos += lineHeight;
            linesPerPage--;
        }

        e.HasMorePages = lineIndex < lines.Length;
    };

    printDoc.Print();
    Console.WriteLine("Text file print job completed");
}

static void PrintImage(string imagePath, PrintOptions options)
{
    Console.WriteLine("Printing image file...");
    using var image = Image.FromFile(imagePath);
    using var printDoc = new PrintDocument();
    
    ConfigurePrintSettings(printDoc, options);

    printDoc.PrintPage += (sender, e) =>
    {
        if (e == null || e.Graphics == null) return;
        
        var pageRect = e.PageBounds;
        var imgRect = GetScaledImageRectangle(image, pageRect);
        
        e.Graphics.DrawImage(image, imgRect);
        e.HasMorePages = false;
    };

    printDoc.Print();
    Console.WriteLine("Image print job completed");
}

static Rectangle GetScaledImageRectangle(Image image, Rectangle pageRect)
{
    var pageWidth = pageRect.Width;
    var pageHeight = pageRect.Height;
    var imgWidth = image.Width;
    var imgHeight = image.Height;
    
    // Calculate scaling factor
    var scaleX = (float)pageWidth / imgWidth;
    var scaleY = (float)pageHeight / imgHeight;
    var scale = Math.Min(scaleX, scaleY);
    
    // Calculate new dimensions
    var newWidth = (int)(imgWidth * scale);
    var newHeight = (int)(imgHeight * scale);
    
    // Center the image
    var x = pageRect.X + (pageWidth - newWidth) / 2;
    var y = pageRect.Y + (pageHeight - newHeight) / 2;
    
    return new Rectangle(x, y, newWidth, newHeight);
}

static void ConfigurePrintSettings(PrintDocument printDoc, PrintOptions options)
{
    Console.WriteLine("Configuring print settings...");
    
    // Set printer name
    if (!string.IsNullOrEmpty(options.PrinterName))
    {
        printDoc.PrinterSettings.PrinterName = options.PrinterName;
        Console.WriteLine($"Printer: {options.PrinterName}");
    }
    else
    {
        Console.WriteLine($"Using default printer: {printDoc.PrinterSettings.PrinterName}");
    }

    // Verify printer exists
    if (!printDoc.PrinterSettings.IsValid)
    {
        throw new Exception($"Printer '{printDoc.PrinterSettings.PrinterName}' is not valid or not found");
    }

    // Set number of copies
    printDoc.PrinterSettings.Copies = (short)options.Copies;
    Console.WriteLine($"Copies: {options.Copies}");

    // Set color
    printDoc.DefaultPageSettings.Color = options.Color;
    Console.WriteLine($"Color: {options.Color}");

    // Set duplex mode
    var duplexMode = (options.Duplex ?? "simplex").ToLower();
    printDoc.PrinterSettings.Duplex = duplexMode switch
    {
        "vertical" => Duplex.Vertical,
        "horizontal" => Duplex.Horizontal,
        _ => Duplex.Simplex
    };
    Console.WriteLine($"Duplex: {duplexMode}");

    // Set paper size
    var paperSizeName = options.PaperSize ?? "A4";
    foreach (PaperSize paperSize in printDoc.PrinterSettings.PaperSizes)
    {
        if (paperSize.PaperName != null && 
            paperSize.PaperName.Equals(paperSizeName, StringComparison.OrdinalIgnoreCase))
        {
            printDoc.DefaultPageSettings.PaperSize = paperSize;
            Console.WriteLine($"Paper Size: {paperSize.PaperName}");
            break;
        }
    }

    // Set orientation
    var orientation = (options.Orientation ?? "portrait").ToLower();
    printDoc.DefaultPageSettings.Landscape = orientation == "landscape";
    Console.WriteLine($"Orientation: {orientation}");
}

static List<int> GetPagesToPrint(string pageRange, int totalPages)
{
    var pages = new List<int>();

    if (pageRange.ToLower() == "all")
    {
        for (int i = 0; i < totalPages; i++)
            pages.Add(i);
        return pages;
    }

    var ranges = pageRange.Split(',');
    foreach (var range in ranges)
    {
        if (range.Contains('-'))
        {
            var parts = range.Split('-');
            var start = int.Parse(parts[0].Trim()) - 1;
            var end = int.Parse(parts[1].Trim()) - 1;
            for (int i = start; i <= end && i < totalPages; i++)
                pages.Add(i);
        }
        else
        {
            var page = int.Parse(range.Trim()) - 1;
            if (page < totalPages)
                pages.Add(page);
        }
    }

    return pages.Distinct().OrderBy(p => p).ToList();
}

// Models
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