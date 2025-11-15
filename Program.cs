using System.Drawing;
using System.Drawing.Printing;
using System.Text.Json;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using PdfiumViewer;

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
                PrintImage(filePath, options);
                break;
            default:
                throw new NotSupportedException($"File type {extension} is not supported");
        }

        return jobId;
    });
}

static void PrintPdf(string pdfPath, PrintOptions options)
{
    using var document = PdfDocument.Load(pdfPath);
    using var printDoc = new PrintDocument();
    
    ConfigurePrintSettings(printDoc, options);
    
    var pageIndex = 0;
    var pagesToPrint = GetPagesToPrint(options.PageRange ?? "all", document.PageCount);

    printDoc.PrintPage += (sender, e) =>
    {
        if (e == null || e.Graphics == null) return;
        
        if (pageIndex < pagesToPrint.Count)
        {
            var currentPage = pagesToPrint[pageIndex];
            using var image = document.Render(currentPage, 300, 300, true);
            e.Graphics.DrawImage(image, e.PageBounds);
            pageIndex++;
            e.HasMorePages = pageIndex < pagesToPrint.Count;
        }
        else
        {
            e.HasMorePages = false;
        }
    };

    printDoc.Print();
}

static void PrintTextFile(string textPath, PrintOptions options)
{
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
}

static void PrintImage(string imagePath, PrintOptions options)
{
    using var image = Image.FromFile(imagePath);
    using var printDoc = new PrintDocument();
    
    ConfigurePrintSettings(printDoc, options);

    printDoc.PrintPage += (sender, e) =>
    {
        if (e == null || e.Graphics == null) return;
        e.Graphics.DrawImage(image, e.PageBounds);
        e.HasMorePages = false;
    };

    printDoc.Print();
}

static void ConfigurePrintSettings(PrintDocument printDoc, PrintOptions options)
{
    // Set printer name
    if (!string.IsNullOrEmpty(options.PrinterName))
    {
        printDoc.PrinterSettings.PrinterName = options.PrinterName;
    }

    // Set number of copies
    printDoc.PrinterSettings.Copies = (short)options.Copies;

    // Set color
    printDoc.DefaultPageSettings.Color = options.Color;

    // Set duplex mode
    var duplexMode = (options.Duplex ?? "simplex").ToLower();
    printDoc.PrinterSettings.Duplex = duplexMode switch
    {
        "vertical" => Duplex.Vertical,
        "horizontal" => Duplex.Horizontal,
        _ => Duplex.Simplex
    };

    // Set paper size
    var paperSizeName = options.PaperSize ?? "A4";
    foreach (PaperSize paperSize in printDoc.PrinterSettings.PaperSizes)
    {
        if (paperSize.PaperName != null && 
            paperSize.PaperName.Equals(paperSizeName, StringComparison.OrdinalIgnoreCase))
        {
            printDoc.DefaultPageSettings.PaperSize = paperSize;
            break;
        }
    }

    // Set orientation
    var orientation = (options.Orientation ?? "portrait").ToLower();
    printDoc.DefaultPageSettings.Landscape = orientation == "landscape";
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