using DocToolkit;

Console.WriteLine("DocToolkit console sample");
Console.WriteLine("==========================");

// 1. HTML -> DOCX
Console.WriteLine("\n1. HTML -> DOCX");
byte[] docx = await HtmlToDocxConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");
Console.WriteLine($"   Generated {docx.Length} bytes of DOCX.");

// 2. HTML -> PDF (pivots through DOCX internally)
Console.WriteLine("\n2. HTML -> PDF");
byte[] pdf = await HtmlToPdfConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");
Console.WriteLine($"   Generated {pdf.Length} bytes of PDF.");

// 3. DOCX -> PDF
Console.WriteLine("\n3. DOCX -> PDF");
byte[] rendered = DocxToPdfConverter.Convert(docx);
Console.WriteLine($"   Rendered {rendered.Length} bytes of PDF from the DOCX above.");

// 4. Fill a DOCX template, then extract text back out
Console.WriteLine("\n4. DOCX template fill + text extraction");
byte[] template = await HtmlToDocxConverter.ConvertAsync("<p>Customer: {{customer}}</p>");
byte[] filled = DocxEditor.ReplaceText(template, new Dictionary<string, string>
{
    ["{{customer}}"] = "Contoso Ltd",
});
string text = DocxEditor.ExtractText(filled);
Console.WriteLine($"   Extracted text: \"{text.Trim()}\"");

// 5. Spreadsheets: create, read a cell, update it, read again
Console.WriteLine("\n5. XLSX create/read/edit");
byte[] xlsx = WorkbookEditor.Create("Sales", new object?[][]
{
    new object?[] { "Region", "Total" },
    new object?[] { "North", 1200 },
});
string cellBefore = WorkbookEditor.ReadCell(xlsx, "Sales", "B2");
byte[] updated = WorkbookEditor.SetCell(xlsx, "Sales", "B2", 1500);
string cellAfter = WorkbookEditor.ReadCell(updated, "Sales", "B2");
Console.WriteLine($"   B2 before: {cellBefore}, after SetCell: {cellAfter}");

// 6. Presentations: read the shared test fixture PPTX
Console.WriteLine("\n6. PPTX read/edit");
string pptxPath = Path.Combine(AppContext.BaseDirectory, "assets", "sample.pptx");
byte[] pptx = await File.ReadAllBytesAsync(pptxPath);
int slideCount = PresentationEditor.SlideCount(pptx);
IReadOnlyList<string> slideText = PresentationEditor.ExtractText(pptx);
string firstSlide = slideText.Count > 0 ? slideText[0] : "(empty)";
Console.WriteLine($"   {slideCount} slide(s); first slide text: \"{firstSlide}\"");

Console.WriteLine("\nDone.");
