using DocToolkit;

// Everything this package claims, exercised inside a container with no fonts, no LibreOffice, no
// browser and - once the image is built - no network.
//
// The interesting line is the base image in the Dockerfile: `runtime-deps`, not `aspnet` or `sdk`.
// It is the smallest official .NET image there is, and it contains no ICU, no fonts and none of the
// native libraries an imaging or rendering stack would expect to find. If any of this needed one,
// it would fail here rather than on somebody's laptop.

Console.WriteLine("Container");
Console.WriteLine("=========");
Console.WriteLine($"OS  : {System.Runtime.InteropServices.RuntimeInformation.OSDescription.Trim()}");
Console.WriteLine($"Arch: {System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture}");

const string Html = "<h1>Invoice 2026-114</h1><p>Total: <strong>18,100.00</strong></p>";

byte[] docx = await HtmlToDocxConverter.ConvertAsync(Html);
byte[] pdf = await HtmlToPdfConverter.ConvertAsync(Html);

byte[] xlsx = WorkbookEditor.Create("Sales", new object?[][]
{
    new object?[] { "Region", "Total" },
    new object?[] { "North", 1200 },
});
byte[] sheetPdf = XlsxToPdfConverter.Convert(xlsx);

byte[] pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Quarterly", "Revenue up 12%") });
byte[] deckPdf = PptxToPdfConverter.Convert(pptx);

string markdown = DocxToMarkdownConverter.Convert(docx);

Console.WriteLine($"\nHTML  -> DOCX : {docx.Length,7:N0} bytes");
Console.WriteLine($"HTML  -> PDF  : {pdf.Length,7:N0} bytes");
Console.WriteLine($"XLSX  -> PDF  : {sheetPdf.Length,7:N0} bytes");
Console.WriteLine($"PPTX  -> PDF  : {deckPdf.Length,7:N0} bytes");
Console.WriteLine($"DOCX  -> MD   : {markdown.Length,7:N0} chars");

// A conversion that produced nothing would still "succeed" above, so check the content.
bool ok = DocxEditor.ExtractText(docx).Contains("Invoice 2026-114", StringComparison.Ordinal)
          && markdown.Contains("# Invoice 2026-114", StringComparison.Ordinal)
          && pdf.Length > 1000 && sheetPdf.Length > 1000 && deckPdf.Length > 1000;

// --- D14 diagnostics: what does the PDF actually contain here? ------------------------------
var raw = System.Text.Encoding.Latin1.GetString(pdf);
int fontFile2 = System.Text.RegularExpressions.Regex.Matches(raw, "FontFile2").Count;
int fontFile3 = System.Text.RegularExpressions.Regex.Matches(raw, "FontFile3").Count;
long biggest = 0;
foreach (System.Text.RegularExpressions.Match m in
         System.Text.RegularExpressions.Regex.Matches(raw, @"/Length\s+(\d+)"))
    biggest = Math.Max(biggest, long.Parse(m.Groups[1].Value));

Console.WriteLine($"\nPDF FontFile2 streams: {fontFile2}");
Console.WriteLine($"PDF FontFile3 streams: {fontFile3}");
Console.WriteLine($"largest stream       : {biggest:N0} bytes");
foreach (System.Text.RegularExpressions.Match m in
         System.Text.RegularExpressions.Regex.Matches(raw, @"/BaseFont\s*/([A-Za-z0-9+,-]+)"))
    Console.WriteLine($"  BaseFont: {m.Groups[1].Value}");

Console.WriteLine($"\nContent verified: {ok}");
return ok ? 0 : 1;
