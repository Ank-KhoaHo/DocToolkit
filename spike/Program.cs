using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

// Spike: prove the HTML -> DOCX -> PDF chain works with permissive, pure-managed packages.
//   stage 1  HtmlToOpenXml    (MIT)
//   stage 2  OfficeIMO.Word.Pdf (MIT)

var root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", ".."));
var htmlPath = Path.Combine(root, "sample.html");
var outDir = Path.Combine(root, "out");
Directory.CreateDirectory(outDir);
var docxPath = Path.Combine(outDir, "result.docx");
var pdfPath = Path.Combine(outDir, "result.pdf");

var html = File.ReadAllText(htmlPath);
Console.WriteLine($"input html      : {html.Length:N0} chars");

// ---------- stage 1: HTML -> DOCX ----------
try
{
    using (var doc = WordprocessingDocument.Create(docxPath, WordprocessingDocumentType.Document))
    {
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var converter = new HtmlConverter(mainPart);
        await converter.ParseBody(html);
        mainPart.Document.Save();
    } // dispose flushes the package to disk before we measure it
    Console.WriteLine($"STAGE 1 OK      : docx {new FileInfo(docxPath).Length:N0} bytes -> {docxPath}");
}
catch (Exception ex)
{
    Console.WriteLine($"STAGE 1 FAILED  : {ex.GetType().Name}: {ex.Message}");
    return 1;
}

// ---------- inspect the DOCX so we know what actually survived ----------
try
{
    using var doc = WordprocessingDocument.Open(docxPath, false);
    var body = doc.MainDocumentPart!.Document.Body!;
    Console.WriteLine($"  paragraphs    : {body.Descendants<Paragraph>().Count()}");
    Console.WriteLine($"  tables        : {body.Descendants<Table>().Count()}");
    Console.WriteLine($"  table rows    : {body.Descendants<TableRow>().Count()}");
    Console.WriteLine($"  hyperlinks    : {body.Descendants<Hyperlink>().Count()}");
    Console.WriteLine($"  bold runs     : {body.Descendants<Bold>().Count()}");
    Console.WriteLine($"  italic runs   : {body.Descendants<Italic>().Count()}");
    Console.WriteLine($"  underline runs: {body.Descendants<Underline>().Count()}");
    Console.WriteLine($"  numbering refs: {body.Descendants<NumberingProperties>().Count()}");
    Console.WriteLine($"  coloured runs : {body.Descendants<Color>().Count()}");
}
catch (Exception ex)
{
    Console.WriteLine($"  (docx inspect failed: {ex.Message})");
}

// ---------- stage 2: DOCX -> PDF ----------
try
{
    using var word = WordDocument.Load(docxPath);
    var result = word.SaveAsPdf(pdfPath);
    var len = new FileInfo(pdfPath).Length;
    Console.WriteLine($"STAGE 2 OK      : pdf {len:N0} bytes -> {pdfPath}");

    // PdfSaveResult carries fidelity info - dump whatever it exposes rather than guess.
    Console.WriteLine("  --- PdfSaveResult ---");
    foreach (var p in result.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
    {
        object? v;
        try { v = p.GetValue(result); } catch (Exception ex) { v = $"<{ex.GetType().Name}>"; }
        if (v is System.Collections.IEnumerable e and not string)
        {
            var items = e.Cast<object?>().Select(x => x?.ToString()).ToList();
            Console.WriteLine($"  {p.Name} ({items.Count}):");
            foreach (var it in items.Take(25)) Console.WriteLine($"      - {it}");
        }
        else Console.WriteLine($"  {p.Name} = {v}");
    }

    // Sanity-check the bytes really are a PDF, and count pages.
    var bytes = File.ReadAllBytes(pdfPath);
    var header = System.Text.Encoding.ASCII.GetString(bytes, 0, Math.Min(8, bytes.Length));
    var raw = System.Text.Encoding.Latin1.GetString(bytes);
    var pageObjs = System.Text.RegularExpressions.Regex.Matches(raw, @"/Type\s*/Page[^s]").Count;
    var treeCount = System.Text.RegularExpressions.Regex.Match(raw, @"/Type\s*/Pages.*?/Count\s+(\d+)",
                        System.Text.RegularExpressions.RegexOptions.Singleline);
    Console.WriteLine($"  header        : {header.TrimEnd()}");
    Console.WriteLine($"  page objects  : {pageObjs}");
    Console.WriteLine($"  /Pages /Count : {(treeCount.Success ? treeCount.Groups[1].Value : "not found")}");

    // ---------- stage 3: did the CONTENT survive, or was it truncated? ----------
    // OfficeIMO writes content streams UNCOMPRESSED (no /Filter) and emits text as hex-string
    // Tj operators, e.g. <41636D65> Tj == "Acme". So decode hex literals rather than inflating
    // or substring-searching the raw bytes - both of those silently find nothing here.
    var sb = new System.Text.StringBuilder();
    foreach (System.Text.RegularExpressions.Match m in
             System.Text.RegularExpressions.Regex.Matches(raw, @"<([0-9A-Fa-f]+)>\s*Tj"))
    {
        var h = m.Groups[1].Value;
        if (h.Length % 2 != 0) continue;
        for (int i = 0; i < h.Length; i += 2)
            sb.Append((char)Convert.ToInt32(h.Substring(i, 2), 16));
    }
    var pdfText = sb.ToString();
    Console.WriteLine($"  decoded text  : {pdfText.Length:N0} chars from content streams");

    string[] markers = {
        "Acme Corp", "Invoice", "Billed to", "Contoso", "Summary", "billing portal",
        "net 30", "Line items", "Discovery workshop", "Implementation", "18,100",
        "Total due", "Scope delivered", "Stakeholder interviews", "Next steps",
        "Appendix", "Lorem ipsum", "perspiciatis", "consequuntur magni", "mollitia animi"
    };
    var missing = markers.Where(k => !pdfText.Contains(k, StringComparison.OrdinalIgnoreCase)).ToList();
    Console.WriteLine($"  markers found : {markers.Length - missing.Count}/{markers.Length}");
    if (missing.Count > 0) Console.WriteLine($"  MISSING       : {string.Join(", ", missing)}");
}
catch (Exception ex)
{
    Console.WriteLine($"STAGE 2 FAILED  : {ex.GetType().Name}: {ex.Message}");
    Console.WriteLine(ex.StackTrace);
    return 2;
}

// ---------- stage 4: pagination ----------
// The sample fits one page, so multi-page flow is still unproven. Force it.
try
{
    var rows = string.Concat(Enumerable.Range(1, 60).Select(i =>
        $"<tr><td>Line item number {i} with a reasonably long description</td>" +
        $"<td align=\"right\">{i}</td><td align=\"right\">$950.00</td>" +
        $"<td align=\"right\">${i * 950:N2}</td></tr>"));
    var paras = string.Concat(Enumerable.Range(1, 25).Select(i =>
        $"<p>Paragraph {i}. Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod " +
        "tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud " +
        "exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.</p>"));
    var bigHtml = $"<h1>Pagination stress test</h1>{paras}" +
                  $"<h2>Many line items</h2><table border=\"1\"><thead><tr><th>Description</th>" +
                  $"<th>Qty</th><th>Rate</th><th>Amount</th></tr></thead><tbody>{rows}</tbody></table>" +
                  "<p>FINAL-MARKER-END-OF-DOCUMENT</p>";

    var bigDocx = Path.Combine(outDir, "big.docx");
    var bigPdf = Path.Combine(outDir, "big.pdf");

    using (var doc = WordprocessingDocument.Create(bigDocx, WordprocessingDocumentType.Document))
    {
        var mp = doc.AddMainDocumentPart();
        mp.Document = new Document(new Body());
        await new HtmlConverter(mp).ParseBody(bigHtml);
        mp.Document.Save();
    }

    using var bigWord = WordDocument.Load(bigDocx);
    var bigResult = bigWord.SaveAsPdf(bigPdf);

    var bigBytes = File.ReadAllBytes(bigPdf);
    var bigRaw = System.Text.Encoding.Latin1.GetString(bigBytes);
    var pages = System.Text.RegularExpressions.Regex.Match(bigRaw, @"/Type\s*/Pages.*?/Count\s+(\d+)",
                    System.Text.RegularExpressions.RegexOptions.Singleline);

    // decode hex-string text operators to confirm the last marker made it in
    var hex = new System.Text.StringBuilder();
    foreach (System.Text.RegularExpressions.Match m in
             System.Text.RegularExpressions.Regex.Matches(bigRaw, @"<([0-9A-Fa-f]+)>\s*Tj"))
    {
        var h = m.Groups[1].Value;
        if (h.Length % 2 != 0) continue;
        for (int i = 0; i < h.Length; i += 2)
            hex.Append((char)Convert.ToInt32(h.Substring(i, 2), 16));
    }
    var bigText = hex.ToString();
    var ys = System.Text.RegularExpressions.Regex.Matches(bigRaw, @"1 0 0 1 [-\d.]+ ([-\d.]+) Tm")
                .Select(m => double.Parse(m.Groups[1].Value)).ToList();

    Console.WriteLine($"\nSTAGE 4 (pagination)");
    Console.WriteLine($"  pdf bytes     : {bigBytes.Length:N0}");
    Console.WriteLine($"  succeeded     : {bigResult.Succeeded}, warnings: {bigResult.HasWarnings}");
    Console.WriteLine($"  PAGES         : {(pages.Success ? pages.Groups[1].Value : "?")}");
    Console.WriteLine($"  decoded text  : {bigText.Length:N0} chars");
    Console.WriteLine($"  last row #60  : {bigText.Contains("Line item number 60")}");
    Console.WriteLine($"  final marker  : {bigText.Contains("FINAL-MARKER-END-OF-DOCUMENT")}");
    Console.WriteLine($"  min Y         : {(ys.Count > 0 ? Math.Round(ys.Min(), 1) : 0)}  (negative = drawn off-page)");
    Console.WriteLine($"  positions <0  : {ys.Count(y => y < 0)}");
}
catch (Exception ex)
{
    Console.WriteLine($"STAGE 4 FAILED  : {ex.GetType().Name}: {ex.Message}");
}

Console.WriteLine("\nPIPELINE OK");
return 0;
