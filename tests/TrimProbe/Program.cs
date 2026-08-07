using DocToolkit;

// Exercises every capability of the public API after trimming, and CHECKS what comes back.
//
// A trim warning says a reflection call might break; only running the trimmed binary says whether
// it did. That distinction is the whole point of this program: the trimmer can remove a type that
// ClosedXML or OpenXML looks up by name, and the failure arrives at runtime as a
// TypeInitializationException or a silently empty document - never at publish time.
//
// So every step asserts on its result rather than merely completing. An exception exits non-zero
// through the runtime's own handler; a wrong-but-not-throwing result exits 1 via Fail.

static void Fail(string what) { Console.Error.WriteLine($"FAIL: {what}"); Environment.Exit(1); }

// --- DOCX: create, read back -------------------------------------------------------------------
byte[] docx = DocxEditor.Create(new DocxBlock[]
{
    DocxBlock.Heading("Quarterly Report", 1),
    DocxBlock.Paragraph("Revenue rose 12% against a flat cost base."),
    DocxBlock.Table(new[] { "Region", "Revenue" }, new[] { new object?[] { "EMEA", 1200 } }),
});
string docxText = DocxEditor.ExtractText(docx);
if (!docxText.Contains("Quarterly Report")) Fail("DOCX heading did not survive the round trip");
if (!docxText.Contains("EMEA")) Fail("DOCX table cell did not survive the round trip");
Console.WriteLine($"docx        ok ({docx.Length} bytes)");

// --- DOCX templating: the reflection-heaviest editing path -------------------------------------
byte[] filled = DocxEditor.ReplaceText(docx, new Dictionary<string, string> { ["EMEA"] = "APAC" });
if (!DocxEditor.ExtractText(filled).Contains("APAC")) Fail("DOCX ReplaceText produced no substitution");
Console.WriteLine("replace     ok");

// --- PPTX --------------------------------------------------------------------------------------
byte[] pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Roadmap", "Ship it") });
if (PresentationEditor.SlideCount(pptx) != 1) Fail("PPTX slide count wrong");
if (!PresentationEditor.ExtractText(pptx).Any(t => t.Contains("Roadmap"))) Fail("PPTX title missing");
Console.WriteLine($"pptx        ok ({pptx.Length} bytes)");

// --- XLSX: ClosedXML, the one assembly that produces a trim warning ----------------------------
byte[] xlsx = WorkbookEditor.Create("Sheet1", new[] { new object?[] { "Item", "Qty" }, new object?[] { "Widget", 42 } });
if (WorkbookEditor.ReadCell(xlsx, "Sheet1", "B2") != "42") Fail("XLSX cell read back wrong");
if (WorkbookEditor.SheetNames(xlsx) is not ["Sheet1"]) Fail("XLSX sheet names wrong");
Console.WriteLine($"xlsx        ok ({xlsx.Length} bytes)");

// --- HTML -> DOCX ------------------------------------------------------------------------------
byte[] fromHtml = await HtmlToDocxConverter.ConvertAsync("<h1>Hello</h1><p>World</p>");
if (!DocxEditor.ExtractText(fromHtml).Contains("Hello")) Fail("HTML to DOCX lost the heading");
Console.WriteLine($"html->docx  ok ({fromHtml.Length} bytes)");

// --- DOCX -> PDF: font handling, the most reflection-dependent path in the graph ---------------
byte[] pdf = DocxToPdfConverter.Convert(fromHtml);
if (pdf.Length < 1000) Fail($"PDF suspiciously small ({pdf.Length} bytes)");
if (!pdf.AsSpan(0, 5).SequenceEqual("%PDF-"u8)) Fail("PDF lacks its magic bytes");
Console.WriteLine($"docx->pdf   ok ({pdf.Length} bytes)");

Console.WriteLine("\nAll capabilities work after trimming.");
