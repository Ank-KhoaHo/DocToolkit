using DocToolkit;

// All three conversions in one file on purpose: HTML -> PDF has no direct renderer under this
// package's constraints (the only free ones are browsers, and a browser is a native binary), so it
// pivots through DOCX internally. Seeing the three side by side is what makes that visible.

const string Html = "<h1>Invoice</h1><p>Total: 18,100.00</p>";

Console.WriteLine("HTML conversion");
Console.WriteLine("===============");

#region convert
byte[] docx = await HtmlToDocxConverter.ConvertAsync(Html);
byte[] pdf = await HtmlToPdfConverter.ConvertAsync(Html);
byte[] rendered = DocxToPdfConverter.Convert(docx);
#endregion

Console.WriteLine($"\nHTML -> DOCX : {docx.Length,7:N0} bytes");
Console.WriteLine($"HTML -> PDF  : {pdf.Length,7:N0} bytes  (pivots through DOCX internally)");
Console.WriteLine($"DOCX -> PDF  : {rendered.Length,7:N0} bytes  (from the DOCX above)");

// --- Getting the bytes out of the process ---------------------------------------------------
// The conversions above hand back a byte[], which is deliberate - this library does not decide
// where your document goes. Writing one to disk is the obvious case and is one line; the other
// two forms below exist so you never have to hold the whole document in memory to do it.
//
// Added after issue #321: every conversion sample stopped at the array and left the reader to
// guess. Path.Join rather than Path.Combine, for the reason given in the Presentations sample.

#region save
// 1. You have the bytes and want a file.
string docxPath = Path.Join(AppContext.BaseDirectory, "invoice.docx");
await File.WriteAllBytesAsync(docxPath, docx);

// 2. You want a file and never needed the array: write straight to any Stream.
string pdfPath = Path.Join(AppContext.BaseDirectory, "invoice.pdf");
await using (var file = File.Create(pdfPath))
{
    await HtmlToPdfConverter.ConvertAsync(Html, file);
}

// 3. Both ends are already paths - no stream, no array.
DocxToPdfConverter.ConvertFile(docxPath, Path.Join(AppContext.BaseDirectory, "from-docx.pdf"));
#endregion

Console.WriteLine($"\nWrote {docxPath}");
Console.WriteLine($"Wrote {pdfPath}");

// The same byte[] is what you would return from an HTTP endpoint, put in a blob store, or pass
// straight back into this library - DocxEditor.ExtractText(docx), DocxToPdfConverter.Convert(docx)
// and the rest all take one. Nothing about it is a file until you make it one.
Console.WriteLine($"\nText read back: {DocxEditor.ExtractText(docx).Replace("\n", " ").Trim()}");

// --- Page setup ----------------------------------------------------------------------------
// Every producer lays out on A4 unless told otherwise. PageSetup is immutable: Landscape() and
// WithMargins() return a NEW instance rather than mutating the shared PageSetup.Letter - which is
// what makes those static properties safe to hand around a running application.

#region page-setup
PageSetup wide = PageSetup.Letter.Landscape().WithMargins(36);

byte[] landscape = await HtmlToDocxConverter.ConvertAsync(Html, wide);
byte[] landscapePdf = await HtmlToPdfConverter.ConvertAsync(Html, wide);
#endregion

Console.WriteLine($"\nDefault      : {PageSetup.A4}");
Console.WriteLine($"This one     : {wide}");
Console.WriteLine($"Landscape    : {landscape.Length,7:N0} bytes DOCX, {landscapePdf.Length:N0} bytes PDF");
Console.WriteLine($"Letter intact: {PageSetup.Letter.WidthPoints < PageSetup.Letter.HeightPoints}  (Landscape() did not mutate it)");

// --- When the input is not what it claims to be ---------------------------------------------
// Everything this library raises on your behalf arrives as one exception type, so a caller needs
// one catch rather than one per underlying library. The original is kept as InnerException.

#region errors
byte[] notADocx = "This is a text file, not a Word document."u8.ToArray();

try
{
    DocxToPdfConverter.Convert(notADocx);
}
catch (DocumentConversionException ex)
{
    Console.WriteLine($"\nRejected     : {ex.Message}");
    Console.WriteLine($"Inner cause  : {ex.InnerException?.GetType().Name ?? "(none)"}");
}
#endregion

Console.WriteLine("\nDone.");
