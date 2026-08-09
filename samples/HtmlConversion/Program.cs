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
