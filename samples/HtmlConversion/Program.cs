using DocToolkit;

// All three conversions in one file on purpose: HTML -> PDF has no direct renderer under this
// package's constraints (the only free ones are browsers, and a browser is a native binary), so it
// pivots through DOCX internally. Seeing the three side by side is what makes that visible.

const string Html = "<h1>Invoice</h1><p>Total: 18,100.00</p>";

Console.WriteLine("HTML conversion");
Console.WriteLine("===============");

byte[] docx = await HtmlToDocxConverter.ConvertAsync(Html);
Console.WriteLine($"\nHTML -> DOCX : {docx.Length,7:N0} bytes");

byte[] pdf = await HtmlToPdfConverter.ConvertAsync(Html);
Console.WriteLine($"HTML -> PDF  : {pdf.Length,7:N0} bytes  (pivots through DOCX internally)");

byte[] rendered = DocxToPdfConverter.Convert(docx);
Console.WriteLine($"DOCX -> PDF  : {rendered.Length,7:N0} bytes  (from the DOCX above)");

Console.WriteLine("\nDone.");
