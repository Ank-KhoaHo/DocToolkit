using DocToolkit;

Console.WriteLine("LegacyDoc");
Console.WriteLine("=========");

// A real Word 97-2003 binary, saved by Word itself. It cannot be generated at run time the way the
// other samples build their inputs: this library reads .doc and deliberately does not write it.
string fixture = Path.Join(AppContext.BaseDirectory, "assets", "quarterly-report.doc");
byte[] doc = await File.ReadAllBytesAsync(fixture);

Console.WriteLine($"\nSource: {Path.GetFileName(fixture)} ({doc.Length:N0} bytes)");
Console.WriteLine($"First bytes: {doc[0]:x2} {doc[1]:x2} {doc[2]:x2} {doc[3]:x2} "
                  + "- a compound file, not the PK of a .docx");

// ---------------------------------------------------------------------------------------------
Console.WriteLine("\n1. Reading it - the common case, and it never refuses");

#region legacy-doc-extract-text
// ExtractText takes no options and cannot fail over content loss: what a .doc keeps in its binary
// stream is pictures and form fields, and none of those are text. Table cells are included.
string text = DocToDocxConverter.ExtractText(doc);
#endregion

foreach (var line in text.Split('\n'))
{
    Console.WriteLine($"   {line}");
}

// ---------------------------------------------------------------------------------------------
Console.WriteLine("\n2. Converting it - refuses by default, and that is not an edge case");

#region legacy-doc-convert-refuses
// A legacy .doc keeps pictures, drawings and form fields in a binary stream a .docx cannot carry.
// Rather than hand back a document quietly missing them, Convert throws.
try
{
    byte[] docx = DocToDocxConverter.Convert(doc);
    Console.WriteLine($"   converted without asking: {docx.Length:N0} bytes");
}
catch (DocumentConversionException ex)
{
    Console.WriteLine($"   refused: {ex.Message}");
}
#endregion

Console.WriteLine("\n   Measured: any .doc containing a TABLE has such a stream. Plain text, bold");
Console.WriteLine("   runs and headings do not. Tables are ordinary, so expect to meet this.");

// ---------------------------------------------------------------------------------------------
Console.WriteLine("\n3. Accepting the loss on purpose, and seeing exactly what it was");

#region legacy-doc-accept-the-loss
// ConvertWithReport returns the same bytes Convert would, plus a list of what was dropped - so the
// loss is recorded rather than merely permitted.
ConversionResult<byte[]> result = DocToDocxConverter.ConvertWithReport(
    doc, new LegacyDocOptions { AllowContentLoss = true });

byte[] converted = result.Value;
#endregion

Console.WriteLine($"   converted: {converted.Length:N0} bytes");
foreach (var warning in result.Warnings)
{
    Console.WriteLine($"   [{warning.Kind}] {warning.Code}");
    Console.WriteLine($"       {warning.Message}");
}

// The point of the opt-in: text, tables and formatting all survive. Only the binary payload goes.
Console.WriteLine("\n   Text read back out of the converted .docx:");
foreach (var line in DocxEditor.ExtractText(converted).Split('\n'))
{
    Console.WriteLine($"   {line}");
}

string outputPath = Path.Join(AppContext.BaseDirectory, "quarterly-report.docx");
await File.WriteAllBytesAsync(outputPath, converted);
Console.WriteLine($"\nWrote {outputPath}");

Console.WriteLine("\nDone.");
