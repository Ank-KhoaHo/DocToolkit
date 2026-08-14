using DocToolkit;

Console.WriteLine("Markdown conversion");
Console.WriteLine("===================");

// --- Markdown in ---------------------------------------------------------------------------
// Markdown is the format documentation, changelogs and LLM output already arrive in. These two
// converters take it straight to a document without a round trip through HTML.

const string markdown = """
    # Q3 report

    Revenue was **up 12%**, driven by:

    - EMEA renewals
    - two enterprise wins in APAC

    | Region | Revenue |
    |--------|---------|
    | EMEA   | 1200    |
    | APAC   | 980     |
    """;

#region convert
byte[] docx = MarkdownToDocxConverter.Convert(markdown);
byte[] pdf = MarkdownToPdfConverter.Convert(markdown);
#endregion

Console.WriteLine($"\nDOCX         : {docx.Length:N0} bytes");
Console.WriteLine($"PDF          : {pdf.Length:N0} bytes");
Console.WriteLine($"Tables in it : {DocxEditor.TableCount(docx)}");

// --- What the conversion could not carry across --------------------------------------------
// The plain Convert overloads discard the report. ConvertWithReport hands it back. The image
// below is remote, and remote images are never fetched - so this is a conversion that succeeds
// while quietly not being what the author wrote, which is exactly the case worth surfacing.

const string withRemoteImage = """
    # Release notes

    ![build status](https://img.example.com/badge.svg)

    Everything shipped.
    """;

#region report
ConversionResult<byte[]> result = MarkdownToDocxConverter.ConvertWithReport(withRemoteImage);

byte[] document = result.Value;          // the same bytes Convert would have returned
if (result.HasLoss)
{
    foreach (ConversionWarning w in result.Warnings)
        Console.WriteLine($"  {w.Kind,-13} {w.Code,-24} {w.Message}");
}
#endregion

Console.WriteLine($"\nHasLoss      : {result.HasLoss}");
Console.WriteLine($"Still valid  : {document.Length:N0} bytes");

// Approximation, not Omission - the image became a hyperlink, so the content is still reachable.
// Reading the text back is how you tell those two apart without trusting the label.
Console.WriteLine($"What it says : {DocxEditor.ExtractText(document).Replace("\n", " / ")}");

// --- Markdown out --------------------------------------------------------------------------
// The other direction, and the lossier one: a Word document can express far more than Markdown
// can. Same shape of API, so the same report tells you what was flattened.

#region round-trip
ConversionResult<string> back = DocxToMarkdownConverter.ConvertWithReport(docx);
#endregion

Console.WriteLine($"\nBack to MD   : {back.Value.Split('\n')[0].Trim()}");
Console.WriteLine($"Warnings     : {back.Warnings.Count}");

foreach (ConversionWarning w in back.Warnings)
    Console.WriteLine($"  {w.Kind,-13} {w.Code,-24} {w.Message}");

Console.WriteLine("\nDone.");
