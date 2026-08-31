using DocToolkit;
using OfficeIMO.PowerPoint;

Console.WriteLine("Presentations");
Console.WriteLine("=============");

// Built from data, not from a template - there is no source file to read.
#region create
byte[] pptx = PresentationEditor.Create(new[]
{
    PptxSlide.Titled("Hello {{who}}", "Built from a typed model", "No template file involved"),
    PptxSlide.Titled("Second slide", "Bullets are optional"),
});

int slides = PresentationEditor.SlideCount(pptx);
IReadOnlyList<string> text = PresentationEditor.ExtractText(pptx);
#endregion

Console.WriteLine($"\nSlides       : {slides}");
// ExtractText returns one entry per text-bearing body, not per slide - and Create emits two shapes
// per slide, a title and a content placeholder, so a two-slide deck reports four bodies. Use
// SlideCount for the slide count; text.Count is not it.
Console.WriteLine($"Bodies       : {text.Count}");
Console.WriteLine($"First body   : \"{(text.Count > 0 ? text[0] : "(empty)")}\"");

pptx = PresentationEditor.ReplaceText(pptx, new Dictionary<string, string>
{
    ["{{who}}"] = "World",
});

IReadOnlyList<string> editedText = PresentationEditor.ExtractText(pptx);
Console.WriteLine($"After replace: \"{(editedText.Count > 0 ? editedText[0] : "(empty)")}\"");

// --- A chart on a slide ----------------------------------------------------------------------
// PresentationEditor.AddChart and WorkbookEditor.AddChart (see the Spreadsheets sample) share one
// ChartType/ChartData model.

#region chart
var chartData = new ChartData(
    new[] { "North", "South" },
    new[] { new ChartSeries("Total", new double[] { 1200, 980 }) });

pptx = PresentationEditor.AddChart(
    pptx, slideIndex: 1, ChartType.ColumnClustered, chartData, title: "Regional Totals");
#endregion

// The exact byte count varies between runs - DocumentFormat.OpenXml assigns each saved part a
// fresh random relationship id, which shifts a compressed ZIP's size independent of content.
Console.WriteLine($"\nWith chart   : {pptx.Length:N0} bytes (reaches PptxToPdfConverter's output)");

// --- Reading SmartArt --------------------------------------------------------------------------
// This library cannot AUTHOR a SmartArt diagram yet (see the guide for why), so the deck below is
// built with OfficeIMO.PowerPoint directly - the same way the real diagrams ReadSmartArt is meant
// to read were made. PresentationEditor.AddChart above and everything else in this sample never
// needs that: OfficeIMO.PowerPoint is only reached for this one section.

#region smartart
using (var source = new MemoryStream(pptx, writable: false))
using (var doc = PowerPointPresentation.Load(source))
{
    var box = PowerPointLayoutBox.FromInches(1, 3, 6, 2);
    doc.Slides[0].AddSmartArt(PowerPointSmartArtType.BasicProcess, new[] { "Plan", "Build", "Ship" }, box);

    using var output = new MemoryStream();
    doc.Save(output);
    pptx = output.ToArray();
}

IReadOnlyList<string> diagrams = PresentationEditor.ReadSmartArt(pptx, index: 1);
#endregion

Console.WriteLine($"\nSmartArt     : {diagrams.Count} diagram(s) on slide 1");
// Each entry is already newline-joined across its diagram's nodes - collapsed to one line here
// for a readable console print, same convention as the Spreadsheets sample's CSV output.
string diagramText = diagrams.Count > 0 ? diagrams[0].Replace("\n", " / ") : "(none)";
Console.WriteLine($"Diagram text : \"{diagramText}\"");
Console.WriteLine($"In ExtractText too: {PresentationEditor.ExtractText(pptx).Any(t => t.Contains("Plan"))}");

// Path.Join, not Path.Combine: Combine silently discards everything before an argument that turns
// out to be rooted, which is a trap in copied sample code the moment the filename becomes a
// variable. Join always concatenates. (CodeQL cs/path-combine flags the Combine form here.)
string outputPath = Path.Join(AppContext.BaseDirectory, "deck.pptx");
await File.WriteAllBytesAsync(outputPath, pptx);
Console.WriteLine($"\nWrote {outputPath}");

Console.WriteLine("\nDone.");
