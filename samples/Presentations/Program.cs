using DocToolkit;

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

byte[] edited = PresentationEditor.ReplaceText(pptx, new Dictionary<string, string>
{
    ["{{who}}"] = "World",
});

IReadOnlyList<string> editedText = PresentationEditor.ExtractText(edited);
Console.WriteLine($"After replace: \"{(editedText.Count > 0 ? editedText[0] : "(empty)")}\"");

// Path.Join, not Path.Combine: Combine silently discards everything before an argument that turns
// out to be rooted, which is a trap in copied sample code the moment the filename becomes a
// variable. Join always concatenates. (CodeQL cs/path-combine flags the Combine form here.)
string outputPath = Path.Join(AppContext.BaseDirectory, "deck.pptx");
await File.WriteAllBytesAsync(outputPath, edited);
Console.WriteLine($"\nWrote {outputPath}");

Console.WriteLine("\nDone.");
