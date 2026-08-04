using DocToolkit;

Console.WriteLine("Presentations");
Console.WriteLine("=============");

string path = Path.Combine(AppContext.BaseDirectory, "assets", "sample.pptx");
byte[] pptx = await File.ReadAllBytesAsync(path);

int slides = PresentationEditor.SlideCount(pptx);
IReadOnlyList<string> text = PresentationEditor.ExtractText(pptx);

Console.WriteLine($"\nSlides       : {slides}");
// ExtractText returns one entry per text-bearing body, not per slide - text[0] is the first
// body's text. This fixture has exactly one shape on one slide, so "first body" and "first slide"
// happen to coincide here; that would not hold on a deck with more than one shape, or a table.
Console.WriteLine($"First body   : \"{(text.Count > 0 ? text[0] : "(empty)")}\"");

byte[] edited = PresentationEditor.ReplaceText(pptx, new Dictionary<string, string>
{
    ["{{who}}"] = "World",
});

IReadOnlyList<string> editedText = PresentationEditor.ExtractText(edited);
Console.WriteLine($"After replace: \"{(editedText.Count > 0 ? editedText[0] : "(empty)")}\"");

Console.WriteLine("\nDone.");
