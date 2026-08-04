using DocToolkit;

Console.WriteLine("Presentations");
Console.WriteLine("=============");

string path = Path.Combine(AppContext.BaseDirectory, "assets", "sample.pptx");
byte[] pptx = await File.ReadAllBytesAsync(path);

int slides = PresentationEditor.SlideCount(pptx);
IReadOnlyList<string> text = PresentationEditor.ExtractText(pptx);

Console.WriteLine($"\nSlides       : {slides}");
Console.WriteLine($"First slide  : \"{(text.Count > 0 ? text[0] : "(empty)")}\"");

byte[] edited = PresentationEditor.ReplaceText(pptx, new Dictionary<string, string>
{
    ["{{who}}"] = "World",
});

IReadOnlyList<string> editedText = PresentationEditor.ExtractText(edited);
Console.WriteLine($"After replace: \"{(editedText.Count > 0 ? editedText[0] : "(empty)")}\"");

Console.WriteLine("\nDone.");
