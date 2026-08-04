using DocToolkit;

Console.WriteLine("DOCX images");
Console.WriteLine("===========");

// A real 64x64 PNG, 137 bytes, inline as base64 so this sample carries no binary asset and needs
// no fixture from anywhere else in the repo. Any PNG or JPEG works the same way.
const string LogoBase64 =
    "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAAAUElEQVR42u3PQQkAAAgEsOtlFCsZ2gi+hcEKLNXzWgQEBA" +
    "QEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQErsACwGghD5ay/wAAAAAASUVORK5CYII=";

byte[] logo = Convert.FromBase64String(LogoBase64);

byte[] letterhead = await HtmlToDocxConverter.ConvertAsync(
    "<p>{{logo}}</p><p>Dear {{customer}}, please find your invoice attached.</p>");

// Size is in points. Give one dimension and the other scales to preserve the aspect ratio; give
// neither and the image's own header decides, read at 96 DPI.
byte[] branded = DocxEditor.ReplaceImage(letterhead, "{{logo}}", logo, widthPoints: 96);

branded = DocxEditor.ReplaceText(branded, new Dictionary<string, string>
{
    ["{{customer}}"] = "Contoso Ltd",
});

string text = DocxEditor.ExtractText(branded);

Console.WriteLine($"\nLogo         : {logo.Length:N0}-byte PNG, placed 96pt wide");
Console.WriteLine($"Document grew: {letterhead.Length:N0} -> {branded.Length:N0} bytes");
Console.WriteLine($"Placeholder replaced: {!text.Contains("{{logo}}")}");
Console.WriteLine($"Customer set : {text.Contains("Contoso Ltd")}");

Console.WriteLine("\nDone.");
