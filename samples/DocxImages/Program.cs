using DocToolkit;

Console.WriteLine("DOCX images");
Console.WriteLine("===========");

// Both images are inline as base64, so this sample carries no binary asset and borrows no fixture
// from anywhere else in the repo.
//
// PNG and JPEG are read by completely different code paths, which is why both appear here. A PNG
// declares its size at a fixed offset - bytes 16 and 20 of the IHDR chunk. A JPEG does not: its
// size lives in a Start-Of-Frame segment that has to be found by walking the marker chain, past
// segments that carry a length and standalone markers that do not. **The format is decided by the
// bytes themselves, never by a filename** - a file called logo.png holding JPEG bytes is read as
// the JPEG it is, because a mislabelled part renders as a blank frame in Word with no error.

// A real 64x64 PNG, 137 bytes.
const string LogoBase64 =
    "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAAAUElEQVR42u3PQQkAAAgEsOtlFCsZ2gi+hcEKLNXzWgQEBA" +
    "QEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQErsACwGghD5ay/wAAAAAASUVORK5CYII=";

// A real baseline JPEG, 141 bytes - the smallest legal one: single quantisation table, one
// grayscale component, and a Huffman pair holding a single 2-bit code.
const string StampBase64 =
    "/9j/2wBDABAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBD/" +
    "wAALCAABAAEBAREA/8QAFAAAAQAAAAAAAAAAAAAAAAAAAP/EABQQAAEAAAAAAAAAAAAAAAAAAAD/2gAIAQEAAD8AD//Z";

byte[] logo = Convert.FromBase64String(LogoBase64);
byte[] stamp = Convert.FromBase64String(StampBase64);

byte[] letterhead = await HtmlToDocxConverter.ConvertAsync(
    "<p>{{logo}}</p><p>Dear {{customer}}, please find your invoice attached.</p>"
    + "<p>{{stamp}}</p>");

// Size is in points. Give one dimension and the other scales to preserve the aspect ratio; give
// neither and the image's own header decides, read at 96 DPI.
byte[] branded = DocxEditor.ReplaceImage(letterhead, "{{logo}}", logo, widthPoints: 96);

// The same call, a different format. Nothing here tells DocToolkit which it is.
branded = DocxEditor.ReplaceImage(branded, "{{stamp}}", stamp, widthPoints: 24);

branded = DocxEditor.ReplaceText(branded, new Dictionary<string, string>
{
    ["{{customer}}"] = "Contoso Ltd",
});

string text = DocxEditor.ExtractText(branded);

Console.WriteLine($"\nLogo         : {logo.Length:N0}-byte PNG,  placed 96pt wide");
Console.WriteLine($"Stamp        : {stamp.Length:N0}-byte JPEG, placed 24pt wide");
Console.WriteLine($"Document grew: {letterhead.Length:N0} -> {branded.Length:N0} bytes");
Console.WriteLine($"Placeholders replaced: {!text.Contains("{{logo}}") && !text.Contains("{{stamp}}")}");
Console.WriteLine($"Customer set : {text.Contains("Contoso Ltd")}");

Console.WriteLine("\nDone.");
