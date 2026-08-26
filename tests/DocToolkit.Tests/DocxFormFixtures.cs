using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit.Tests;

/// <summary>
/// Documents carrying content controls. Hand-built OOXML, because no public API in this library can
/// author a <c>w:sdt</c> — which is half the gap <see cref="DocxForm"/> closes.
/// </summary>
internal static class DocxFormFixtures
{
    /// <summary>A control with a tag, an alias and some current text.</summary>
    internal static SdtBlock Control(
        string tag, string alias, string shown, params OpenXmlElement[] typeMarkers)
    {
        var properties = new SdtProperties(new Tag { Val = tag }, new SdtAlias { Val = alias });
        foreach (OpenXmlElement marker in typeMarkers) properties.Append(marker);

        return new SdtBlock(properties, new SdtContentBlock(
            new Paragraph(new Run(new Text(shown)))));
    }

    /// <summary>A drop-down control offering <paramref name="options"/>.</summary>
    internal static SdtBlock DropDown(string tag, string alias, string shown, params string[] options)
    {
        var list = new SdtContentDropDownList();
        foreach (string option in options)
            list.Append(new ListItem { DisplayText = option, Value = option });

        return Control(tag, alias, shown, list);
    }

    /// <summary>A four-field form: text, drop-down, date, and a second plain text field.</summary>
    internal static byte[] Form() => Build(
        Control("FullName", "Full name", "Khoa Ho", new SdtContentText()),
        DropDown("Plan", "Plan", "Pro", "Free", "Pro", "Team"),
        Control("Start", "Start date", "15 January 2026", new SdtContentDate
        { FullDate = new DateTime(2026, 1, 15, 0, 0, 0, DateTimeKind.Utc) }),
        Control("Notes", "Notes", "none"));

    internal static byte[] Build(params OpenXmlElement[] blocks)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            MainDocumentPart main = document.AddMainDocumentPart();
            var body = new Body();
            foreach (OpenXmlElement block in blocks) body.Append(block);
            main.Document = new Document(body);
            main.Document.Save();
        }
        return ms.ToArray();
    }
}
