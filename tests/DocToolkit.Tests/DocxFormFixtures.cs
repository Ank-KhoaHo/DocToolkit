using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace DocToolkit.Tests;

/// <summary>
/// Documents carrying content controls, for <see cref="DocxForm"/>.
///
/// <b>These are authored through OfficeIMO's own API, and an earlier version of this file was not —
/// which invalidated every measurement taken against it.</b> Hand-built <c>SdtBlock</c> markup with
/// a <c>w:dropDownList</c> or <c>w:date</c> child <i>looks</i> like a typed control and is not one:
/// OfficeIMO recognises typed controls only on an <c>SdtRun</c>, so those markers were inert and
/// every fixture control was a plain structured document tag, whose value validator is
/// <c>value =&gt; true</c>. Measured against the old fixture, <b>no</b> wrong-typed value could be
/// made to fail; against these, three kinds fire immediately.
///
/// The rule this earns: <b>author a fixture the way the library under test authors one</b>, or the
/// thing being measured is the fixture rather than the library.
/// </summary>
internal static class DocxFormFixtures
{
    /// <summary>
    /// A four-field form with three REAL typed controls — drop-down, date picker and check box —
    /// plus a plain text tag.
    /// </summary>
    internal static byte[] Form()
    {
        byte[] blank = Authored(document =>
        {
            document.AddParagraph().AddStructuredDocumentTag("Khoa Ho", "Full name", "FullName");
            document.AddParagraph().AddDropDownList(["Free", "Pro", "Team"], "Plan", "Plan");
            document.AddParagraph().AddDatePicker(new DateTime(2026, 1, 15), "Start date", "Start");
            document.AddParagraph().AddCheckBox(false, "Signed", "Signed");
        });

        // A drop-down added with no selection reads back as null, which is a real state but a poor
        // starting point for a form fixture - so select one, the way a template with a default has.
        return DocxForm.Fill(blank,
            new Dictionary<string, DocxFormValue> { ["Plan"] = DocxFormValue.FromChoice("Pro") });
    }

    /// <summary>A document whose only controls share one tag.</summary>
    internal static byte[] DuplicateTags() => Authored(document =>
    {
        document.AddParagraph().AddStructuredDocumentTag("one", "Same", "Same");
        document.AddParagraph().AddStructuredDocumentTag("two", "Same", "Same");
    });

    /// <summary>
    /// A document with a picture content control.
    /// </summary>
    /// <remarks>
    /// Writes a temporary file because OfficeIMO's <c>AddPictureControl</c> takes a path. That is a
    /// fixture-side concern only — <see cref="DocxForm"/> itself never accepts a path, which is the
    /// point of <see cref="DocxFormValue"/>.
    /// </remarks>
    internal static byte[] WithPictureControl()
    {
        string path = Path.Combine(Path.GetTempPath(), $"docxform-{Guid.NewGuid():N}.png");
        File.WriteAllBytes(path, ImageFixtures.Png());
        try
        {
            return Authored(document =>
                document.AddParagraph().AddPictureControl(path, 32, 32, "Logo", "Logo"));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    /// <summary>A document with no content controls at all.</summary>
    internal static byte[] NoControls() => Authored(document =>
        document.AddParagraph().Text = "an ordinary document");

    /// <summary>Builds a document through OfficeIMO, as Word itself would write one.</summary>
    internal static byte[] Authored(Action<WordDocument> build)
    {
        using var buffer = new MemoryStream();
        using (WordDocument document = WordDocument.Create(buffer))
        {
            build(document);
            document.Save();
        }
        return buffer.ToArray();
    }

    /// <summary>
    /// A hand-built <c>SdtBlock</c> — deliberately NOT a typed control.
    /// </summary>
    /// <remarks>
    /// Kept for the one test that pins the distinction, so the mistake this file's summary describes
    /// cannot be made again silently. Do not build form fixtures with it.
    /// </remarks>
    internal static byte[] UntypedBlockControl(string tag, string shown)
    {
        var properties = new SdtProperties(new Tag { Val = tag }, new SdtAlias { Val = tag });
        var control = new SdtBlock(properties,
            new SdtContentBlock(new Paragraph(new Run(new Text(shown)))));

        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            MainDocumentPart main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(control));
            main.Document.Save();
        }
        return ms.ToArray();
    }
}
