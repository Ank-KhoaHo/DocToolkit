using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocToolkit.Tests;

public class DocxEditorReplaceImageTests
{
    /// <summary>
    /// Asserts the package is schema-valid, not merely readable. Extracted text tells you what a
    /// document says, never whether Word will open it — a lesson from the table fixture that built
    /// invalid tables while every text assertion passed.
    /// </summary>
    private static void AssertValid(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var errors = new OpenXmlValidator().Validate(doc).ToList();
        Assert.True(errors.Count == 0,
            "expected a schema-valid package, got:\n" +
            string.Join("\n", errors.Take(3).Select(e => "  " + e.Description)));
    }

    [Fact]
    public void ReplaceImage_SwapsThePlaceholderForADrawing()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("Logo: {{logo}} end")));

        var filled = DocxEditor.ReplaceImage(docx, "{{logo}}", ImageFixtures.Png());

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);

        Assert.Single(doc.MainDocumentPart!.Document!.Body!.Descendants<Drawing>());
        Assert.Single(doc.MainDocumentPart.ImageParts);

        // Only the matched span goes; the text around it stays where it was.
        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("Logo: ", text);
        Assert.Contains(" end", text);
        Assert.DoesNotContain("{{logo}}", text);

        AssertValid(filled);
    }

    [Fact]
    public void ReplaceImage_SizesFromTheImageHeaderByDefault()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("{{logo}}")));

        var filled = DocxEditor.ReplaceImage(docx, "{{logo}}", ImageFixtures.Png(width: 2, height: 3));

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var extent = doc.MainDocumentPart!.Document!.Body!.Descendants<DW.Extent>().Single();

        Assert.Equal(2L * 9525, extent.Cx!.Value);
        Assert.Equal(3L * 9525, extent.Cy!.Value);
    }

    [Fact]
    public void ReplaceImage_ThrowsWhenThePlaceholderIsAbsent()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("nothing to replace")));

        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxEditor.ReplaceImage(docx, "{{logo}}", ImageFixtures.Png()));

        Assert.Contains("{{logo}}", ex.Message);
    }

    // =====================================================================================
    // The traps. Every one of these produces a document that OPENS, so nothing already in
    // the suite would catch them.
    // =====================================================================================

    [Fact]
    public void ReplaceImage_PutsAHeaderImageInTheHeaderPart()
    {
        var docx = DocxFixtures.Build(
            headerText: "Company {{logo}}",
            footerText: null,
            DocxFixtures.P(DocxFixtures.R("body text")));

        var filled = DocxEditor.ReplaceImage(docx, "{{logo}}", ImageFixtures.Png());

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        var header = main.HeaderParts.Single();

        // The image must belong to the HEADER, not the main document part. Getting this wrong gives
        // a relationship id that resolves in the wrong scope - the file opens, showing nothing.
        Assert.Single(header.ImageParts);
        Assert.Empty(main.ImageParts);
        Assert.Single(header.Header!.Descendants<Drawing>());

        AssertValid(filled);
    }

    [Fact]
    public void ReplaceImage_GivesEveryOccurrenceItsOwnIdAndImagePart()
    {
        var docx = DocxFixtures.Build(
            DocxFixtures.P(DocxFixtures.R("{{logo}} first")),
            DocxFixtures.P(DocxFixtures.R("{{logo}} second")));

        var filled = DocxEditor.ReplaceImage(docx, "{{logo}}", ImageFixtures.Png());

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var ids = doc.MainDocumentPart!.Document!.Body!
            .Descendants<DW.DocProperties>().Select(p => p.Id!.Value).ToList();

        Assert.Equal(2, ids.Count);
        // Duplicate ids are what make Word declare the file corrupt and offer to repair it.
        Assert.Equal(ids.Count, ids.Distinct().Count());

        AssertValid(filled);
    }

    [Fact]
    public void ReplaceImage_DoesNotReuseAnIdAlreadyInTheDocument()
    {
        // The previous test only proves two NEW images differ from each other, which stays true even
        // if the id counter is seeded wrongly - nextId++ still yields 1 and 2. The trap is colliding
        // with a drawing the document ALREADY has, which is what makes Word offer to repair the
        // file. So: insert one image, then insert another into the result.
        var docx = DocxFixtures.Build(
            DocxFixtures.P(DocxFixtures.R("{{first}}")),
            DocxFixtures.P(DocxFixtures.R("{{second}}")));

        var once = DocxEditor.ReplaceImage(docx, "{{first}}", ImageFixtures.Png());
        var twice = DocxEditor.ReplaceImage(once, "{{second}}", ImageFixtures.Png());

        using var ms = new MemoryStream(twice);
        using var doc = WordprocessingDocument.Open(ms, false);
        var ids = doc.MainDocumentPart!.Document!.Body!
            .Descendants<DW.DocProperties>().Select(p => p.Id!.Value).ToList();

        Assert.Equal(2, ids.Count);
        Assert.Equal(ids.Count, ids.Distinct().Count());

        AssertValid(twice);
    }

    [Fact]
    public void ReplaceImage_DeclaresTheContentTypeTheBytesActuallyAre()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("{{photo}}")));

        var filled = DocxEditor.ReplaceImage(docx, "{{photo}}", ImageFixtures.Jpeg());

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);

        // A part declaring image/png while holding JPEG bytes renders as a blank frame, silently.
        Assert.Equal("image/jpeg", doc.MainDocumentPart!.ImageParts.Single().ContentType);
        AssertValid(filled);
    }

    [Fact]
    public void ReplaceImage_MatchesAPlaceholderSplitAcrossRuns()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(
            DocxFixtures.R("start {{lo"),
            DocxFixtures.R("go}} end")));

        var filled = DocxEditor.ReplaceImage(docx, "{{logo}}", ImageFixtures.Png());

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("start ", text);
        Assert.Contains(" end", text);
        Assert.DoesNotContain("{{lo", text);

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Single(doc.MainDocumentPart!.Document!.Body!.Descendants<Drawing>());

        AssertValid(filled);
    }

    [Fact]
    public void ReplaceImage_ScalesTheOtherSideWhenOnlyOneIsGiven()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("{{logo}}")));

        var filled = DocxEditor.ReplaceImage(
            docx, "{{logo}}", ImageFixtures.Png(width: 2, height: 3), widthPoints: 10);

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var extent = doc.MainDocumentPart!.Document!.Body!.Descendants<DW.Extent>().Single();

        Assert.Equal(127000L, extent.Cx!.Value);   // 10pt
        Assert.Equal(190500L, extent.Cy!.Value);   // 15pt, keeping 2:3
    }

    [Fact]
    public void ReplaceImage_NamesTheDrawingAfterThePlaceholder()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("{{logo}}")));

        var filled = DocxEditor.ReplaceImage(docx, "{{logo}}", ImageFixtures.Png());

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var properties = doc.MainDocumentPart!.Document!.Body!
            .Descendants<DW.DocProperties>().Single();

        // Braces stripped, so the selection and accessibility panes show something meaningful.
        Assert.Equal("logo", properties.Name!.Value);
        Assert.Equal("logo", properties.Description!.Value);
    }

    [Fact]
    public void ReplaceImage_RejectsAnUnsupportedFormatByName()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("{{logo}}")));

        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxEditor.ReplaceImage(docx, "{{logo}}", ImageFixtures.Gif()));

        Assert.Contains("GIF", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ReplaceImage_RejectsBadArguments()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("{{logo}}")));
        var png = ImageFixtures.Png();

        Assert.Throws<ArgumentNullException>(() => DocxEditor.ReplaceImage(null!, "{{logo}}", png));
        Assert.Throws<ArgumentNullException>(() => DocxEditor.ReplaceImage(docx, null!, png));
        Assert.Throws<ArgumentNullException>(() => DocxEditor.ReplaceImage(docx, "{{logo}}", null!));
        Assert.Throws<ArgumentException>(() => DocxEditor.ReplaceImage(Array.Empty<byte>(), "{{logo}}", png));
        Assert.Throws<ArgumentException>(() => DocxEditor.ReplaceImage(docx, " ", png));
        Assert.Throws<ArgumentException>(() => DocxEditor.ReplaceImage(docx, "{{logo}}", Array.Empty<byte>()));
        Assert.Throws<ArgumentOutOfRangeException>(
            () => DocxEditor.ReplaceImage(docx, "{{logo}}", png, widthPoints: 0));
    }
}
