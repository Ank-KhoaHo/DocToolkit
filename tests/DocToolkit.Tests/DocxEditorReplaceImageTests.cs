using DocumentFormat.OpenXml.Packaging;
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
        var errors = DocxFixtures.Validate(docx);
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

    /// <summary>
    /// The same image embedded once, however many times its placeholder appears.
    /// </summary>
    /// <remarks>
    /// <b>It was embedded once per OCCURRENCE.</b> Measured 2026-08-20 with a 40 KB image: one
    /// occurrence produced one media part and a 41,983-byte package; three produced
    /// <c>media/image.png</c>, <c>media/image2.png</c> and <c>media/image3.png</c> - three
    /// identical copies - and 122,483 bytes. Linear in occurrences, and a letterhead logo repeated
    /// down a document is the ordinary case rather than a contrived one.
    ///
    /// <b>The count is what discriminates, not the size.</b> Asserting the package merely got
    /// smaller would pass on a change that compressed better while still storing three copies.
    /// </remarks>
    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(5)]
    public void ReplaceImage_EmbedsTheImageOnce_HoweverManyOccurrences(int occurrences)
    {
        var paragraphs = Enumerable.Range(1, occurrences)
            .Select(i => DocxFixtures.P(DocxFixtures.R($"Line {i} {{{{logo}}}} end")))
            .ToArray();

        var filled = DocxEditor.ReplaceImage(DocxFixtures.Build(paragraphs), "{{logo}}", ImageFixtures.Png());

        AssertValid(filled);

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;

        // Every occurrence still becomes a drawing...
        Assert.Equal(occurrences, main.Document!.Body!.Descendants<Drawing>().Count());

        // ...all pointing at ONE part.
        Assert.Single(main.ImageParts);

        // And they really do all resolve to it: a shared relationship id is the point, and a
        // drawing left holding a stale id would open in Word showing nothing.
        var partId = main.GetIdOfPart(main.ImageParts.Single());
        var referenced = main.Document.Body.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()
            .Select(b => b.Embed?.Value)
            .ToList();
        Assert.Equal(occurrences, referenced.Count);
        Assert.All(referenced, id => Assert.Equal(partId, id));
    }

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

    /// <summary>
    /// Every occurrence gets its own <c>wp:docPr</c> id.
    /// </summary>
    /// <remarks>
    /// <b>Renamed 2026-08-20, from <c>...ItsOwnIdAndImagePart</c>.</b> It never asserted the part
    /// count - only that the ids differ - and the "AndImagePart" half became false when occurrences
    /// started sharing one part. A test name is a claim, and this one would have been read as
    /// licensing the duplication it was silently sitting beside.
    ///
    /// The ids must still be distinct, and that half is unchanged: a duplicate <c>wp:docPr/@id</c>
    /// is what makes Word declare the file corrupt. Sharing a relationship id is fine; sharing a
    /// drawing id is not. <see cref="ReplaceImage_EmbedsTheImageOnce_HoweverManyOccurrences"/> holds
    /// the other half.
    /// </remarks>
    [Fact]
    public void ReplaceImage_GivesEveryOccurrenceItsOwnDrawingId()
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
    public void ReplaceImage_GivesDistinctDrawingIdsAcrossBodyAndHeader()
    {
        // The counter that hands out wp:docPr/@id is threaded through every owner InsertImagesIn is
        // called for -- body, header, footer, footnotes, endnotes -- via one `ref uint nextId`. None
        // of the other tests in this file exercise more than one owner in a single ReplaceImage call
        // with the SAME placeholder in both, so a regression that dropped the cross-owner thread
        // (each owner restarting its own id counter at 1) would pass every one of them. This is that
        // case.
        var docx = DocxFixtures.Build(
            headerText: "Company {{logo}}",
            footerText: null,
            DocxFixtures.P(DocxFixtures.R("Body {{logo}}")));

        var filled = DocxEditor.ReplaceImage(docx, "{{logo}}", ImageFixtures.Png());

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        var header = main.HeaderParts.Single();

        var ids = main.Document!.Body!.Descendants<DW.DocProperties>()
            .Concat(header.Header!.Descendants<DW.DocProperties>())
            .Select(p => p.Id!.Value)
            .ToList();

        Assert.Equal(2, ids.Count);
        // Duplicate ids across owners are exactly what make Word declare the file corrupt.
        Assert.Equal(ids.Count, ids.Distinct().Count());

        AssertValid(filled);
    }

    [Fact]
    public void ReplaceImage_DoesNotReuseAnIdAlreadyInTheDocument()
    {
        // ReplaceImage_GivesEveryOccurrenceItsOwnDrawingId only proves two NEW images differ from
        // each other, which stays true even if the id counter is seeded wrongly - nextId++ still
        // yields 1 and 2. The trap is colliding with a drawing the document ALREADY has, which is
        // what makes Word offer to repair the file. So: insert one image, then insert another into
        // the result.
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
        var jpeg = ImageFixtures.Jpeg();

        var filled = DocxEditor.ReplaceImage(docx, "{{photo}}", jpeg);

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);

        // A part declaring image/png while holding JPEG bytes renders as a blank frame, silently.
        var part = doc.MainDocumentPart!.ImageParts.Single();
        Assert.Equal("image/jpeg", part.ContentType);

        // And the part must actually HOLD those bytes. Until this line, deleting the stream.Write in
        // AddImagePart passed the entire suite: the result is a schema-valid document with a
        // correctly typed but EMPTY image part, which Word renders as a blank frame. Nothing else in
        // this repo reads an ImagePart's stream back, on either the edit or the create path.
        using var stored = part.GetStream();
        using var buffer = new MemoryStream();
        stored.CopyTo(buffer);
        Assert.Equal(jpeg, buffer.ToArray());

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
    public async Task ReplaceImageAsync_MatchesTheByteArrayOverload()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("Logo: {{logo}} end")));
        var png = ImageFixtures.Png();

        var expected = DocxEditor.ReplaceImage(docx, "{{logo}}", png, widthPoints: 24);

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        await DocxEditor.ReplaceImageAsync(
            source, "{{logo}}", png, destination, widthPoints: 24);
        var actual = destination.ToArray();

        Assert.Equal(DocxEditor.ExtractText(expected), DocxEditor.ExtractText(actual));

        using var ms = new MemoryStream(actual);
        using var doc = WordprocessingDocument.Open(ms, false);
        var extent = doc.MainDocumentPart!.Document!.Body!.Descendants<DW.Extent>().Single();
        Assert.Equal(24L * 12700, extent.Cx!.Value);

        AssertValid(actual);
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

    /// <summary>
    /// The upper size bound, which applies to THIS already-shipped method and not only to the newer
    /// create path. Before the bound existed, an oversized value returned a document whose drawing
    /// extent had overflowed to a negative number — schema-invalid, produced with no exception.
    ///
    /// This test exists because the bound was added in a commit whose tests all sat on the create
    /// side: reverting the bound reddened three tests and not one of them was here, even though
    /// ReplaceImage shares the same choke point and its observable behaviour changed.
    /// </summary>
    [Fact]
    public void ReplaceImage_RejectsASizeTooLargeForADrawingExtent()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("{{logo}}")));
        var png = ImageFixtures.Png();

        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => DocxEditor.ReplaceImage(docx, "{{logo}}", png, widthPoints: 1_000_000));
        Assert.Equal("widthPoints", ex.ParamName);

        // The derived side counts too: a legal width on a tall image overflows the height.
        Assert.Throws<ArgumentOutOfRangeException>(
            () => DocxEditor.ReplaceImage(docx, "{{logo}}", png, widthPoints: 150_000));
    }
}
