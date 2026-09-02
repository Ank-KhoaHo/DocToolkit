using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

/// <summary>
/// The four page operations core added in 0.21.0: RemovePages, RotatePages, ReorderPages and
/// InsertPages. PdfEditorService's other members (PageCount, Merge, ExtractPages, the metadata
/// pair) are exercised through DI resolution in ServiceCollectionExtensionsTests instead - this
/// file follows the direct-instantiation style the other *ServiceTests.cs files use.
///
/// Each fixture page carries its own PDF page ROTATION (0, 90 or 180) rather than distinct
/// visible text. PdfProbe, which decodes OfficeIMO's hex-string text operators well enough to
/// find a heading, lives in the core test project and is not referenced here. A rotation degree
/// read back with PdfSharp directly (a transitive dependency of Ank.DocToolkit, so no new
/// package is needed) works just as well as a page fingerprint, and it discriminates ORDER, not
/// just count - a wrapper that dropped or moved the wrong page fails on the exact numbers rather
/// than only on how many pages came back.
/// </summary>
public class PdfEditorServiceTests
{
    private static async Task<byte[]> PageAsync(string heading, int rotationDegrees = 0)
    {
        var pdf = await HtmlToPdfConverter.ConvertAsync($"<h1>{heading}</h1>");
        return rotationDegrees == 0 ? pdf : PdfEditor.RotatePages(pdf, 1, 1, rotationDegrees);
    }

    private static int[] RotationsOf(byte[] pdf)
    {
        using var source = new MemoryStream(pdf, writable: false);
        using var document = PdfReader.Open(source, PdfDocumentOpenMode.Import);
        var rotations = new int[document.PageCount];
        for (var i = 0; i < document.PageCount; i++)
        {
            rotations[i] = document.Pages[i].Rotate;
        }
        return rotations;
    }

    [Fact]
    public async Task RemovePages_DropsTheRightPageAndKeepsTheOthers()
    {
        var merged = PdfEditor.Merge(
            [await PageAsync("A"), await PageAsync("B", 90), await PageAsync("C", 180)]);
        var sut = new PdfEditorService();

        var fromWrapper = sut.RemovePages(merged, 2, 1);

        // Concrete: proves the SECOND page (rotation 90) is the one that went missing, not just
        // that the count dropped by one - a wrapper that removed the wrong page would still
        // shrink the count by one and pass a count-only check.
        Assert.Equal([0, 180], RotationsOf(fromWrapper));

        // Also proves delegation: the wrapper agrees with the static method on the same input.
        Assert.Equal(RotationsOf(PdfEditor.RemovePages(merged, 2, 1)), RotationsOf(fromWrapper));
    }

    [Fact]
    public async Task RemovePagesAsync_DropsTheRightPage()
    {
        var merged = PdfEditor.Merge([await PageAsync("A"), await PageAsync("B", 90)]);
        var sut = new PdfEditorService();

        using var source = new MemoryStream(merged, writable: false);
        using var destination = new MemoryStream();
        await sut.RemovePagesAsync(source, 2, 1, destination);

        Assert.Equal([0], RotationsOf(destination.ToArray()));
    }

    [Fact]
    public async Task RotatePages_TurnsOnlyTheRequestedPage()
    {
        var merged = PdfEditor.Merge([await PageAsync("A"), await PageAsync("B")]);
        var sut = new PdfEditorService();

        var fromWrapper = sut.RotatePages(merged, firstPage: 2, count: 1, degrees: 90);

        // Concrete: page 1 is untouched, page 2 carries exactly the requested rotation.
        Assert.Equal([0, 90], RotationsOf(fromWrapper));
        Assert.Equal(
            RotationsOf(PdfEditor.RotatePages(merged, 2, 1, 90)),
            RotationsOf(fromWrapper));
    }

    [Fact]
    public async Task RotatePagesAsync_TurnsTheRequestedRange()
    {
        var merged = PdfEditor.Merge([await PageAsync("A"), await PageAsync("B")]);
        var sut = new PdfEditorService();

        using var source = new MemoryStream(merged, writable: false);
        using var destination = new MemoryStream();
        await sut.RotatePagesAsync(source, 1, 2, 90, destination);

        Assert.Equal([90, 90], RotationsOf(destination.ToArray()));
    }

    [Fact]
    public async Task ReorderPages_PutsThePagesInTheOrderGiven()
    {
        var merged = PdfEditor.Merge(
            [await PageAsync("A"), await PageAsync("B", 90), await PageAsync("C", 180)]);
        var sut = new PdfEditorService();

        var fromWrapper = sut.ReorderPages(merged, [3, 1, 2]);

        // Position-by-position, not just membership - the identity order would pass a
        // count-only or set-only check just as well.
        Assert.Equal([180, 0, 90], RotationsOf(fromWrapper));
        Assert.Equal(
            RotationsOf(PdfEditor.ReorderPages(merged, [3, 1, 2])),
            RotationsOf(fromWrapper));
    }

    [Fact]
    public async Task ReorderPagesAsync_PutsThePagesInTheOrderGiven()
    {
        var merged = PdfEditor.Merge([await PageAsync("A"), await PageAsync("B", 90)]);
        var sut = new PdfEditorService();

        using var source = new MemoryStream(merged, writable: false);
        using var destination = new MemoryStream();
        await sut.ReorderPagesAsync(source, [2, 1], destination);

        Assert.Equal([90, 0], RotationsOf(destination.ToArray()));
    }

    [Fact]
    public async Task InsertPages_PutsTheSourcePagesAtTheGivenPosition()
    {
        var target = PdfEditor.Merge([await PageAsync("A"), await PageAsync("C", 180)]);
        var source = await PageAsync("B", 90);
        var sut = new PdfEditorService();

        var fromWrapper = sut.InsertPages(target, source, atPage: 2);

        Assert.Equal([0, 90, 180], RotationsOf(fromWrapper));
        Assert.Equal(
            RotationsOf(PdfEditor.InsertPages(target, source, 2)),
            RotationsOf(fromWrapper));
    }

    [Fact]
    public async Task InsertPagesAsync_PutsTheSourcePagesAtTheGivenPosition()
    {
        var target = PdfEditor.Merge([await PageAsync("A"), await PageAsync("C", 180)]);
        var source = await PageAsync("B", 90);
        var sut = new PdfEditorService();

        using var targetSource = new MemoryStream(target, writable: false);
        using var sourceSource = new MemoryStream(source, writable: false);
        using var destination = new MemoryStream();
        await sut.InsertPagesAsync(targetSource, sourceSource, 2, destination);

        Assert.Equal([0, 90, 180], RotationsOf(destination.ToArray()));
    }

    /// <summary>
    /// A18-DI. Asserts LITERAL text, not just that the wrapper agrees with the static method:
    /// "the wrapper matches the thing it wraps" holds identically when both return nothing, which
    /// is the tautology B16 exists to catch. The delegation check is kept as a second assertion,
    /// underneath one that can actually fail on its own.
    /// </summary>
    [Fact]
    public async Task ExtractText_ReturnsOneEntryPerPage_HoldingThatPagesOwnText()
    {
        var merged = PdfEditor.Merge([await PageAsync("Acme Corporation"), await PageAsync("Invoice 42")]);
        var sut = new PdfEditorService();

        var pages = sut.ExtractText(merged);

        Assert.Equal(2, pages.Count);

        // Page 2's text is page 2's, not page 1's - a wrapper that returned the first page twice,
        // or concatenated everything, fails here rather than on a count.
        Assert.Contains("Acme Corporation", pages[0], StringComparison.Ordinal);
        Assert.DoesNotContain("Invoice 42", pages[0], StringComparison.Ordinal);
        Assert.Contains("Invoice 42", pages[1], StringComparison.Ordinal);

        Assert.Equal(PdfEditor.ExtractText(merged), pages);
    }

    [Fact]
    public async Task ExtractTextAsync_MatchesTheByteArrayForm_AndLeavesTheSourceOpen()
    {
        var pdf = await PageAsync("Acme Corporation");
        var sut = new PdfEditorService();

        using var source = new MemoryStream(pdf, writable: false);
        var pages = await sut.ExtractTextAsync(source);

        Assert.Contains("Acme Corporation", Assert.Single(pages), StringComparison.Ordinal);

        // The caller owns the stream. CanRead is false once a MemoryStream is disposed.
        Assert.True(source.CanRead, "ExtractTextAsync disposed a source stream it does not own.");
    }

    [Fact]
    public async Task ExtractTextAsync_RejectsASourceItCannotRead()
    {
        var sut = new PdfEditorService();

        await Assert.ThrowsAsync<ArgumentNullException>(() => sut.ExtractTextAsync(null!));

        using var empty = new MemoryStream();
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => sut.ExtractTextAsync(empty));
        Assert.Equal("source", ex.ParamName);
    }

    // ---------------------------------------------------------------------------------------
    // A113-DI: the Stream overloads core added in 0.48.0 for metadata, protection.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public async Task ReadMetadataAsync_MatchesTheStaticMethod()
    {
        var sut = new PdfEditorService();
        var pdf = PdfEditor.WithMetadata(
            await PageAsync("One"), new PdfMetadata { Title = "T", Author = "A" });

        using var source = new MemoryStream(pdf, writable: false);
        var metadata = await sut.ReadMetadataAsync(source);

        Assert.Equal("T", metadata.Title);
        Assert.Equal("A", metadata.Author);
    }

    [Fact]
    public async Task WithMetadataAsync_MatchesTheStaticMethod()
    {
        var sut = new PdfEditorService();

        using var source = new MemoryStream(await PageAsync("One"), writable: false);
        using var destination = new MemoryStream();
        await sut.WithMetadataAsync(source, new PdfMetadata { Title = "Stamped" }, destination);

        Assert.Equal("Stamped", PdfEditor.ReadMetadata(destination.ToArray()).Title);
    }

    // ---------- A110-DI: word positions and embedded images ----------
    //
    // The image fixture cannot come from HTML: HtmlToPdfConverter has no <img> path at all, which
    // A110 measured. The route that works uses the published package's own machinery -
    // DocxBlock.Image -> DocxEditor.Create -> DocxToPdfConverter.Convert - and is available here
    // because the extensions package references core 0.52.0, where all three ship.

    /// <summary>
    /// 8x8 truecolour PNG, every pixel rgb(222,17,99), hand-built from IHDR/IDAT/IEND. Its
    /// dimensions are asserted rather than trusted, so the literal 8 is the oracle.
    /// </summary>
    private static byte[] KnownPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEUlEQVR4nGO4J5iMFTEMLQkAI3dUgf4kDHQAAAAASUVORK5CYII=");

    private static byte[] PdfWithKnownImage() =>
        DocxToPdfConverter.Convert(DocxEditor.Create(
            [DocxBlock.Paragraph("Acme Corporation"), DocxBlock.Image(KnownPng(), widthPoints: 48)]));

    [Fact]
    public async Task ExtractWords_ReturnsWordsWithRealPositions_NotAConstantRectangle()
    {
        var pdf = await HtmlToPdfConverter.ConvertAsync("<h1>Acme Corporation</h1><p>Invoice 42</p>");
        var sut = new PdfEditorService();

        var words = Assert.Single(sut.ExtractWords(pdf));

        // Literals first, per this class's rule. A wrapper returning one constant rectangle for
        // every word passes any count assertion; only a comparison between two words that are
        // genuinely on different lines fails against it. PDF's origin is the page's BOTTOM-left,
        // so the heading's Bottom must be the LARGER of the two.
        var heading = Assert.Single(words, w => w.Text == "Acme");
        var body = Assert.Single(words, w => w.Text == "Invoice");
        Assert.True(
            heading.Bounds.Bottom > body.Bounds.Bottom,
            $"heading Bottom {heading.Bounds.Bottom} should exceed body Bottom {body.Bounds.Bottom}");

        Assert.Equal(PdfEditor.ExtractWords(pdf)[0].Select(w => w.Text), words.Select(w => w.Text));
    }

    [Fact]
    public async Task ExtractWordsAsync_MatchesTheByteArrayForm_AndLeavesTheSourceOpen()
    {
        var pdf = await HtmlToPdfConverter.ConvertAsync("<p>Acme Corporation</p>");
        var sut = new PdfEditorService();

        using var source = new MemoryStream(pdf, writable: false);
        var pages = await sut.ExtractWordsAsync(source);

        Assert.Contains(Assert.Single(pages), w => w.Text == "Acme");
        Assert.True(source.CanRead, "ExtractWordsAsync disposed a source stream it does not own.");
    }

    [Fact]
    public async Task ExtractImages_FindsAnEmbeddedImage_AndReportsNoneWhenThereAreNone()
    {
        var sut = new PdfEditorService();

        var image = Assert.Single(sut.ExtractImages(PdfWithKnownImage())[0]);
        Assert.Equal(8, image.PixelWidth);
        Assert.Equal(8, image.PixelHeight);
        Assert.NotNull(image.Png);
        // Placed at 48pt, which is NOT the 8px stored size - the two are different measurements.
        Assert.Equal(48d, image.Bounds.Width, 1);

        // THE control. Without it a wrapper that always reported one image looks identical, which
        // is the A38 "the API is reachable" shape this repository has been caught by before.
        var textOnly = await HtmlToPdfConverter.ConvertAsync("<p>No pictures here</p>");
        Assert.Empty(Assert.Single(sut.ExtractImages(textOnly)));
    }

    [Fact]
    public async Task ExtractImagesAsync_MatchesTheByteArrayForm_AndLeavesTheSourceOpen()
    {
        var pdf = PdfWithKnownImage();
        var sut = new PdfEditorService();

        using var source = new MemoryStream(pdf, writable: false);
        var pages = await sut.ExtractImagesAsync(source);

        Assert.Equal(8, Assert.Single(pages[0]).PixelHeight);
        Assert.True(source.CanRead, "ExtractImagesAsync disposed a source stream it does not own.");
    }
}
