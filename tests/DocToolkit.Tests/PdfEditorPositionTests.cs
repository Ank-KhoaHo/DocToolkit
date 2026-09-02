using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// Reading WHERE things are in a PDF, not just what they say (A110) —
/// <see cref="PdfEditor.ExtractWords(byte[])"/> and
/// <see cref="PdfEditor.ExtractImages(byte[])"/>.
///
/// The trap this file exists to avoid is the one A110's own spec names: the original probe
/// asserted that <c>GetImages()</c> was REACHABLE on a page with no images, which passes whether
/// extraction works or not. That is the A38 AcroForm shape. So every assertion here is a literal,
/// and the image tests come in pairs — a literal 1 on a page with an image, and a literal 0 on a
/// page without one. Neither alone discriminates.
/// </summary>
public class PdfEditorPositionTests
{
    private static Task<byte[]> PdfAsync(string html) => HtmlToPdfConverter.ConvertAsync(html);

    /// <summary>
    /// 8x8 truecolour PNG, every pixel rgb(222,17,99), built by hand from IHDR/IDAT/IEND.
    /// Its dimensions are asserted below rather than trusted, so the literal 8 is the oracle.
    /// </summary>
    private static byte[] KnownPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAgAAAAICAIAAABLbSncAAAAEUlEQVR4nGO4J5iMFTEMLQkAI3dUgf4kDHQAAAAASUVORK5CYII=");

    /// <summary>A PDF carrying one known image, built through this library's own machinery.</summary>
    private static byte[] PdfWithKnownImage() =>
        DocxToPdfConverter.Convert(DocxEditor.Create(
            [DocxBlock.Paragraph("Acme Corporation"), DocxBlock.Image(KnownPng(), widthPoints: 48)]));

    // ---------- ExtractWords ----------

    [Fact]
    public async Task FindsAWordAndItsPosition()
    {
        var pdf = await PdfAsync("<h1>Acme Corporation</h1><p>Invoice 42</p>");

        var pages = PdfEditor.ExtractWords(pdf);

        Assert.Single(pages);
        var acme = Assert.Single(pages[0], w => w.Text == "Acme");
        Assert.True(acme.Bounds.Left > 0, $"Left was {acme.Bounds.Left}, expected a real coordinate");
        Assert.True(acme.Bounds.Bottom > 0, $"Bottom was {acme.Bounds.Bottom}, expected a real coordinate");
        Assert.True(acme.Bounds.Width > 0, $"Width was {acme.Bounds.Width}");
    }

    [Fact]
    public async Task GivesDifferentPositionsToWordsOnDifferentLines()
    {
        // THE test of this feature. Returning a constant rectangle for every word - zeroes, or the
        // page box - passes every other assertion in this file. Only a comparison between two
        // words that are genuinely in different places fails against that.
        var pdf = await PdfAsync("<h1>Acme Corporation</h1><p>Invoice 42</p>");

        var words = PdfEditor.ExtractWords(pdf)[0];
        var heading = Assert.Single(words, w => w.Text == "Acme");
        var body = Assert.Single(words, w => w.Text == "Invoice");

        // The heading is printed above the paragraph, and PDF's origin is the page's BOTTOM-left,
        // so the heading's Bottom must be the LARGER of the two.
        Assert.True(
            heading.Bounds.Bottom > body.Bounds.Bottom,
            $"heading Bottom {heading.Bounds.Bottom} should exceed body Bottom {body.Bounds.Bottom}");
    }

    [Fact]
    public async Task BoundsRightAndTopAreDerivedFromLeftAndBottom()
    {
        var pdf = await PdfAsync("<p>Acme</p>");

        var word = Assert.Single(PdfEditor.ExtractWords(pdf)[0], w => w.Text == "Acme");

        Assert.Equal(word.Bounds.Left + word.Bounds.Width, word.Bounds.Right, 6);
        Assert.Equal(word.Bounds.Bottom + word.Bounds.Height, word.Bounds.Top, 6);
    }

    [Fact]
    public async Task TheStreamOverloadAgreesWithTheByteArrayOverload()
    {
        var pdf = await PdfAsync("<p>Acme Corporation</p>");
        using var source = new MemoryStream(pdf);

        var fromStream = await PdfEditor.ExtractWordsAsync(source);

        Assert.Equal(
            PdfEditor.ExtractWords(pdf)[0].Select(w => w.Text),
            fromStream[0].Select(w => w.Text));
    }

    [Fact]
    public async Task TheStreamOverloadDoesNotDisposeOrCloseItsSource()
    {
        var pdf = await PdfAsync("<p>Acme</p>");
        using var source = new MemoryStream(pdf);

        await PdfEditor.ExtractWordsAsync(source);

        Assert.True(source.CanRead, "the caller's stream was closed");
    }

    [Fact]
    public async Task ThePathOverloadAgreesWithTheByteArrayOverload()
    {
        var pdf = await PdfAsync("<p>Acme Corporation</p>");
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
        await File.WriteAllBytesAsync(path, pdf);
        try
        {
            var fromPath = await PdfEditor.ExtractWordsAsync(path);

            Assert.Equal(
                PdfEditor.ExtractWords(pdf)[0].Select(w => w.Text),
                fromPath[0].Select(w => w.Text));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ExtractWordsRejectsNull() =>
        Assert.Throws<ArgumentNullException>(() => PdfEditor.ExtractWords(null!));

    [Fact]
    public void ExtractWordsRejectsEmpty() =>
        Assert.Throws<ArgumentException>(() => PdfEditor.ExtractWords([]));

    [Fact]
    public void ExtractWordsRejectsBytesThatAreNotAPdf()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => PdfEditor.ExtractWords("this is not a PDF"u8.ToArray()));

        Assert.Contains("not actually a PDF", ex.Message, StringComparison.Ordinal);
    }

    // ---------- ExtractImages ----------

    [Fact]
    public void FindsAnEmbeddedImageWithItsPixelSizeAndPlacement()
    {
        var pages = PdfEditor.ExtractImages(PdfWithKnownImage());

        var image = Assert.Single(pages[0]);
        Assert.Equal(8, image.PixelWidth);
        Assert.Equal(8, image.PixelHeight);
        Assert.NotNull(image.Png);
        Assert.NotEmpty(image.Png!);
        // Placed at 48pt wide, which is NOT the 8px stored size - the two are different things.
        Assert.Equal(48d, image.Bounds.Width, 1);
        Assert.True(image.Bounds.Bottom > 0, $"Bottom was {image.Bounds.Bottom}");
    }

    [Fact]
    public async Task ReportsNoImagesForAPageThatHasNone()
    {
        // THE control for the test above. Without it, an implementation that always reported one
        // image - or the original probe's "the API is reachable" - would look identical.
        var pdf = await PdfAsync("<h1>Acme Corporation</h1><p>No pictures here</p>");

        var pages = PdfEditor.ExtractImages(pdf);

        Assert.Single(pages);
        Assert.Empty(pages[0]);
    }

    [Fact]
    public async Task TheImageStreamOverloadAgreesWithTheByteArrayOverload()
    {
        var pdf = PdfWithKnownImage();
        using var source = new MemoryStream(pdf);

        var fromStream = await PdfEditor.ExtractImagesAsync(source);

        Assert.Equal(PdfEditor.ExtractImages(pdf)[0].Count, fromStream[0].Count);
        Assert.Equal(8, Assert.Single(fromStream[0]).PixelWidth);
        Assert.True(source.CanRead, "the caller's stream was closed");
    }

    [Fact]
    public async Task TheImagePathOverloadAgreesWithTheByteArrayOverload()
    {
        var pdf = PdfWithKnownImage();
        var path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".pdf");
        await File.WriteAllBytesAsync(path, pdf);
        try
        {
            var fromPath = await PdfEditor.ExtractImagesAsync(path);

            Assert.Equal(8, Assert.Single(fromPath[0]).PixelHeight);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void ExtractImagesRejectsNull() =>
        Assert.Throws<ArgumentNullException>(() => PdfEditor.ExtractImages(null!));

    [Fact]
    public void ExtractImagesRejectsEmpty() =>
        Assert.Throws<ArgumentException>(() => PdfEditor.ExtractImages([]));

    [Fact]
    public void ExtractImagesRejectsBytesThatAreNotAPdf()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => PdfEditor.ExtractImages("this is not a PDF"u8.ToArray()));

        Assert.Contains("not actually a PDF", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ExtractWordsAsyncRejectsAnUnreadableStream()
    {
        using var closed = new MemoryStream();
        await closed.DisposeAsync();

        await Assert.ThrowsAsync<ArgumentException>(() => PdfEditor.ExtractWordsAsync(closed));
    }

    [Fact]
    public async Task ExtractImagesAsyncRejectsAnUnreadableStream()
    {
        using var closed = new MemoryStream();
        await closed.DisposeAsync();

        await Assert.ThrowsAsync<ArgumentException>(() => PdfEditor.ExtractImagesAsync(closed));
    }

    // ---------- ToString, which is what a failing assertion elsewhere will print ----------
    //
    // Constructed through the internal constructors rather than through a PDF. The null-Png case
    // is a documented normal outcome — an image PdfPig cannot re-encode — and no fixture here
    // produces one, so building it directly is the only way to exercise that branch at all.

    [Fact]
    public void PdfWordToStringNamesTheWordAndItsCorner()
    {
        // Deliberately away from a .x5 midpoint: 72.05 formats as "72.0", because the nearest
        // double to 72.05 is below it. A test sitting on that boundary would be asserting IEEE
        // rounding rather than this method.
        var word = new PdfWord("Acme", new PdfBounds(72.34, 759.94, 30, 12));

        Assert.Equal("Acme @ (72.3,759.9)", word.ToString());
    }

    [Fact]
    public void PdfImageToStringReportsTheByteCountWhenThePixelsDecoded()
    {
        var image = new PdfImage([1, 2, 3], 8, 8, new PdfBounds(72, 700, 48, 48));

        Assert.Equal("8x8px at (72.0,700.0) 48.0x48.0pt, 3 PNG bytes", image.ToString());
    }

    [Fact]
    public void PdfImageToStringSaysSoWhenThePixelsCouldNotBeDecoded()
    {
        var image = new PdfImage(null, 8, 8, new PdfBounds(72, 700, 48, 48));

        Assert.Equal("8x8px at (72.0,700.0) 48.0x48.0pt, no PNG", image.ToString());
    }
}
