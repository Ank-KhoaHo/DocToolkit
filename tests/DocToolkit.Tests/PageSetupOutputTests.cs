using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// What page setup actually lands in the produced file.
///
/// These assert on the <c>w:sectPr</c> element rather than on a round trip, deliberately: a test
/// that writes a document and reads its text back passes identically against a document with no
/// page setup at all, which is exactly the defect this feature exists to fix.
/// </summary>
public class PageSetupOutputTests
{
    private static readonly DocxBlock[] Blocks = { DocxBlock.Paragraph("Hello.") };

    /// <summary>
    /// The <c>w:sectPr</c> of the main document part, or null if there is none. Asserting on null
    /// is what makes "no page setup at all" a visible failure.
    /// </summary>
    private static SectionProperties? SectionPropertiesOf(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.Document!.Body!.Elements<SectionProperties>().SingleOrDefault();
    }

    [Fact]
    public void Create_WithNoPageSetup_EmitsA4()
    {
        var sectPr = SectionPropertiesOf(DocxEditor.Create(Blocks));

        Assert.NotNull(sectPr);
        var size = sectPr!.GetFirstChild<PageSize>();
        Assert.NotNull(size);
        Assert.Equal(11906U, size!.Width!.Value);
        Assert.Equal(16838U, size.Height!.Value);
    }

    [Fact]
    public void Create_WithLetter_EmitsLetter()
    {
        var sectPr = SectionPropertiesOf(DocxEditor.Create(Blocks, PageSetup.Letter));

        var size = sectPr!.GetFirstChild<PageSize>();
        Assert.Equal(12240U, size!.Width!.Value);
        Assert.Equal(15840U, size.Height!.Value);
    }

    [Fact]
    public void Create_WritesTheMarginsInTwentiethsOfAPoint()
    {
        var page = PageSetup.A4.WithMargins(10, 20, 30, 40);

        var sectPr = SectionPropertiesOf(DocxEditor.Create(Blocks, page));

        var margin = sectPr!.GetFirstChild<PageMargin>();
        Assert.NotNull(margin);
        Assert.Equal(200, margin!.Top!.Value);
        Assert.Equal(400U, margin.Right!.Value);
        Assert.Equal(600, margin.Bottom!.Value);
        Assert.Equal(800U, margin.Left!.Value);
    }

    // Word reads the dimensions, but its page-setup UI and several renderers read w:orient. A
    // landscape page whose orient still says portrait is a document that disagrees with itself.
    [Fact]
    public void Create_WithLandscape_SwapsTheDimensionsAndSaysSo()
    {
        var sectPr = SectionPropertiesOf(DocxEditor.Create(Blocks, PageSetup.A4.Landscape()));

        var size = sectPr!.GetFirstChild<PageSize>();
        Assert.Equal(16838U, size!.Width!.Value);
        Assert.Equal(11906U, size.Height!.Value);
        Assert.Equal(PageOrientationValues.Landscape, size.Orient!.Value);
    }

    [Fact]
    public void Create_WithPortraitPage_SaysPortrait()
    {
        var sectPr = SectionPropertiesOf(DocxEditor.Create(Blocks, PageSetup.A4));

        Assert.Equal(
            PageOrientationValues.Portrait,
            sectPr!.GetFirstChild<PageSize>()!.Orient!.Value);
    }

    // sectPr anywhere but last makes Word declare the file corrupt. Nothing else catches this: the
    // document is schema-valid either way and every text-reading test still passes.
    [Fact]
    public void Create_PutsSectionPropertiesLastInTheBody()
    {
        byte[] docx = DocxEditor.Create(
            new[] { DocxBlock.Heading("Title", 1), DocxBlock.Paragraph("Body.") });

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        Assert.IsType<SectionProperties>(body.LastChild);
    }

    [Fact]
    public void Create_WithNoBlocks_StillEmitsPageSetup()
    {
        var sectPr = SectionPropertiesOf(DocxEditor.Create(Array.Empty<DocxBlock>()));

        Assert.NotNull(sectPr);
    }

    [Fact]
    public async Task CreateAsync_HonoursThePageSetup()
    {
        using var destination = new MemoryStream();

        await DocxEditor.CreateAsync(Blocks, PageSetup.Letter, destination);

        var size = SectionPropertiesOf(destination.ToArray())!.GetFirstChild<PageSize>();
        Assert.Equal(12240U, size!.Width!.Value);
    }

    [Fact]
    public async Task CreateToFileAsync_HonoursThePageSetup()
    {
        string path = Path.Join(Path.GetTempPath(), $"{Guid.NewGuid():N}.docx");
        try
        {
            await DocxEditor.CreateToFileAsync(Blocks, PageSetup.Letter, path);

            var size = SectionPropertiesOf(await File.ReadAllBytesAsync(path))!
                .GetFirstChild<PageSize>();
            Assert.Equal(12240U, size!.Width!.Value);
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public void Create_WithNullPageSetup_ThrowsArgumentNullException()
    {
        var ex = Assert.Throws<ArgumentNullException>(() => DocxEditor.Create(Blocks, null!));

        Assert.Equal("page", ex.ParamName);
    }

    // Fractional points are rounded, not truncated: 100.03 pt is 2000.6 twentieths, and truncating
    // would lose a whole twentieth on a value that was only ever an approximation anyway.
    [Fact]
    public void Create_RoundsFractionalPointsToTheNearestTwentieth()
    {
        var page = PageSetup.Custom(100.03, 200);

        var size = SectionPropertiesOf(DocxEditor.Create(Blocks, page))!.GetFirstChild<PageSize>();

        Assert.Equal(2001U, size!.Width!.Value);
    }
}
