namespace DocToolkit.Tests;

/// <summary>
/// PowerPoint 97-2003 binary <c>.ppt</c> decks render to PDF through
/// <see cref="PptxToPdfConverter"/>, and this is what says so.
/// </summary>
/// <remarks>
/// <b>Nothing in this repository can produce the input.</b> DocToolkit writes OOXML only, so every
/// other PPTX fixture in this suite is built by the code under test - and a hand-built substitute
/// would be a file this project wrote, which is the one property these fixtures exist to avoid.
/// The deck is a real one from govdocs1; <c>assets/realworld/README.txt</c> carries its provenance.
///
/// <b>Measured 2026-08-20 over the 88 genuine legacy decks in govdocs1 chunk 000: 53 convert,
/// 60.2%</b>, producing 988 pages of which 51 of 53 carry extractable text. So this is a real
/// capability at a real rate, and the rate is published rather than rounded up to "supported".
///
/// <b>The 35 refusals are not arbitrary, which is why the rate is worth stating.</b> One upstream
/// limitation dominates - 20 of them are <i>"Binary PowerPoint groups must contain at least one
/// drawable child"</i>. Nine more are text-encoding preflight failures, the same glyph family
/// <see cref="PdfFontOptions"/> addresses on the DOCX path, and five are control characters in the
/// deck's own text. Every one is a clean refusal: none produced a corrupt PDF.
///
/// <b>This is inherited behaviour, and that is precisely why it is pinned here.</b>
/// <c>PptxToPdfConverter</c> contains no legacy path at all - it hands the bytes to OfficeIMO,
/// whose <c>LegacyPpt</c> support does the work. An inherited capability is one a dependency bump
/// can withdraw silently, exactly as B24 argued about an inherited guarantee. These tests are what
/// would make that a red build rather than a quiet regression.
/// </remarks>
public class LegacyPowerPointTests
{
    private static byte[] LegacyDeck() => File.ReadAllBytes(
        Path.Join(AppContext.BaseDirectory, "assets", "realworld", "legacy-powerpoint-97.ppt"));

    /// <summary>The deck really is the legacy binary format, not a mislabelled .pptx.</summary>
    /// <remarks>
    /// The premise of every assertion below. A .pptx renamed to .ppt would make them all pass
    /// while proving nothing about legacy support - and govdocs1 is a crawl, so mislabelled files
    /// are exactly the sort of thing it contains.
    /// </remarks>
    [Fact]
    public void TheFixtureIsAGenuineBinaryDeck()
    {
        var bytes = LegacyDeck();

        // D0 CF 11 E0 is the OLE compound file signature. 50 4B is a ZIP, which is what OOXML is.
        Assert.True(bytes.Length > 8);
        Assert.Equal(0xD0, bytes[0]);
        Assert.Equal(0xCF, bytes[1]);
        Assert.Equal(0x11, bytes[2]);
        Assert.Equal(0xE0, bytes[3]);
    }

    [Fact]
    public void ALegacyDeckRendersToPdf()
    {
        var pdf = PptxToPdfConverter.Convert(LegacyDeck());

        Assert.True(PdfProbe.IsPdf(pdf));
    }

    [Fact]
    public void ItRendersOnePagePerSlide()
    {
        // Not "it returned a PDF": a renderer emitting one blank page passes the weaker check, and
        // that silent-success shape is the one this repository keeps finding.
        Assert.Equal(2, PdfProbe.MediaBoxes(PptxToPdfConverter.Convert(LegacyDeck())).Count);
    }

    [Fact]
    public void ItCarriesTheDecksOwnText()
    {
        // The assertion that makes the capability real rather than nominal. The literal comes from
        // the deck itself, so a converter that produced two correctly-sized blank pages fails here.
        var text = PdfProbe.ExtractText(PptxToPdfConverter.Convert(LegacyDeck()));

        Assert.Contains("Milestones", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AnUnreadableDeckIsRefusedRatherThanRenderedEmpty()
    {
        // The boundary. Legacy support must not become "accept anything and emit blank pages",
        // which would be worse than refusing - the caller gets a plausible file and no signal.
        Assert.Throws<DocumentConversionException>(
            () => PptxToPdfConverter.Convert([0xD0, 0xCF, 0x11, 0xE0, 1, 2, 3, 4]));
    }
}
