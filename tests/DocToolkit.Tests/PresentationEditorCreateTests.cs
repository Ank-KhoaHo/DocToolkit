using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocToolkit.Tests;

public class PresentationEditorCreateTests
{
    /// <summary>
    /// Asserts the package is schema-valid, not merely readable. Extracted text tells you what a
    /// deck says, never whether PowerPoint will open it.
    /// </summary>
    private static void AssertValid(byte[] pptx)
    {
        using var ms = new MemoryStream(pptx);
        using var doc = PresentationDocument.Open(ms, false);
        var errors = new OpenXmlValidator().Validate(doc).ToList();
        Assert.True(errors.Count == 0,
            "expected a schema-valid package, got:\n" +
            string.Join("\n", errors.Take(3).Select(e => "  " + e.Description)));
    }

    [Fact]
    public void Create_WithNoSlides_ProducesAValidEmptyDeck()
    {
        var pptx = PresentationEditor.Create(Array.Empty<PptxSlide>());

        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, pptx.Take(4).ToArray());
        AssertValid(pptx);
        Assert.Equal(0, PresentationEditor.SlideCount(pptx));
    }

    /// <summary>
    /// The scaffold PowerPoint needs before any slide exists. A master with no layout, or a slide
    /// size of zero, produces a file that opens wrong rather than a file that fails to open.
    /// </summary>
    [Fact]
    public void Create_BuildsTheMasterLayoutAndThemeScaffold()
    {
        var pptx = PresentationEditor.Create(Array.Empty<PptxSlide>());

        using var ms = new MemoryStream(pptx);
        using var doc = PresentationDocument.Open(ms, false);
        var pres = doc.PresentationPart!;

        var master = Assert.Single(pres.SlideMasterParts);
        Assert.Single(master.SlideLayoutParts);

        // A theme is not required by the validator - measured, both ways pass - but it is what
        // supplies the colour and font scheme the master's clrMap points at. See the design doc.
        Assert.NotNull(master.ThemePart);

        // 16:9 at 12192000 x 6858000 EMU. Absent, PowerPoint substitutes its own default.
        Assert.Equal(12192000, pres.Presentation!.SlideSize!.Cx!.Value);
        Assert.Equal(6858000, pres.Presentation.SlideSize.Cy!.Value);
    }

    /// <summary>
    /// The master must be reachable from the presentation by relationship id, not merely present as
    /// a part. A dangling id is schema-valid and loses every slide its inherited formatting.
    /// </summary>
    [Fact]
    public void Create_LinksTheMasterFromThePresentation()
    {
        var pptx = PresentationEditor.Create(Array.Empty<PptxSlide>());

        using var ms = new MemoryStream(pptx);
        using var doc = PresentationDocument.Open(ms, false);
        var pres = doc.PresentationPart!;

        var masterId = Assert.Single(pres.Presentation!.SlideMasterIdList!.Elements<P.SlideMasterId>());
        Assert.Same(
            Assert.Single(pres.SlideMasterParts),
            pres.GetPartById(masterId.RelationshipId!.Value!));
    }

    /// <summary>
    /// The layout must be reachable from the master by relationship id, not merely present as a
    /// part. p:sldLayoutIdLst is optional in the schema and SlideLayoutParts counts physical
    /// relationships rather than id-list entries, so dropping the id list leaves a package that
    /// validates, keeps its layout part, and still passes a parts-count assertion - measured.
    /// This is the analogue of Create_LinksTheMasterFromThePresentation one level down.
    /// </summary>
    [Fact]
    public void Create_LinksTheLayoutFromTheMaster()
    {
        var pptx = PresentationEditor.Create(Array.Empty<PptxSlide>());

        using var ms = new MemoryStream(pptx);
        using var doc = PresentationDocument.Open(ms, false);
        var master = Assert.Single(doc.PresentationPart!.SlideMasterParts);

        var layoutId = Assert.Single(
            master.SlideMaster!.SlideLayoutIdList!.Elements<P.SlideLayoutId>());

        Assert.Same(
            Assert.Single(master.SlideLayoutParts),
            master.GetPartById(layoutId.RelationshipId!.Value!));
    }

    [Fact]
    public void Create_RejectsNullSlides()
    {
        Assert.Throws<ArgumentNullException>(() => PresentationEditor.Create(null!));
    }

    [Fact]
    public void Create_RejectsANullSlideElement()
    {
        var ex = Assert.Throws<ArgumentException>(
            () => PresentationEditor.Create(new PptxSlide[] { null! }));

        Assert.Equal("slides", ex.ParamName);
    }
}
