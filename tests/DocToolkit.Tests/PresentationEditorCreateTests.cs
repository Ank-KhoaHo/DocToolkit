using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using A = DocumentFormat.OpenXml.Drawing;
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

        Assert.NotNull(master.SlideMaster!.SlideLayoutIdList);
        var layoutId = Assert.Single(master.SlideMaster.SlideLayoutIdList!.Elements<P.SlideLayoutId>());

        var layout = Assert.Single(master.SlideLayoutParts);
        Assert.Same(layout, master.GetPartById(layoutId.RelationshipId!.Value!));

        // The back-relationship: p:sldLayoutIdLst only ever records master->layout. A package
        // missing layout->master is schema-valid (OpenXmlValidator checks part XML against the
        // schema, not the relationship graph) and unopenable in PowerPoint - E_FAIL on
        // Presentations.Open. Measured against PowerPoint 16.0.
        Assert.Same(master, layout.SlideMasterPart);
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

    /// <summary>
    /// Text must round-trip through the EXISTING PresentationEditor.ExtractText. That is the
    /// strongest single test here: it proves a created deck is consumable by the rest of the
    /// package, using code this feature does not own.
    /// </summary>
    [Fact]
    public void Create_WritesTitlesAndBulletsThatExtractTextCanRead()
    {
        var pptx = PresentationEditor.Create(new[]
        {
            PptxSlide.Titled("Q3 Results", "Revenue up 12%", "Costs flat"),
            PptxSlide.Titled("Outlook"),
        });

        AssertValid(pptx);
        Assert.Equal(2, PresentationEditor.SlideCount(pptx));

        // ExtractText reports one entry per text-bearing body (shape), not one per slide - see its
        // own doc comment. A title-and-bullets slide has a separate title shape and a separate
        // bullets shape, so it yields two entries, not one: [0] the title body, [1] the bullets
        // body (one paragraph per bullet, newline-joined), then [2] the title-only second slide.
        var text = PresentationEditor.ExtractText(pptx);
        Assert.Equal(3, text.Count);
        Assert.Contains("Q3 Results", text[0]);
        // Exact equality, not two Assert.Contains: those would pass identically if the bullets came
        // out in the opposite order, which a two-Contains version of this test did not catch.
        Assert.Equal("Revenue up 12%\nCosts flat", text[1]);
        Assert.Contains("Outlook", text[2]);
    }

    /// <summary>
    /// Deck order is p:sldIdLst, NOT part order. PptxFixtures exists partly to prove those can
    /// disagree - it builds slide parts in one order and reverses the id list. A writer that
    /// appends parts and never maintains sldIdLst produces a deck PowerPoint shows in an arbitrary
    /// order, and a test that walks SlideParts would never notice.
    /// </summary>
    [Fact]
    public void Create_OrdersSlidesByTheSlideIdList()
    {
        var pptx = PresentationEditor.Create(new[]
        {
            PptxSlide.Titled("First"),
            PptxSlide.Titled("Second"),
            PptxSlide.Titled("Third"),
        });

        using var ms = new MemoryStream(pptx);
        using var doc = PresentationDocument.Open(ms, false);
        var pres = doc.PresentationPart!;

        var ordered = pres.Presentation!.SlideIdList!.Elements<P.SlideId>()
            .Select(id => (SlidePart)pres.GetPartById(id.RelationshipId!.Value!))
            .Select(part => part.Slide!.InnerText)
            .ToList();

        Assert.Equal(3, ordered.Count);
        Assert.Contains("First", ordered[0]);
        Assert.Contains("Second", ordered[1]);
        Assert.Contains("Third", ordered[2]);
    }

    /// <summary>
    /// Every slide id must be unique and inside 256..2147483647. A duplicate is the PPTX analogue
    /// of a duplicate wp:docPr/@id: PowerPoint declares the file corrupt and offers to repair it,
    /// and no schema check flags it.
    /// </summary>
    [Fact]
    public void Create_GivesEverySlideAUniqueIdInRange()
    {
        var pptx = PresentationEditor.Create(new[]
        {
            PptxSlide.Titled("A"), PptxSlide.Titled("B"), PptxSlide.Titled("C"),
        });

        using var ms = new MemoryStream(pptx);
        using var doc = PresentationDocument.Open(ms, false);

        var ids = doc.PresentationPart!.Presentation!.SlideIdList!
            .Elements<P.SlideId>().Select(s => s.Id!.Value).ToList();

        Assert.Equal(3, ids.Count);
        Assert.Equal(3, ids.Distinct().Count());
        Assert.All(ids, id => Assert.InRange(id, 256U, 2147483647U));
    }

    /// <summary>
    /// Every slide must reference the layout, and by identity rather than by count. Counting parts
    /// passes against a valid-but-wrong reference; only resolving the id catches that. The DOCX
    /// arc's whole-branch review found exactly this hole in its image tests.
    /// </summary>
    [Fact]
    public void Create_PointsEverySlideAtTheLayout()
    {
        var pptx = PresentationEditor.Create(new[]
        {
            PptxSlide.Titled("A"), PptxSlide.Titled("B"),
        });

        using var ms = new MemoryStream(pptx);
        using var doc = PresentationDocument.Open(ms, false);
        var pres = doc.PresentationPart!;
        var layout = Assert.Single(Assert.Single(pres.SlideMasterParts).SlideLayoutParts);

        Assert.Equal(2, pres.SlideParts.Count());
        Assert.All(pres.SlideParts, slide => Assert.Same(layout, slide.SlideLayoutPart));
    }

    /// <summary>
    /// The Stream overload must produce the same deck as the byte[] one — same content, not same
    /// bytes. An OOXML package is a ZIP whose entries carry timestamps, so byte-identity would be a
    /// flaky test wearing the costume of a strict one.
    ///
    /// StreamOverloadTests covers the SHAPE of this overload across every overload at once; what it
    /// does not check is that this particular one writes the right deck.
    /// </summary>
    [Fact]
    public async Task CreateAsync_MatchesTheByteArrayOverload()
    {
        var slides = new[]
        {
            PptxSlide.Titled("Quarterly Report", "Revenue up 12%"),
            PptxSlide.Titled("Outlook"),
        };

        using var destination = new MemoryStream();
        await PresentationEditor.CreateAsync(slides, destination);

        var written = destination.ToArray();
        AssertValid(written);
        Assert.Equal(
            PresentationEditor.ExtractText(PresentationEditor.Create(slides)),
            PresentationEditor.ExtractText(written));
    }

    /// <summary>
    /// Pins the bullet fix directly against the XML: a paragraph named "bullet" is worthless if
    /// nothing in the inheritance chain gives it a:buChar, p:txStyles or a placeholder to inherit
    /// from — none of which this writer supplies, so the paragraph itself must carry it. The title
    /// paragraph must NOT get one; a bulleted title is not what Titled(title, bullets) promises.
    /// This is the only thing that stops a future refactor silently dropping the bullet again.
    /// </summary>
    [Fact]
    public void Create_GivesBulletParagraphsACharacterBulletAndLeavesTheTitleWithout()
    {
        var pptx = PresentationEditor.Create(new[]
        {
            PptxSlide.Titled("Title", "First bullet", "Second bullet"),
        });

        using var ms = new MemoryStream(pptx);
        using var doc = PresentationDocument.Open(ms, false);
        var paragraphs = doc.PresentationPart!.SlideParts.Single().Slide!
            .Descendants<A.Paragraph>().ToList();

        Assert.Equal(3, paragraphs.Count);

        var titleBullet = paragraphs[0].ParagraphProperties?.GetFirstChild<A.CharacterBullet>();
        Assert.Null(titleBullet);

        foreach (var bulletParagraph in paragraphs.Skip(1))
        {
            var bullet = bulletParagraph.ParagraphProperties?.GetFirstChild<A.CharacterBullet>();
            Assert.NotNull(bullet);
            Assert.Equal("•", bullet!.Char!.Value);
        }
    }

    /// <summary>
    /// ReplaceText is a genuinely different code path from Create: it opens the package for write
    /// with PresentationDocument.Open(ms, true) and re-saves the slide part, rather than assembling
    /// fresh parts. This pins that a Create-built deck survives that path rather than assuming a
    /// package this writer produces is automatically compatible with one this project reads back.
    /// </summary>
    [Fact]
    public void ReplaceText_EditsADeckThatCreateBuilt()
    {
        var pptx = PresentationEditor.Create(new[]
        {
            PptxSlide.Titled("Q3 Results", "Revenue up {{pct}}"),
            PptxSlide.Titled("Outlook"),
        });

        var edited = PresentationEditor.ReplaceText(
            pptx, new Dictionary<string, string> { ["{{pct}}"] = "12%" });

        AssertValid(edited);
        Assert.Equal(2, PresentationEditor.SlideCount(edited));

        var text = PresentationEditor.ExtractText(edited);
        Assert.Contains(text, t => t.Contains("Revenue up 12%"));
        Assert.DoesNotContain(text, t => t.Contains("{{pct}}"));
    }

    /// <summary>
    /// Create inherits the promise the rest of this package makes for content sourced from data
    /// rather than markup: a value holding XML metacharacters cannot corrupt the document's
    /// structure. The OpenXml SDK's typed A.Text setter is what is trusted to escape these — this
    /// test is what would catch a future change that concatenates slide text into raw XML instead.
    /// </summary>
    [Fact]
    public void Create_EscapesXmlMetacharactersInTitlesAndBullets()
    {
        const string title = "Report <Q3> & \"Highlights\"";
        const string bullet = "Revenue < costs & \"flat\" 'growth'";

        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled(title, bullet) });

        AssertValid(pptx);

        var text = PresentationEditor.ExtractText(pptx);
        Assert.Equal(2, text.Count);
        Assert.Equal(title, text[0]);
        Assert.Equal(bullet, text[1]);
    }
}
