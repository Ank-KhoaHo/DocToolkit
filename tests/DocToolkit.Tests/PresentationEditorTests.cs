using DocumentFormat.OpenXml.Packaging;
using DocToolkit;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocToolkit.Tests;

public class PresentationEditorTests
{
    // Real one-slide deck with a single text box reading "Hello {{who}}", committed at
    // tests/DocToolkit.Tests/assets/sample.pptx and copied next to the test DLL (see the csproj).
    // It was produced once with ShapeCrawler before that package was removed from the codebase —
    // a real PowerPoint-shaped fixture is more realistic than hand-building the OOXML parts.
    private static readonly string SampleAssetPath =
        Path.Join(AppContext.BaseDirectory, "assets", "sample.pptx");

    private static byte[] SampleDeck() => File.ReadAllBytes(SampleAssetPath);

    /// <summary>
    /// Loads the sample deck and splits the single "Hello {{who}}" run in its text box into two
    /// sibling a:r/a:t runs within the same paragraph. PowerPoint itself commonly splits a single
    /// visible word across several runs (spell-check state, formatting changes), so this
    /// reproduces the failure mode a naive per-run Replace would miss.
    /// </summary>
    private static byte[] SampleDeckWithPlaceholderSplitAcrossRuns()
    {
        var bytes = SampleDeck();

        using var ms = new MemoryStream();
        ms.Write(bytes, 0, bytes.Length);
        ms.Position = 0;

        using (var doc = PresentationDocument.Open(ms, true))
        {
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            var slide = slidePart.Slide!;

            var run = slide.Descendants<A.Run>().Single(r => r.Text?.Text == "Hello {{who}}");
            var text = run.Text!;

            // "Hello {{who}}" -> "Hello {{" (first run) + "who}}" (new sibling run), so the
            // "{{who}}" placeholder straddles two a:t elements in the same a:p.
            text.Text = "Hello {{";
            var secondRun = (A.Run)run.CloneNode(true);
            secondRun.Text!.Text = "who}}";
            run.Parent!.InsertAfter(secondRun, run);

            slide.Save();
        }

        return ms.ToArray();
    }

    [Fact]
    public void SlideCount_CountsSlides()
    {
        Assert.Equal(1, PresentationEditor.SlideCount(SampleDeck()));
    }

    [Fact]
    public void ExtractText_ReturnsSlideText()
    {
        var texts = PresentationEditor.ExtractText(SampleDeck());
        Assert.Contains(texts, t => t.Contains("Hello {{who}}"));
    }

    [Fact]
    public void ReplaceText_SubstitutesPlaceholders()
    {
        var edited = PresentationEditor.ReplaceText(SampleDeck(),
            new Dictionary<string, string> { ["{{who}}"] = "world" });

        var texts = PresentationEditor.ExtractText(edited);
        Assert.Contains(texts, t => t.Contains("Hello world"));
        Assert.DoesNotContain(texts, t => t.Contains("{{who}}"));
    }

    [Fact]
    public void ReplaceText_SubstitutesPlaceholderSplitAcrossRuns()
    {
        var edited = PresentationEditor.ReplaceText(SampleDeckWithPlaceholderSplitAcrossRuns(),
            new Dictionary<string, string> { ["{{who}}"] = "world" });

        var texts = PresentationEditor.ExtractText(edited);
        Assert.Contains(texts, t => t.Contains("Hello world"));
        Assert.DoesNotContain(texts, t => t.Contains("{{who}}"));
    }

    // -----------------------------------------------------------------------------------------
    // Formatting (Blocker 3): the old merge wrote the whole paragraph onto run 0.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void ReplaceText_DoesNotImposeTheFirstRunsFormattingOnTheParagraph()
    {
        var deck = PptxFixtures.SampleWithRuns(("Bold ", true), ("plain {{x}} tail", false));

        var edited = PresentationEditor.ReplaceText(deck,
            new Dictionary<string, string> { ["{{x}}"] = "VALUE" });

        Assert.Equal(
            new[] { ("Bold ", true), ("plain VALUE tail", false) },
            PptxFixtures.RunsOfFirstSlide(edited));
    }

    [Fact]
    public void ReplaceText_KeepsTheReplacementInTheRunThatOwnsTheMatchStart()
    {
        var deck = PptxFixtures.SampleWithRuns(("{{na", true), ("me}} tail", false));

        var edited = PresentationEditor.ReplaceText(deck,
            new Dictionary<string, string> { ["{{name}}"] = "VALUE" });

        Assert.Equal(
            new[] { ("VALUE", true), (" tail", false) },
            PptxFixtures.RunsOfFirstSlide(edited));
    }

    // -----------------------------------------------------------------------------------------
    // Deck order (I-4): SlideParts is part-relationship order, not the order PowerPoint shows.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void ExtractText_ReturnsSlidesInDeckOrderNotPartOrder()
    {
        // Parts created as 1,2,3,4; p:sldIdLst then reversed, so the deck reads 4,3,2,1.
        var deck = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2", "Slide 3", "Slide 4" }, reverseDeckOrder: true);

        Assert.Equal(4, PresentationEditor.SlideCount(deck));
        Assert.Equal(
            new[] { "Slide 4", "Slide 3", "Slide 2", "Slide 1" },
            PresentationEditor.ExtractText(deck));
    }

    // -----------------------------------------------------------------------------------------
    // I-5: ExtractText walked only p:sp while ReplaceText walked every a:p, so ReplaceText could
    // change text ExtractText never reported.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void ExtractText_ReportsTableCellTextThatReplaceTextCanReach()
    {
        var deck = PptxFixtures.SampleWithTableCell("Cell {{t}}");
        Assert.Empty(PptxFixtures.Validate(deck));

        Assert.Contains(PresentationEditor.ExtractText(deck), t => t.Contains("Cell {{t}}"));

        var edited = PresentationEditor.ReplaceText(deck,
            new Dictionary<string, string> { ["{{t}}"] = "filled" });

        Assert.Contains(PresentationEditor.ExtractText(edited), t => t.Contains("Cell filled"));
    }

    // -----------------------------------------------------------------------------------------
    // SmartArt (A98): a diagram's text lives in a diagram data part, not a p:txBody, so it was
    // invisible to ExtractText/ReadSlide before ReadSmartArt existed to report it separately.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void ReadSmartArt_ReturnsEachDiagramsNodeTexts()
    {
        var deck = PptxFixtures.DeckWithSmartArt(
            OfficeIMO.PowerPoint.PowerPointSmartArtType.BasicProcess, "Plan", "Build", "Ship");

        var diagrams = PresentationEditor.ReadSmartArt(deck, 1);

        Assert.Single(diagrams);
        Assert.Equal("Plan\nBuild\nShip", diagrams[0]);
    }

    [Fact]
    public void ReadSmartArt_OnASlideWithNoSmartArt_ReturnsEmpty()
    {
        Assert.Empty(PresentationEditor.ReadSmartArt(SampleDeck(), 1));
    }

    [Fact]
    public void ReadSmartArt_IndexBelowOne_Throws()
    {
        Assert.Throws<ArgumentOutOfRangeException>(
            () => PresentationEditor.ReadSmartArt(SampleDeck(), 0));
    }

    [Fact]
    public void ReadSmartArt_IndexAboveSlideCount_Throws()
    {
        Assert.Throws<ArgumentOutOfRangeException>(
            () => PresentationEditor.ReadSmartArt(SampleDeck(), 2));
    }

    [Fact]
    public void ExtractText_IncludesSmartArtAlongsideOrdinaryText()
    {
        var deck = PptxFixtures.DeckWithSmartArt(
            OfficeIMO.PowerPoint.PowerPointSmartArtType.BasicProcess, "Plan", "Build", "Ship");

        var texts = PresentationEditor.ExtractText(deck);

        // Both halves present: the sample's own text-bearing shape, untouched by the fixture,
        // and the SmartArt diagram this test added.
        Assert.Contains(texts, t => t.Contains("Hello {{who}}"));
        Assert.Contains(texts, t => t == "Plan\nBuild\nShip");
    }

    [Fact]
    public void ExtractText_OnADeckWithNoSmartArt_IsUnchanged()
    {
        // The additive-only claim above needs a negative half: a deck with no SmartArt at all
        // must report exactly what it always did, not an empty diagram entry alongside it.
        var texts = PresentationEditor.ExtractText(SampleDeck());
        Assert.Single(texts);
        Assert.Contains("Hello {{who}}", texts[0]);
    }

    [Fact]
    public void ReplaceText_LeavesSmartArtDataUntouched()
    {
        // Regression pin for a probe finding (artifacts/a98-probe): a no-op ReplaceText round
        // trip must not corrupt or drop SmartArt data already in the document.
        var deck = PptxFixtures.DeckWithSmartArt(
            OfficeIMO.PowerPoint.PowerPointSmartArtType.BasicProcess, "Plan", "Build", "Ship");

        var edited = PresentationEditor.ReplaceText(deck, new Dictionary<string, string>());

        Assert.Equal(1, PresentationEditor.SlideCount(edited));
        var diagrams = PresentationEditor.ReadSmartArt(edited, 1);
        Assert.Single(diagrams);
        Assert.Equal("Plan\nBuild\nShip", diagrams[0]);
    }

    [Fact]
    public async Task ReadSmartArtAsync_FromFile_MatchesTheByteArrayOverload()
    {
        var pptx = PptxFixtures.DeckWithSmartArt(
            OfficeIMO.PowerPoint.PowerPointSmartArtType.BasicProcess, "Plan", "Build", "Ship");

        using var input = new TempFile();
        await File.WriteAllBytesAsync(input.Path, pptx);

        Assert.Equal(
            PresentationEditor.ReadSmartArt(pptx, 1),
            await PresentationEditor.ReadSmartArtAsync(input.Path, 1));
    }

    [Fact]
    public void ReadSmartArt_WrapsCorruptInputInDocumentConversionException()
    {
        Assert.Throws<DocumentConversionException>(
            () => PresentationEditor.ReadSmartArt(new byte[] { 1, 2, 3, 4, 5 }, 1));
    }

    // -----------------------------------------------------------------------------------------
    // Error handling (I-6).
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void ExtractText_WrapsCorruptInputInDocumentConversionException()
    {
        Assert.Throws<DocumentConversionException>(
            () => PresentationEditor.ExtractText(new byte[] { 1, 2, 3, 4, 5 }));
    }

    // -----------------------------------------------------------------------------------------
    // File-path overloads.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public async Task SlideCountAsync_FromFile_MatchesTheByteArrayOverload()
    {
        var pptx = PptxFixtures.Sample();

        using var input = new TempFile();
        await File.WriteAllBytesAsync(input.Path, pptx);

        Assert.Equal(
            PresentationEditor.SlideCount(pptx),
            await PresentationEditor.SlideCountAsync(input.Path));
    }

    [Fact]
    public async Task ExtractTextAsync_FromFile_MatchesTheByteArrayOverload()
    {
        var pptx = PptxFixtures.Sample();

        using var input = new TempFile();
        await File.WriteAllBytesAsync(input.Path, pptx);

        Assert.Equal(
            PresentationEditor.ExtractText(pptx),
            await PresentationEditor.ExtractTextAsync(input.Path));
    }

    [Fact]
    public async Task ReplaceTextAsync_FromFileToFile_SubstitutesThePlaceholder()
    {
        var pptx = PptxFixtures.Sample();
        var replacements = new Dictionary<string, string> { ["{{who}}"] = "World" };

        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, pptx);

        await PresentationEditor.ReplaceTextAsync(input.Path, output.Path, replacements);

        var text = await PresentationEditor.ExtractTextAsync(output.Path);
        Assert.Contains(text, entry => entry.Contains("World"));
        Assert.DoesNotContain(text, entry => entry.Contains("{{who}}"));
    }

    // -----------------------------------------------------------------------------------------
    // Task 3 fixture: DeckWithPlaceholderBox must carry an explicit a:xfrm of its own, not one
    // inherited from a layout, since a layout-inherited shape is exactly what ReplaceImage must
    // reject.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void TheFixtureBoxCarriesTheExactPositionItWasAskedFor()
    {
        // Pins the fixture itself. Deliberately NOT the method's own defaults (x=1000000,
        // y=2000000, cx=4000000, cy=3000000) — passing those would let a fixture that ignored its
        // arguments and always wrote the defaults pass this test too. Every ReplaceImage
        // assertion below is computed FROM these numbers, so a fixture that quietly placed the box
        // elsewhere would make them all agree with each other and with nothing real.
        var pptx = PptxFixtures.DeckWithPlaceholderBox("{{chart}}", 555000, 666000, 777000, 888000);
        Assert.Empty(PptxFixtures.Validate(pptx));

        using var ms = new MemoryStream(pptx);
        using var doc = PresentationDocument.Open(ms, false);

        var shape = doc.PresentationPart!.SlideParts.Single()
                       .Slide!.CommonSlideData!.ShapeTree!.Elements<P.Shape>().Single();
        var xfrm = shape.ShapeProperties!.Transform2D!;

        Assert.Equal(555000L, xfrm.Offset!.X!.Value);
        Assert.Equal(666000L, xfrm.Offset.Y!.Value);
        Assert.Equal(777000L, xfrm.Extents!.Cx!.Value);
        Assert.Equal(888000L, xfrm.Extents.Cy!.Value);
        Assert.Equal("{{chart}}", string.Concat(shape.Descendants<A.Text>().Select(t => t.Text)));
    }

    // -----------------------------------------------------------------------------------------
    // Task 4: ReplaceImage(byte[]) and its two refusals.
    // -----------------------------------------------------------------------------------------

    private static byte[] Png() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAEAAAAAwCAIAAAD1Nh4LAAAAI0lEQVR4nO3BAQ0AAADCoP" +
        "dPbQ8HFAAAAAAAAAAAAAAAAAAA8G0hAAABmmDh1QAAAABJRU5ErkJggg==");

    [Fact]
    public void ReplaceImagePutsAPictureWhereTheBoxWas()
    {
        var pptx = PptxFixtures.DeckWithPlaceholderBox("{{chart}}", 1000000, 2000000, 4000000, 3000000);

        var filled = PresentationEditor.ReplaceImage(pptx, "{{chart}}", Png());

        using var ms = new MemoryStream(filled);
        using var doc = PresentationDocument.Open(ms, false);
        var slidePart = doc.PresentationPart!.SlideParts.Single();
        var tree = slidePart.Slide!.CommonSlideData!.ShapeTree!;

        var picture = Assert.Single(tree.Elements<P.Picture>());
        Assert.Empty(tree.Elements<P.Shape>());     // the text box is GONE, not merely joined

        // The image is 64x48 px at 96 DPI, i.e. 4:3 — the same ratio as the box — so it fills the
        // box exactly and sits at the box origin. Computed, not copied from a previous run.
        var xfrm = picture.ShapeProperties!.Transform2D!;
        Assert.Equal(1000000L, xfrm.Offset!.X!.Value);
        Assert.Equal(2000000L, xfrm.Offset.Y!.Value);
        Assert.Equal(4000000L, xfrm.Extents!.Cx!.Value);
        Assert.Equal(3000000L, xfrm.Extents.Cy!.Value);

        // The blip must resolve against THIS slide part, which is the ownership trap.
        var embed = picture.BlipFill!.Blip!.Embed!.Value!;
        Assert.NotNull(slidePart.GetPartById(embed));
    }

    [Fact]
    public void ReplaceImageFitsAMismatchedAspectRatioRatherThanForwardingTheBox()
    {
        // A SQUARE box (4000000x4000000 EMU) against the 64x48 (4:3) test image. Unlike
        // ReplaceImagePutsAPictureWhereTheBoxWas, box and image ratios deliberately differ, so a
        // wiring bug that forwards the shape's raw a:xfrm instead of PptxPictureFactory.Fit's
        // output produces a DIFFERENT (wrong) result here rather than passing by coincidence.
        var pptx = PptxFixtures.DeckWithPlaceholderBox(
            "{{chart}}", x: 1000000, y: 2000000, cx: 4000000, cy: 4000000);

        var filled = PresentationEditor.ReplaceImage(pptx, "{{chart}}", Png());

        using var ms = new MemoryStream(filled);
        using var doc = PresentationDocument.Open(ms, false);
        var slidePart = doc.PresentationPart!.SlideParts.Single();
        var tree = slidePart.Slide!.CommonSlideData!.ShapeTree!;

        var picture = Assert.Single(tree.Elements<P.Picture>());
        Assert.Empty(tree.Elements<P.Shape>());

        // Fit rule: scale = min(boxCx/imageCx, boxCy/imageCy), image scaled by that factor and
        // centred in the box.
        //   scale = min(4000000/64, 4000000/48) = min(62500, 83333.33) = 62500   (width binds)
        //   cx    = round(64 * 62500) = 4000000
        //   cy    = round(48 * 62500) = 3000000
        //   x     = boxX + (boxCx - cx) / 2 = 1000000 + (4000000 - 4000000) / 2 = 1000000
        //   y     = boxY + (boxCy - cy) / 2 = 2000000 + (4000000 - 3000000) / 2 = 2500000
        // Width fills the box exactly; height is letterboxed with 500000 EMU of equal slack above
        // (2000000..2500000) and below (5500000..6000000).
        var xfrm = picture.ShapeProperties!.Transform2D!;
        Assert.Equal(1000000L, xfrm.Offset!.X!.Value);
        Assert.Equal(2500000L, xfrm.Offset.Y!.Value);
        Assert.Equal(4000000L, xfrm.Extents!.Cx!.Value);
        Assert.Equal(3000000L, xfrm.Extents.Cy!.Value);
    }

    [Fact]
    public void TheImagePartBelongsToTheSlideNotThePresentation()
    {
        var pptx = PptxFixtures.DeckWithPlaceholderBox("{{chart}}");

        var filled = PresentationEditor.ReplaceImage(pptx, "{{chart}}", Png());

        using var ms = new MemoryStream(filled);
        using var doc = PresentationDocument.Open(ms, false);

        Assert.Single(doc.PresentationPart!.SlideParts.Single().ImageParts);
        // PresentationPart has no ImageParts property of its own (unlike SlidePart) — walk parts
        // by type instead, which is exactly what the ownership assertion needs anyway.
        Assert.Empty(doc.PresentationPart.GetPartsOfType<ImagePart>());
    }

    [Fact]
    public void APlaceholderSplitAcrossRunsIsFoundAndReplaced()
    {
        // Every fixture above writes the placeholder into a single a:t run. PowerPoint routinely
        // splits one visible string across several runs (spell-check state, formatting changes),
        // so this is the only test that would catch a regression from matching against each
        // paragraph's concatenated text back to a naive per-run Contains check.
        var pptx = PptxFixtures.SampleWithRuns(("{{ch", false), ("art}}", false));

        var filled = PresentationEditor.ReplaceImage(pptx, "{{chart}}", Png());

        using var ms = new MemoryStream(filled);
        using var doc = PresentationDocument.Open(ms, false);
        var tree = doc.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.ShapeTree!;

        Assert.Single(tree.Elements<P.Picture>());
        Assert.Empty(tree.Elements<P.Shape>());
    }

    [Fact]
    public void ADeckBuiltByCreateWorksWithReplaceImage()
    {
        // The design originally claimed a deck built by PresentationEditor.Create would be
        // refused by ReplaceImage for carrying no explicit position, the same way an
        // unpositioned layout placeholder is. That claim was wrong: PptxDocumentWriter.TextShape
        // gives every shape it builds, title included, its own a:xfrm - there is nothing
        // layout-inherited about it. This test exists so the corrected claim cannot silently
        // regress back to the wrong one.
        var deck = PresentationEditor.Create(new[] { PptxSlide.Titled("{{chart}}") });

        var filled = PresentationEditor.ReplaceImage(deck, "{{chart}}", Png());

        using var ms = new MemoryStream(filled);
        using var doc = PresentationDocument.Open(ms, false);
        var tree = doc.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.ShapeTree!;

        Assert.Single(tree.Elements<P.Picture>());
        Assert.Empty(tree.Elements<P.Shape>());
    }

    [Fact]
    public void AShapeHoldingMoreThanThePlaceholderIsRefused()
    {
        // Refusing beats replacing: the unit swapped is the whole shape, so proceeding would
        // destroy "Chart: " and " (Q3)" with no error and no schema violation.
        var pptx = PptxFixtures.DeckWithPlaceholderBox("Chart: {{chart}} (Q3)");

        var ex = Assert.Throws<DocumentConversionException>(
            () => PresentationEditor.ReplaceImage(pptx, "{{chart}}", Png()));

        Assert.Contains("only", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void APlaceholderThatMatchesNothingIsRefused()
    {
        var pptx = PptxFixtures.DeckWithPlaceholderBox("{{chart}}");

        Assert.Throws<DocumentConversionException>(
            () => PresentationEditor.ReplaceImage(pptx, "{{missing}}", Png()));
    }

    [Fact]
    public void AShapeWithNoExplicitPositionIsRefused()
    {
        // The third of the feature's three refusals — a shape holding extra text and a
        // placeholder matching nothing are both covered above; this is the one that was shipping
        // untested: a shape that inherits its position from a layout rather than carrying its own
        // a:xfrm, so there is nowhere to put the replacement picture.
        var pptx = PptxFixtures.DeckWithUnpositionedPlaceholder("{{chart}}");

        // The fixture's doc claims removing the a:xfrm leaves the deck schema-valid. Prove it,
        // rather than stating it: if this deck were invalid the refusal below could be firing for
        // the wrong reason entirely, and the test would still look green.
        Assert.Empty(PptxFixtures.Validate(pptx));

        var ex = Assert.Throws<DocumentConversionException>(
            () => PresentationEditor.ReplaceImage(pptx, "{{chart}}", Png()));

        Assert.Contains("Draw a text box", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void APlaceholderInsideAGroupIsRefusedWithAnAccurateMessage()
    {
        // ReplaceImage only walks a slide's direct p:sp children - a shape inside a p:grpSp
        // carries coordinates in the group's own space, so placing a picture there from
        // slide-space numbers would put it somewhere unrelated. The placeholder genuinely exists
        // on this slide, so the generic "does not appear in any shape" message would be false;
        // this pins the message that names the real cause instead.
        var pptx = PptxFixtures.SampleWithPlaceholderInGroup("{{chart}}");
        Assert.Empty(PptxFixtures.Validate(pptx));

        var ex = Assert.Throws<DocumentConversionException>(
            () => PresentationEditor.ReplaceImage(pptx, "{{chart}}", Png()));

        Assert.Contains("group", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SomethingThatIsNotAnImageIsRefused()
    {
        var pptx = PptxFixtures.DeckWithPlaceholderBox("{{chart}}");

        Assert.Throws<DocumentConversionException>(
            () => PresentationEditor.ReplaceImage(pptx, "{{chart}}", "not an image"u8.ToArray()));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void AMissingPlaceholderIsRejectedBeforeAnyWork(string? placeholder)
    {
        var pptx = PptxFixtures.DeckWithPlaceholderBox("{{c}}");

        Assert.ThrowsAny<ArgumentException>(
            () => PresentationEditor.ReplaceImage(pptx, placeholder!, Png()));
    }

    [Fact]
    public void AMissingImageIsRejectedBeforeAnyWork()
    {
        var pptx = PptxFixtures.DeckWithPlaceholderBox("{{c}}");

        Assert.Throws<ArgumentNullException>(
            () => PresentationEditor.ReplaceImage(pptx, "{{c}}", null!));
        Assert.Throws<ArgumentException>(
            () => PresentationEditor.ReplaceImage(pptx, "{{c}}", []));
    }

    // -----------------------------------------------------------------------------------------
    // Task 5: ReplaceImageAsync's Stream and file-path overloads.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public async Task ReplaceImageAsync_ThroughStreams_MatchesTheByteArrayOverload()
    {
        var pptx = PptxFixtures.DeckWithPlaceholderBox("{{chart}}");

        using var source = new MemoryStream(pptx, writable: false);
        using var destination = new MemoryStream();
        await PresentationEditor.ReplaceImageAsync(source, "{{chart}}", Png(), destination);

        // ExtractText would compare an empty list against an empty list here — the only shape
        // was replaced by a picture, which ExtractText does not walk — and so would prove nothing
        // about parity between the two paths. The picture's own xfrm does discriminate: a stream
        // path that fit the image differently, or picked the wrong shape, produces different
        // numbers rather than passing by coincidence.
        var viaBytes = PresentationEditor.ReplaceImage(pptx, "{{chart}}", Png());
        Assert.Equal(PictureTransformOf(viaBytes), PictureTransformOf(destination.ToArray()));

        using var ms = new MemoryStream(destination.ToArray());
        using var doc = PresentationDocument.Open(ms, false);
        Assert.Single(doc.PresentationPart!.SlideParts.Single().ImageParts);
    }

    /// <summary>The sole picture's offset and extent, for comparing two ReplaceImage results.</summary>
    private static (long X, long Y, long Cx, long Cy) PictureTransformOf(byte[] pptx)
    {
        using var ms = new MemoryStream(pptx);
        using var doc = PresentationDocument.Open(ms, false);
        var picture = doc.PresentationPart!.SlideParts.Single()
                         .Slide!.CommonSlideData!.ShapeTree!.Elements<P.Picture>().Single();
        var xfrm = picture.ShapeProperties!.Transform2D!;

        return (xfrm.Offset!.X!.Value, xfrm.Offset.Y!.Value, xfrm.Extents!.Cx!.Value, xfrm.Extents.Cy!.Value);
    }

    [Fact]
    public async Task ReplaceImageAsync_FromFileToFile_ReplacesThePlaceholder()
    {
        var pptx = PptxFixtures.DeckWithPlaceholderBox("{{chart}}");

        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, pptx);

        await PresentationEditor.ReplaceImageAsync(input.Path, output.Path, "{{chart}}", Png());

        using var ms = new MemoryStream(await File.ReadAllBytesAsync(output.Path));
        using var doc = PresentationDocument.Open(ms, false);
        Assert.Single(doc.PresentationPart!.SlideParts.Single().ImageParts);
    }

    // -----------------------------------------------------------------------------------------
    // A70: fixture pin. MultiLayoutDeck must genuinely produce two DIFFERENT layouts, not two
    // slides sharing one — otherwise the layout-selection tests in Task 5 would pass vacuously.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void MultiLayoutDeck_ProducesTwoSlidesWithGenuinelyDifferentLayouts()
    {
        var deck = PptxFixtures.MultiLayoutDeck("First", "Second");
        Assert.Empty(PptxFixtures.Validate(deck));

        using var ms = new MemoryStream(deck);
        using var doc = PresentationDocument.Open(ms, false);

        var slideParts = doc.PresentationPart!.Presentation!.SlideIdList!.Elements<P.SlideId>()
            .Select(id => (SlidePart)doc.PresentationPart!.GetPartById(id.RelationshipId!.Value!))
            .ToList();

        Assert.Equal(2, slideParts.Count);
        Assert.Equal("Title Slide", slideParts[0].SlideLayoutPart!.SlideLayout!.CommonSlideData!.Name!.Value);
        Assert.Equal("Second Layout", slideParts[1].SlideLayoutPart!.SlideLayout!.CommonSlideData!.Name!.Value);
        Assert.NotEqual(slideParts[0].SlideLayoutPart, slideParts[1].SlideLayoutPart);
    }

    // -----------------------------------------------------------------------------------------
    // A70: ReadSlide — per-slide text, same granularity as ExtractText.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void ReadSlide_ReturnsOnlyThatSlidesText()
    {
        var deck = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2", "Slide 3" }, reverseDeckOrder: false);

        Assert.Equal(new[] { "Slide 2" }, PresentationEditor.ReadSlide(deck, 2));
    }

    [Fact]
    public void ReadSlide_MatchesTheCorrespondingSliceOfExtractText()
    {
        var deck = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2", "Slide 3" }, reverseDeckOrder: false);

        var all = PresentationEditor.ExtractText(deck);
        for (var i = 1; i <= 3; i++)
        {
            Assert.Equal(new[] { all[i - 1] }, PresentationEditor.ReadSlide(deck, i));
        }
    }

    [Fact]
    public void ReadSlide_RejectsAnIndexBelowOne()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => PresentationEditor.ReadSlide(SampleDeck(), 0));
    }

    [Fact]
    public void ReadSlide_RejectsAnIndexPastTheLastSlide()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => PresentationEditor.ReadSlide(SampleDeck(), 2));
    }

    [Fact]
    public async Task ReadSlideAsync_FromFile_MatchesTheByteArrayOverload()
    {
        var pptx = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2" }, reverseDeckOrder: false);

        using var input = new TempFile();
        await File.WriteAllBytesAsync(input.Path, pptx);

        Assert.Equal(
            PresentationEditor.ReadSlide(pptx, 2),
            await PresentationEditor.ReadSlideAsync(input.Path, 2));
    }

    // -----------------------------------------------------------------------------------------
    // A70: RemoveSlides — an arbitrary set of indices, not a contiguous range.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void RemoveSlides_RemovesTheGivenIndicesAndKeepsTheRest()
    {
        var deck = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2", "Slide 3", "Slide 4" }, reverseDeckOrder: false);

        var edited = PresentationEditor.RemoveSlides(deck, new[] { 2, 4 });
        Assert.Empty(PptxFixtures.Validate(edited));

        Assert.Equal(2, PresentationEditor.SlideCount(edited));
        Assert.Equal(new[] { "Slide 1", "Slide 3" }, PresentationEditor.ExtractText(edited));
    }

    [Fact]
    public void RemoveSlides_AcceptsANonContiguousSetInAnyOrder()
    {
        var deck = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2", "Slide 3" }, reverseDeckOrder: false);

        var edited = PresentationEditor.RemoveSlides(deck, new[] { 3, 1 });
        Assert.Empty(PptxFixtures.Validate(edited));

        Assert.Equal(new[] { "Slide 2" }, PresentationEditor.ExtractText(edited));
    }

    [Fact]
    public void RemoveSlides_RejectsRemovingEverySlide()
    {
        var deck = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2" }, reverseDeckOrder: false);

        Assert.Throws<ArgumentOutOfRangeException>(
            () => PresentationEditor.RemoveSlides(deck, new[] { 1, 2 }));
    }

    [Fact]
    public void RemoveSlides_RejectsADuplicateIndex()
    {
        var deck = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2" }, reverseDeckOrder: false);

        Assert.Throws<ArgumentException>(() => PresentationEditor.RemoveSlides(deck, new[] { 1, 1 }));
    }

    [Fact]
    public void RemoveSlides_RejectsAnOutOfRangeIndex()
    {
        var deck = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2" }, reverseDeckOrder: false);

        Assert.Throws<ArgumentOutOfRangeException>(() => PresentationEditor.RemoveSlides(deck, new[] { 3 }));
    }

    [Fact]
    public async Task RemoveSlidesAsync_FromFileToFile_RemovesTheGivenIndices()
    {
        var pptx = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2", "Slide 3" }, reverseDeckOrder: false);

        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, pptx);

        await PresentationEditor.RemoveSlidesAsync(input.Path, output.Path, new[] { 2 });

        var outputBytes = await File.ReadAllBytesAsync(output.Path);
        Assert.Empty(PptxFixtures.Validate(outputBytes));

        var text = await PresentationEditor.ExtractTextAsync(output.Path);
        Assert.Equal(new[] { "Slide 1", "Slide 3" }, text);
    }

    [Fact]
    public void RemoveSlides_DeletesASlidesOwnNotesSlidePart_LeavingNoOrphan()
    {
        // A removed slide's OWN uniquely-referenced child part (speaker notes) must not survive
        // as an orphan in the output package. Verified by hand first, separately from this test:
        // DeletePart on a SlidePart DOES cascade to its own NotesSlidePart, confirmed against a
        // real build+run before writing this - so this pins correct behaviour rather than
        // testing for a defect that turned out not to exist.
        var deck = BuildDeckWhereSlideOneHasItsOwnNotesSlidePart();

        // Positive control: the notes part must actually be there before removal, or the
        // negative assertion below would pass vacuously.
        using (var beforeZip = new System.IO.Compression.ZipArchive(
            new MemoryStream(deck), System.IO.Compression.ZipArchiveMode.Read))
        {
            Assert.Contains(
                beforeZip.Entries, e => e.FullName.Contains("notesSlide", StringComparison.OrdinalIgnoreCase));
        }

        var edited = PresentationEditor.RemoveSlides(deck, new[] { 1 });
        Assert.Empty(PptxFixtures.Validate(edited));

        using var zip = new System.IO.Compression.ZipArchive(
            new MemoryStream(edited), System.IO.Compression.ZipArchiveMode.Read);
        Assert.DoesNotContain(
            zip.Entries, e => e.FullName.Contains("notesSlide", StringComparison.OrdinalIgnoreCase));
    }

    private static byte[] BuildDeckWhereSlideOneHasItsOwnNotesSlidePart()
    {
        var deck = PptxFixtures.MultiSlideDeck(new[] { "Slide 1", "Slide 2" }, reverseDeckOrder: false);

        using var ms = new MemoryStream();
        ms.Write(deck, 0, deck.Length);
        ms.Position = 0;

        using (var doc = PresentationDocument.Open(ms, true))
        {
            var slide1 = doc.PresentationPart!.SlideParts.First();
            var notesPart = slide1.AddNewPart<NotesSlidePart>();
            notesPart.NotesSlide = new P.NotesSlide(new P.CommonSlideData(new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties())));
            notesPart.NotesSlide.Save();
        }

        return ms.ToArray();
    }

    // -----------------------------------------------------------------------------------------
    // A70: ReorderSlides — a full permutation, matching PdfEditor.ReorderPages exactly.
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void ReorderSlides_AppliesTheGivenPermutation()
    {
        var deck = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2", "Slide 3" }, reverseDeckOrder: false);

        var edited = PresentationEditor.ReorderSlides(deck, new[] { 3, 1, 2 });
        Assert.Empty(PptxFixtures.Validate(edited));

        Assert.Equal(new[] { "Slide 3", "Slide 1", "Slide 2" }, PresentationEditor.ExtractText(edited));
    }

    [Fact]
    public void ReorderSlides_RejectsAnOrderMissingASlide()
    {
        var deck = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2", "Slide 3" }, reverseDeckOrder: false);

        Assert.Throws<ArgumentException>(() => PresentationEditor.ReorderSlides(deck, new[] { 1, 2 }));
    }

    [Fact]
    public void ReorderSlides_RejectsAnOrderWithADuplicate()
    {
        var deck = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2", "Slide 3" }, reverseDeckOrder: false);

        Assert.Throws<ArgumentException>(() => PresentationEditor.ReorderSlides(deck, new[] { 1, 1, 2 }));
    }

    [Fact]
    public async Task ReorderSlidesAsync_FromFileToFile_AppliesThePermutation()
    {
        var pptx = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2" }, reverseDeckOrder: false);

        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, pptx);

        await PresentationEditor.ReorderSlidesAsync(input.Path, output.Path, new[] { 2, 1 });

        var text = await PresentationEditor.ExtractTextAsync(output.Path);
        Assert.Equal(new[] { "Slide 2", "Slide 1" }, text);
    }

    // -----------------------------------------------------------------------------------------
    // A70: InsertSlides — new content via PptxSlide, matching PdfEditor.InsertPages' atIndex
    // convention (SlideCount + 1 appends).
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void InsertSlides_AtOne_PutsTheNewSlidesInFront()
    {
        var deck = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2" }, reverseDeckOrder: false);

        var edited = PresentationEditor.InsertSlides(deck, 1, new[] { PptxSlide.Titled("New") });
        Assert.Empty(PptxFixtures.Validate(edited));

        Assert.Equal(3, PresentationEditor.SlideCount(edited));
        Assert.Contains("New", PresentationEditor.ReadSlide(edited, 1)[0]);
    }

    [Fact]
    public void InsertSlides_AtSlideCountPlusOne_Appends()
    {
        var deck = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2" }, reverseDeckOrder: false);

        var edited = PresentationEditor.InsertSlides(deck, 3, new[] { PptxSlide.Titled("New") });
        Assert.Empty(PptxFixtures.Validate(edited));

        Assert.Equal(3, PresentationEditor.SlideCount(edited));
        Assert.Contains("New", PresentationEditor.ReadSlide(edited, 3)[0]);
    }

    [Fact]
    public void InsertSlides_InTheMiddle_PreservesTheSurroundingOrder()
    {
        var deck = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2", "Slide 3" }, reverseDeckOrder: false);

        var edited = PresentationEditor.InsertSlides(deck, 2, new[] { PptxSlide.Titled("New") });

        var titles = new[] { 1, 2, 3, 4 }.Select(i => PresentationEditor.ReadSlide(edited, i)[0]);
        Assert.Equal(new[] { "Slide 1", "New", "Slide 2", "Slide 3" }, titles);
    }

    [Fact]
    public void InsertSlides_AttachesToTheLayoutOfTheSlideBeforeTheInsertionPoint()
    {
        var deck = PptxFixtures.MultiLayoutDeck("First", "Second");

        var edited = PresentationEditor.InsertSlides(deck, 2, new[] { PptxSlide.Titled("New") });
        Assert.Empty(PptxFixtures.Validate(edited));

        using var ms = new MemoryStream(edited);
        using var doc = PresentationDocument.Open(ms, false);
        var insertedSlide = doc.PresentationPart!.SlideParts
            .First(p => p.Slide!.Descendants<A.Text>().Any(t => t.Text == "New"));

        Assert.Equal(
            "Title Slide",
            insertedSlide.SlideLayoutPart!.SlideLayout!.CommonSlideData!.Name!.Value);
    }

    [Fact]
    public void InsertSlides_AfterTheSecondSlide_AttachesToTheSecondSlidesLayoutNotTheFirst()
    {
        // The discriminating case: insert AFTER "Second" (layout B), so the new slide must
        // attach to layout B, not the deck's first layout (layout A) - proving the rule is
        // genuinely "the adjacent slide", not "always slide one".
        var deck = PptxFixtures.MultiLayoutDeck("First", "Second");

        var edited = PresentationEditor.InsertSlides(deck, 3, new[] { PptxSlide.Titled("New") });

        using var ms = new MemoryStream(edited);
        using var doc = PresentationDocument.Open(ms, false);
        var insertedSlide = doc.PresentationPart!.SlideParts
            .First(p => p.Slide!.Descendants<A.Text>().Any(t => t.Text == "New"));

        Assert.Equal(
            "Second Layout",
            insertedSlide.SlideLayoutPart!.SlideLayout!.CommonSlideData!.Name!.Value);
    }

    [Fact]
    public void InsertSlides_IntoAnEmptyDeck_UsesTheFirstMastersFirstLayout()
    {
        var empty = PresentationEditor.Create(Array.Empty<PptxSlide>());

        var edited = PresentationEditor.InsertSlides(empty, 1, new[] { PptxSlide.Titled("New") });
        Assert.Empty(PptxFixtures.Validate(edited));

        Assert.Equal(1, PresentationEditor.SlideCount(edited));
        Assert.Contains("New", PresentationEditor.ReadSlide(edited, 1)[0]);

        using var ms = new MemoryStream(edited);
        using var doc = PresentationDocument.Open(ms, false);
        var insertedSlide = doc.PresentationPart!.SlideParts.Single();
        Assert.NotNull(insertedSlide.SlideLayoutPart);
    }

    [Fact]
    public void InsertSlides_ABatch_KeepsThemInTheOrderGiven()
    {
        var deck = PptxFixtures.MultiSlideDeck(new[] { "Slide 1" }, reverseDeckOrder: false);

        var edited = PresentationEditor.InsertSlides(
            deck, 1, new[] { PptxSlide.Titled("A"), PptxSlide.Titled("B") });

        var titles = new[] { 1, 2, 3 }.Select(i => PresentationEditor.ReadSlide(edited, i)[0]);
        Assert.Equal(new[] { "A", "B", "Slide 1" }, titles);
    }

    [Fact]
    public void InsertSlides_RejectsAnIndexPastTheAppendPosition()
    {
        var deck = PptxFixtures.MultiSlideDeck(new[] { "Slide 1" }, reverseDeckOrder: false);

        Assert.Throws<ArgumentOutOfRangeException>(
            () => PresentationEditor.InsertSlides(deck, 3, new[] { PptxSlide.Titled("New") }));
    }

    [Fact]
    public void InsertSlides_RejectsAnIndexBelowOne()
    {
        var deck = PptxFixtures.MultiSlideDeck(new[] { "Slide 1" }, reverseDeckOrder: false);

        Assert.Throws<ArgumentOutOfRangeException>(
            () => PresentationEditor.InsertSlides(deck, 0, new[] { PptxSlide.Titled("New") }));
    }

    [Fact]
    public void InsertSlides_SlideIdsStayUniqueEvenWhenExistingIdsAreNotContiguous()
    {
        // Simulates a deck that has already had a slide removed and re-added: existing ids are
        // not a tidy 256, 257, 258, ... run. The new slide's id must still not collide.
        var deck = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2", "Slide 3" }, reverseDeckOrder: false);
        deck = PresentationEditor.RemoveSlides(deck, new[] { 2 });
        deck = PresentationEditor.InsertSlides(deck, 2, new[] { PptxSlide.Titled("Replacement") });

        Assert.Empty(PptxFixtures.Validate(deck));
        Assert.Equal(3, PresentationEditor.SlideCount(deck));
    }

    [Fact]
    public async Task InsertSlidesAsync_FromFileToFile_InsertsAtTheGivenPosition()
    {
        var pptx = PptxFixtures.MultiSlideDeck(
            new[] { "Slide 1", "Slide 2" }, reverseDeckOrder: false);

        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, pptx);

        await PresentationEditor.InsertSlidesAsync(
            input.Path, output.Path, 2, new[] { PptxSlide.Titled("New") });

        var titles = await PresentationEditor.ExtractTextAsync(output.Path);
        Assert.Equal(new[] { "Slide 1", "New", "Slide 2" }, titles);
    }

    [Fact]
    public void InsertSlides_ScalesContentToFitADeckOfADifferentSize()
    {
        // sample.pptx (and every other fixture) is 16:9, the same canvas BuildSlide's hard-coded
        // geometry assumes - so inserting into a DIFFERENTLY-SIZED deck is the one case that can
        // silently overhang the canvas edge without any existing test noticing. Built by hand
        // rather than from a fixture: a deck whose width AND height both differ from the 16:9
        // design size (12192000 x 6858000), using the sample's own layout/master unchanged - a
        // deck that only varied one axis could not catch the two scale factors being swapped.
        var deck = PptxFixtures.Sample();
        using var ms = new MemoryStream();
        ms.Write(deck, 0, deck.Length);
        ms.Position = 0;
        using (var doc = PresentationDocument.Open(ms, true))
        {
            var slideSize = doc.PresentationPart!.Presentation!.SlideSize!;
            slideSize.Cx = 9144000; // narrower than the 16:9 design width (12192000)
            slideSize.Cy = 5143500; // shorter than the 16:9 design height (6858000)
            doc.PresentationPart.Presentation.Save();
        }
        var differentlySizedDeck = ms.ToArray();

        var edited = PresentationEditor.InsertSlides(differentlySizedDeck, 2, new[] { PptxSlide.Titled("New") });
        Assert.Empty(PptxFixtures.Validate(edited));

        using var checkMs = new MemoryStream(edited);
        using var checkDoc = PresentationDocument.Open(checkMs, false);
        var insertedSlide = checkDoc.PresentationPart!.SlideParts
            .First(p => p.Slide!.Descendants<A.Text>().Any(t => t.Text == "New"));
        var titleXfrm = insertedSlide.Slide!.Descendants<A.Transform2D>().First();

        var rightEdge = titleXfrm.Offset!.X!.Value + titleXfrm.Extents!.Cx!.Value;
        var bottomEdge = titleXfrm.Offset.Y!.Value + titleXfrm.Extents.Cy!.Value;
        Assert.True(rightEdge <= 9144000,
            $"Inserted title shape's right edge ({rightEdge}) overhangs the deck's width (9144000).");
        Assert.True(bottomEdge <= 5143500,
            $"Inserted title shape's bottom edge ({bottomEdge}) overhangs the deck's height (5143500).");
    }

    [Fact]
    public void InsertSlides_WhenTheTargetLayoutHasAMatchingPlaceholder_InheritsItsPosition()
    {
        // The layout's title/body placeholders are relocated FAR from PptxDocumentWriter's own
        // constants (TitleXEmu=838200/TitleYEmu=365125, BodyXEmu=838200/BodyYEmu=1825625) while
        // keeping their TYPES identical (title, body idx=1) -- the one case the fix is meant to
        // improve. Measured via PdfProbe against the real render pipeline, not merely reading the
        // OOXML back: text presence and position, not just structure.
        var deck = PptxFixtures.DeckWithRelocatedLayoutPlaceholders();

        var edited = PresentationEditor.InsertSlides(
            deck, 2, new[] { PptxSlide.Titled("Inserted Title", "Bullet A") });
        Assert.Empty(PptxFixtures.Validate(edited));

        var pdf = PptxToPdfConverter.Convert(edited);
        var text = PdfProbe.ExtractText(pdf);
        Assert.Contains("Inserted Title", text);
        Assert.Contains("Bullet A", text);

        var yPositions = PdfProbe.TextYPositions(pdf);

        // TextYPositions returns values in DOCUMENT order, not grouped or filterable by slide -- a
        // value-range heuristic (e.g. "near the top"/"near the bottom") would risk matching the
        // WRONG slide's text, since slide 1 ("First"/"One", built by Create with its OWN unrelated
        // explicit geometry, unaffected by relocating the shared layout) also contributes two
        // entries to the same pooled list. Index into it instead: slide 1 contributes exactly 2
        // entries (title, one bullet), so the inserted slide's title is index 2 and its bullet is
        // index 3 -- deterministic from the fixture's own known shape counts, not a guess about
        // absolute position.
        Assert.Equal(4, yPositions.Count);
        var insertedTitleY = yPositions[2];
        var insertedBulletY = yPositions[3];

        // The relocated layout puts the title near the BOTTOM of the slide (off.Y=5000000 of
        // 6858000 EMU tall) and the body near the TOP (off.Y=500000) -- inverted from
        // PptxDocumentWriter's own near-top-title/mid-page-body constants. In PDF points (bottom-up
        // origin), a title inheriting the relocated position lands LOW; a body inheriting it lands
        // HIGH. If the fix instead kept the fixed geometry, both would land at their usual
        // (opposite) relative order instead.
        Assert.True(insertedTitleY < insertedBulletY,
            $"Expected the inherited title (relocated near the bottom, low Y) to sit below the " +
            $"inherited bullet (relocated near the top, high Y) in PDF coordinates. " +
            $"Title Y={insertedTitleY}, Bullet Y={insertedBulletY}.");
    }

    [Fact]
    public void InsertSlides_WhenTheTargetLayoutHasNoMatchingPlaceholder_StillRendersTheContent()
    {
        // The measured failure mode this test pins: sample.pptx's real "Title Slide" layout uses
        // ctrTitle/subTitle types, which do not match what BuildSlide writes (title/body). Omitting
        // a:xfrm unconditionally would make this content vanish from the render entirely -- schema-
        // valid, but invisible. This proves the fallback keeps it visible, exactly like today.
        var deck = PptxFixtures.Sample();

        var edited = PresentationEditor.InsertSlides(
            deck, 2, new[] { PptxSlide.Titled("Inserted Title", "Bullet A") });
        Assert.Empty(PptxFixtures.Validate(edited));

        var pdf = PptxToPdfConverter.Convert(edited);
        var text = PdfProbe.ExtractText(pdf);
        Assert.Contains("Inserted Title", text);
        Assert.Contains("Bullet A", text);
    }

    [Fact]
    public void InsertSlides_WhenTheTargetLayoutHasNoMatchingPlaceholder_KeepsTheFixedGeometry()
    {
        // The structural twin of the render-based test above: confirms the SHAPE still carries its
        // own explicit a:xfrm (today's behavior, unchanged) rather than inheriting nothing.
        var deck = PptxFixtures.Sample();

        var edited = PresentationEditor.InsertSlides(
            deck, 2, new[] { PptxSlide.Titled("Inserted Title", "Bullet A") });

        using var ms = new MemoryStream(edited);
        using var doc = PresentationDocument.Open(ms, false);
        var insertedSlide = doc.PresentationPart!.SlideParts
            .First(p => p.Slide!.Descendants<A.Text>().Any(t => t.Text == "Inserted Title"));

        var xfrms = insertedSlide.Slide!.Descendants<A.Transform2D>().ToList();
        Assert.Equal(2, xfrms.Count); // title AND body, both unmatched, both keep explicit geometry
    }

    [Fact]
    public void InsertSlides_WhenTheTargetLayoutHasAMatchingPlaceholder_RemovesTheExplicitGeometry()
    {
        // The structural twin of the render-based match test: confirms the shape's a:xfrm is
        // genuinely GONE, not merely coincidentally equal to the layout's position.
        var deck = PptxFixtures.DeckWithRelocatedLayoutPlaceholders();

        var edited = PresentationEditor.InsertSlides(
            deck, 2, new[] { PptxSlide.Titled("Inserted Title", "Bullet A") });

        using var ms = new MemoryStream(edited);
        using var doc = PresentationDocument.Open(ms, false);
        var insertedSlide = doc.PresentationPart!.SlideParts
            .First(p => p.Slide!.Descendants<A.Text>().Any(t => t.Text == "Inserted Title"));

        var xfrms = insertedSlide.Slide!.Descendants<A.Transform2D>().ToList();
        Assert.Empty(xfrms); // both shapes matched, both inherit -- no explicit geometry left
    }

    // ---------------------------------------------------------------------------------------
    // The THIRD case: the layout names the role but positions nothing. A role match alone used
    // to be enough to strip the inserted shape's own a:xfrm, and these two real stock layouts
    // are the ones that then had nowhere to draw -- this render pipeline resolves slide ->
    // layout, never layout -> master, so nothing downstream supplies the missing box.
    //
    // Measured through PptxToPdfConverter + PdfProbe against the pre-fix code, per layout:
    //
    //   "Title and Content"        1 a:xfrm kept, page 2 rendered "* Bullet A"  -- title GONE
    //   "Title and Vertical Text"  0 a:xfrm kept, page 2 rendered EMPTY         -- both GONE
    //
    // and after the fix, 2 kept and "Inserted Title" + "Bullet A" on page 2 for both. The two
    // cases differ and both are worth pinning: "Title and Content" (what PowerPoint gives a new
    // body slide by default) has an UNTYPED idx=1 body placeholder, so its body already fell
    // back on the type check and only the title exercises the geometry guard; "Title and
    // Vertical Text" types both, so both do.
    // ---------------------------------------------------------------------------------------

    [Theory]
    [InlineData("Title and Content")]
    [InlineData("Title and Vertical Text")]
    public void InsertSlides_WhenTheTargetLayoutNamesTheRoleButPositionsNothing_StillRendersTheContent(
        string layoutName)
    {
        var deck = PptxFixtures.SampleAttachedToLayout(layoutName);
        Assert.Empty(PptxFixtures.Validate(deck)); // the re-pointed fixture is itself schema-valid

        var edited = PresentationEditor.InsertSlides(
            deck, 2, new[] { PptxSlide.Titled("Inserted Title", "Bullet A") });
        Assert.Empty(PptxFixtures.Validate(edited));

        var pdf = PptxToPdfConverter.Convert(edited);
        var text = PdfProbe.ExtractText(pdf);
        Assert.Contains("Inserted Title", text);
        Assert.Contains("Bullet A", text);
    }

    [Theory]
    [InlineData("Title and Content")]
    [InlineData("Title and Vertical Text")]
    public void InsertSlides_WhenTheTargetLayoutNamesTheRoleButPositionsNothing_KeepsTheFixedGeometry(
        string layoutName)
    {
        // The structural twin of the render test above. It discriminates where the render test
        // cannot: a render only says the text appeared, while this says WHICH box put it there.
        var deck = PptxFixtures.SampleAttachedToLayout(layoutName);

        var edited = PresentationEditor.InsertSlides(
            deck, 2, new[] { PptxSlide.Titled("Inserted Title", "Bullet A") });

        using var ms = new MemoryStream(edited);
        using var doc = PresentationDocument.Open(ms, false);
        var insertedSlide = doc.PresentationPart!.SlideParts
            .First(p => p.Slide!.Descendants<A.Text>().Any(t => t.Text == "Inserted Title"));

        var xfrms = insertedSlide.Slide!.Descendants<A.Transform2D>().ToList();
        Assert.Equal(2, xfrms.Count); // title AND body keep their own box -- nothing to inherit
    }

    [Fact]
    public void InsertSlides_WhenARealStockLayoutPositionsItsPlaceholders_StillInheritsFromIt()
    {
        // The positive control for the geometry guard, and the reason it is not simply a way of
        // switching the feature off: "Section Header" is a real PowerPoint-authored layout in the
        // SAME package as the two above, with title/body idx=1 placeholders that DO carry their own
        // a:xfrm. Inheritance still happens here -- 0 a:xfrm on the inserted slide -- and the
        // content still renders. Without this, the two tests above would pass just as happily
        // against a guard that never inherited from anything.
        var deck = PptxFixtures.SampleAttachedToLayout("Section Header");

        var edited = PresentationEditor.InsertSlides(
            deck, 2, new[] { PptxSlide.Titled("Inserted Title", "Bullet A") });
        Assert.Empty(PptxFixtures.Validate(edited));

        using (var ms = new MemoryStream(edited))
        using (var doc = PresentationDocument.Open(ms, false))
        {
            var insertedSlide = doc.PresentationPart!.SlideParts
                .First(p => p.Slide!.Descendants<A.Text>().Any(t => t.Text == "Inserted Title"));
            Assert.Empty(insertedSlide.Slide!.Descendants<A.Transform2D>());
        }

        var text = PdfProbe.ExtractText(PptxToPdfConverter.Convert(edited));
        Assert.Contains("Inserted Title", text);
        Assert.Contains("Bullet A", text);
    }

    // ---------------------------------------------------------------------------------------
    // Re-review fix round 2, Finding 1: a present Transform2D is not necessarily a USABLE box.
    // CT_Transform2D declares both a:off and a:ext as optional, so a layout placeholder can carry
    // an a:xfrm with only a:off, or only a:ext. Measured against the pre-fix guard
    // (Transform2D-presence only): stripping just <a:ext> from "Section Header"'s title
    // placeholder -- the SAME real, PowerPoint-authored layout used as the positive control above
    // -- still passed it, still stripped the inserted title's own a:xfrm, and "Inserted Title"
    // vanished from the render entirely. That is the identical failure class the whole geometry
    // guard exists to prevent.
    //
    // Re-review fix round 3, Finding 2: only the a:ext-missing half of that was ever exercised --
    // nothing failed if the "layoutXfrm?.Offset is null" half of the guard's condition were
    // deleted. Both cases below are now [Theory] cases over XfrmPart, so both halves
    // of the completeness check are pinned.
    // ---------------------------------------------------------------------------------------

    [Theory]
    [InlineData(XfrmPart.Extents)]
    [InlineData(XfrmPart.Offset)]
    public void InsertSlides_WhenTheLayoutTitleXfrmIsIncomplete_StillRendersTheTitle(
        XfrmPart missingPart)
    {
        var deck = PptxFixtures.SampleAttachedToLayoutWithAnIncompleteTitleBox("Section Header", missingPart);
        Assert.Empty(PptxFixtures.Validate(deck)); // the mutated fixture is itself schema-valid

        var edited = PresentationEditor.InsertSlides(
            deck, 2, new[] { PptxSlide.Titled("Inserted Title", "Bullet A") });
        Assert.Empty(PptxFixtures.Validate(edited));

        var text = PdfProbe.ExtractText(PptxToPdfConverter.Convert(edited));
        Assert.Contains("Inserted Title", text);
        Assert.Contains("Bullet A", text);
    }

    [Theory]
    [InlineData(XfrmPart.Extents)]
    [InlineData(XfrmPart.Offset)]
    public void InsertSlides_WhenTheLayoutTitleXfrmIsIncomplete_TitleFallsBackButBodyStillInherits(
        XfrmPart missingPart)
    {
        // The structural twin of the render test above, and it discriminates PER PLACEHOLDER: only
        // the title's box was made incomplete, so only the title keeps its own fixed geometry -- the
        // body placeholder's box is still complete and is still inherited, exactly like the
        // positive control. That split is what proves the guard checks completeness placeholder by
        // placeholder rather than merely refusing the whole layout the moment one box is bad.
        var deck = PptxFixtures.SampleAttachedToLayoutWithAnIncompleteTitleBox("Section Header", missingPart);

        var edited = PresentationEditor.InsertSlides(
            deck, 2, new[] { PptxSlide.Titled("Inserted Title", "Bullet A") });

        using var ms = new MemoryStream(edited);
        using var doc = PresentationDocument.Open(ms, false);
        var insertedSlide = doc.PresentationPart!.SlideParts
            .First(p => p.Slide!.Descendants<A.Text>().Any(t => t.Text == "Inserted Title"));

        var xfrms = insertedSlide.Slide!.Descendants<A.Transform2D>().ToList();
        Assert.Single(xfrms); // title keeps its own box (incomplete layout box); body still inherits
    }

    // ---------------------------------------------------------------------------------------
    // Re-review fix round 3, Finding 1: the match walked into GROUPED shapes on the layout via
    // Descendants<P.Shape>(), but this repo's render pipeline (PptxToPdfConverter/OfficeIMO) only
    // resolves a layout's TOP-LEVEL shape tree. Measured against the pre-fix guard: wrapping
    // "Section Header"'s title placeholder in a p:grpSp -- matching type, complete a:xfrm,
    // schema-valid before and after -- still matched, still stripped the inserted title's own
    // a:xfrm, and the title vanished from the render. Same failure class as the two findings
    // above, one level further in -- and the same class of mistake CLAUDE.md already records for
    // DocxEditor (w:txbxContent) and TableRowFinder (nested tables).
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void InsertSlides_WhenTheMatchingLayoutPlaceholderIsInsideAGroup_DoesNotInheritFromIt()
    {
        var deck = PptxFixtures.SampleAttachedToLayoutWithTitleInGroup("Section Header");
        Assert.Empty(PptxFixtures.Validate(deck)); // the mutated fixture is itself schema-valid

        var edited = PresentationEditor.InsertSlides(
            deck, 2, new[] { PptxSlide.Titled("Inserted Title", "Bullet A") });
        Assert.Empty(PptxFixtures.Validate(edited));

        using (var ms = new MemoryStream(edited))
        using (var doc = PresentationDocument.Open(ms, false))
        {
            var insertedSlide = doc.PresentationPart!.SlideParts
                .First(p => p.Slide!.Descendants<A.Text>().Any(t => t.Text == "Inserted Title"));

            // The title's layout match is inside a group and must not count -- it keeps its own
            // explicit geometry (ScaleToFitDeck fallback). The body has no such wrinkle and still
            // inherits, exactly like the positive control -- proving this is per-placeholder, not
            // a refusal of the whole layout.
            var xfrms = insertedSlide.Slide!.Descendants<A.Transform2D>().ToList();
            Assert.Single(xfrms);
        }

        // And it renders correctly via the fallback -- proving the fallback works, not merely that
        // the structural (non-)match is correct.
        var text = PdfProbe.ExtractText(PptxToPdfConverter.Convert(edited));
        Assert.Contains("Inserted Title", text);
        Assert.Contains("Bullet A", text);
    }

    // ---------------------------------------------------------------------------------------
    // Re-review fix round 2, Finding 2: the null-SlideLayout guard in
    // LayoutHasMatchingPositionedPlaceholder had no test proving it fires. Before this branch,
    // InsertSlides never read a layout's XML content at all, so a SlideLayoutPart with no root
    // element -- a genuinely reachable OOXML state -- was not a failure mode InsertSlides could
    // reach. Reading a layout's placeholders to decide on geometry inheritance made it reachable,
    // and without a test nothing fails if the guard is deleted.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void InsertSlides_WhenTheTargetLayoutHasNoRootElement_DegradesToTheFixedGeometryRatherThanThrowing()
    {
        var deck = PptxFixtures.SampleAttachedToAnUnassignedLayout();

        var edited = PresentationEditor.InsertSlides(
            deck, 2, new[] { PptxSlide.Titled("Inserted Title", "Bullet A") });

        using var ms = new MemoryStream(edited);
        using var doc = PresentationDocument.Open(ms, false);
        var insertedSlide = doc.PresentationPart!.SlideParts
            .First(p => p.Slide!.Descendants<A.Text>().Any(t => t.Text == "Inserted Title"));

        var xfrms = insertedSlide.Slide!.Descendants<A.Transform2D>().ToList();
        Assert.Equal(2, xfrms.Count); // title AND body degrade to the fixed-geometry fallback
    }

    [Fact]
    public void AddChart_AddsAChart_ThatSurvivesThePptxRoundTrip()
    {
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Slide 1") });
        var data = new ChartData(
            new[] { "North", "South" },
            new[] { new ChartSeries("Total", new double[] { 1200, 980 }) });

        var result = PresentationEditor.AddChart(
            pptx, 1, ChartType.ColumnClustered, data, title: "Regional Totals");

        using var source = new MemoryStream(result, writable: false);
        using var doc = OfficeIMO.PowerPoint.PowerPointPresentation.Load(source);
        var slide = doc.Slides[0];
        Assert.Single(slide.Charts);
        Assert.Equal("Regional Totals", slide.Charts.Single().Title);
    }

    [Fact]
    public void AddChart_TheChartAndItsTitleSurviveThePptxToPdfRender()
    {
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Slide 1") });
        var data = new ChartData(
            new[] { "North", "South" },
            new[] { new ChartSeries("Total", new double[] { 1200, 980 }) });

        var result = PresentationEditor.AddChart(
            pptx, 1, ChartType.ColumnClustered, data, title: "Regional Totals");

        var text = PdfProbe.ExtractText(PptxToPdfConverter.Convert(result));
        Assert.Contains("Regional Totals", text);
        Assert.Contains("North", text);
        Assert.Contains("South", text);
    }

    [Fact]
    public void AddChart_SlideIndexBelowOne_Throws()
    {
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Slide 1") });
        var data = new ChartData(new[] { "A" }, new[] { new ChartSeries("S", new double[] { 1 }) });

        Assert.Throws<ArgumentOutOfRangeException>(
            () => PresentationEditor.AddChart(pptx, 0, ChartType.Line, data));
    }

    [Fact]
    public void AddChart_SlideIndexAboveSlideCount_Throws()
    {
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Slide 1") });
        var data = new ChartData(new[] { "A" }, new[] { new ChartSeries("S", new double[] { 1 }) });

        Assert.Throws<ArgumentOutOfRangeException>(
            () => PresentationEditor.AddChart(pptx, 2, ChartType.Line, data));
    }

    [Fact]
    public void AddChart_NullData_Throws()
    {
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Slide 1") });

        Assert.Throws<ArgumentNullException>(
            () => PresentationEditor.AddChart(pptx, 1, ChartType.Line, null!));
    }

    [Fact]
    public async Task AddChartAsync_FromStream_MatchesTheByteArrayOverload()
    {
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Slide 1") });
        var data = new ChartData(new[] { "A" }, new[] { new ChartSeries("S", new double[] { 1 }) });
        var expected = PresentationEditor.AddChart(
            pptx, 1, ChartType.Line, data, title: "Regional Totals");

        using var source = new MemoryStream(pptx, writable: false);
        var actual = await PresentationEditor.AddChartAsync(
            source, 1, ChartType.Line, data, title: "Regional Totals");

        // NOT Assert.Equal(expected, actual) on the raw bytes. OfficeIMO.PowerPoint's Save() mints
        // fresh internal ids on every independent call -- the same nondeterminism PresentationEditor
        // .Create's own doc comment already records for the raw OpenXml SDK path ("two calls with
        // identical slides in the same process differ [in bytes]"), measured here to hold for
        // OfficeIMO's writer too: two independent AddChartPoints+Save calls from byte-identical
        // input produced outputs of different LENGTH. So the two overloads are compared structurally
        // -- both produce a deck with the same slide/chart shape -- rather than byte-for-byte.
        using var expectedSource = new MemoryStream(expected, writable: false);
        using var actualSource = new MemoryStream(actual, writable: false);
        using var expectedDoc = OfficeIMO.PowerPoint.PowerPointPresentation.Load(expectedSource);
        using var actualDoc = OfficeIMO.PowerPoint.PowerPointPresentation.Load(actualSource);
        Assert.Equal(expectedDoc.Slides.Count, actualDoc.Slides.Count);
        var expectedChart = Assert.Single(expectedDoc.Slides[0].Charts);
        var actualChart = Assert.Single(actualDoc.Slides[0].Charts);

        // The slide/chart-count checks above only prove "exactly one chart landed on each side" --
        // they say nothing about whether it is the RIGHT chart. TryGetOfficeSnapshot exposes the
        // chart's type/title/data without relying on Save()'s non-deterministic byte output (see
        // above), so compare those directly between the two independently-produced decks.
        Assert.True(expectedChart.TryGetOfficeSnapshot(out var expectedSnapshot));
        Assert.True(actualChart.TryGetOfficeSnapshot(out var actualSnapshot));
        Assert.Equal(expectedSnapshot.ChartKind, actualSnapshot.ChartKind);
        Assert.Equal(expectedSnapshot.Title, actualSnapshot.Title);
        Assert.Equal(expectedSnapshot.Data.Categories, actualSnapshot.Data.Categories);
        Assert.Equal(expectedSnapshot.Data.Series.Count, actualSnapshot.Data.Series.Count);
        for (var i = 0; i < expectedSnapshot.Data.Series.Count; i++)
        {
            Assert.Equal(expectedSnapshot.Data.Series[i].Name, actualSnapshot.Data.Series[i].Name);
            Assert.Equal(expectedSnapshot.Data.Series[i].Values, actualSnapshot.Data.Series[i].Values);
        }
    }

    [Fact]
    public async Task AddChartAsync_FromFile_MatchesTheByteArrayOverload()
    {
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Slide 1") });
        var data = new ChartData(new[] { "A" }, new[] { new ChartSeries("S", new double[] { 1 }) });

        using var input = new TempFile();
        await File.WriteAllBytesAsync(input.Path, pptx);

        var expected = PresentationEditor.AddChart(
            pptx, 1, ChartType.Line, data, title: "Regional Totals");
        var actual = await PresentationEditor.AddChartAsync(
            input.Path, 1, ChartType.Line, data, title: "Regional Totals");

        // Streams hoisted so they are disposed, and declared BEFORE the documents so `using`
        // disposes them after — a presentation outliving its own source stream would break.
        using var expectedStream = new MemoryStream(expected, writable: false);
        using var actualStream = new MemoryStream(actual, writable: false);
        using var expectedDoc = OfficeIMO.PowerPoint.PowerPointPresentation.Load(expectedStream);
        using var actualDoc = OfficeIMO.PowerPoint.PowerPointPresentation.Load(actualStream);
        var expectedChart = Assert.Single(expectedDoc.Slides[0].Charts);
        var actualChart = Assert.Single(actualDoc.Slides[0].Charts);
        Assert.True(expectedChart.TryGetOfficeSnapshot(out var expectedSnapshot));
        Assert.True(actualChart.TryGetOfficeSnapshot(out var actualSnapshot));
        Assert.Equal(expectedSnapshot.ChartKind, actualSnapshot.ChartKind);
        Assert.Equal(expectedSnapshot.Title, actualSnapshot.Title);
    }

    public static IEnumerable<object[]> AllChartTypes() =>
        Enum.GetValues<ChartType>().Select(t => new object[] { t });

    [Theory]
    [MemberData(nameof(AllChartTypes))]
    public void AddChart_EveryChartTypeProducesALoadableChart(ChartType type)
    {
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Slide 1") });
        var data = new ChartData(
            new[] { "North", "South" },
            new[] { new ChartSeries("Total", new double[] { 1200, 980 }) });

        var result = PresentationEditor.AddChart(pptx, 1, type, data);

        using var source = new MemoryStream(result, writable: false);
        using var doc = OfficeIMO.PowerPoint.PowerPointPresentation.Load(source);
        var chart = Assert.Single(doc.Slides[0].Charts);
        Assert.NotNull(chart);
    }

    [Fact]
    public void EveryOperation_PreservesALinkedOleObject()
    {
        var basePptx = PresentationEditor.Create(new[]
        {
            PptxSlide.Titled("Slide 1", "{{name}}"),
            PptxSlide.Titled("Slide 2", "bullet"),
        });

        byte[] withOle;
        using (var ms = new MemoryStream())
        {
            ms.Write(basePptx, 0, basePptx.Length);
            ms.Position = 0;
            using (var doc = OfficeIMO.PowerPoint.PowerPointPresentation.Load(ms))
            {
                // AddOleObject requires a real OLE compound-document blob, validated by magic
                // bytes - a LINKED object exercises the identical PowerPointOleObject and
                // relationship machinery without needing to hand-author a CFBF file.
                doc.Slides[0].AddLinkedOleObject(
                    new Uri("file:///C:/does-not-need-to-exist/workbook.xlsx"),
                    "Excel.Sheet.12", false, 100, 100, 500000, 500000);
                doc.Save();
            }
            withOle = ms.ToArray();
        }

        static void AssertOleSurvivesOnExactlyOneSlide(byte[] pptx, int expectedSlideIndex)
        {
            using var ms = new MemoryStream(pptx, writable: false);
            using var doc = OfficeIMO.PowerPoint.PowerPointPresentation.Load(ms);
            for (var i = 0; i < doc.Slides.Count; i++)
            {
                var oleObjects = doc.Slides[i].OleObjects.ToList();
                if (i == expectedSlideIndex)
                {
                    var ole = Assert.Single(oleObjects);
                    Assert.Equal("Excel.Sheet.12", ole.ProgId);
                    Assert.True(ole.IsLinked);
                    Assert.Equal(new Uri("file:///C:/does-not-need-to-exist/workbook.xlsx"), ole.LinkUri);
                }
                else
                {
                    Assert.Empty(oleObjects);
                }
            }
        }

        AssertOleSurvivesOnExactlyOneSlide(withOle, expectedSlideIndex: 0);

        var afterReplaceText = PresentationEditor.ReplaceText(
            withOle, new Dictionary<string, string> { ["{{name}}"] = "Alice" });
        AssertOleSurvivesOnExactlyOneSlide(afterReplaceText, expectedSlideIndex: 0);

        var afterReplaceImage = PresentationEditor.ReplaceImage(withOle, "{{name}}", Png());
        AssertOleSurvivesOnExactlyOneSlide(afterReplaceImage, expectedSlideIndex: 0);

        // 1-based: removes "Slide 2", the UNRELATED slide - slide 0's OLE object must survive.
        var afterRemoveSlides = PresentationEditor.RemoveSlides(withOle, new[] { 2 });
        AssertOleSurvivesOnExactlyOneSlide(afterRemoveSlides, expectedSlideIndex: 0);

        // 1-based [2, 1]: swaps the two slides, so the OLE object (originally on slide 1) is now
        // on slide 2 - proving it moves WITH its slide rather than being lost or left behind.
        var afterReorder = PresentationEditor.ReorderSlides(withOle, new[] { 2, 1 });
        AssertOleSurvivesOnExactlyOneSlide(afterReorder, expectedSlideIndex: 1);

        // 1-based atIndex 1: inserts a new first slide, so the original slide 1 (with the OLE
        // object) shifts to position 2 (0-based index 1).
        var afterInsert = PresentationEditor.InsertSlides(
            withOle, 1, new[] { PptxSlide.Titled("New first slide") });
        AssertOleSurvivesOnExactlyOneSlide(afterInsert, expectedSlideIndex: 1);

        var afterProtect = PresentationEditor.Protect(withOle, "pw123");
        var afterUnprotect = PresentationEditor.Unprotect(afterProtect, "pw123");
        AssertOleSurvivesOnExactlyOneSlide(afterUnprotect, expectedSlideIndex: 0);

        // AddChart is OfficeIMO-backed like Protect/Unprotect above, but was never actually
        // exercised against an embedded object until now - closing the same shape of gap the
        // task review already caught once on the XLSX side (WorkbookEditor.AddChart).
        // 1-based slideIndex 1 targets the slide carrying the OLE object.
        var chartData = new ChartData(
            new[] { "North", "South" },
            new[] { new ChartSeries("Total", new double[] { 1200, 980 }) });
        var afterAddChart = PresentationEditor.AddChart(withOle, 1, ChartType.ColumnClustered, chartData);
        AssertOleSurvivesOnExactlyOneSlide(afterAddChart, expectedSlideIndex: 0);
    }

    [Fact]
    public void InspectSignatures_ReportsAnUnsignedDeckCleanly()
    {
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Slide 1") });

        var info = PresentationEditor.InspectSignatures(pptx);

        Assert.False(info.HasSignatures);
        Assert.Equal(0, info.SignatureCount);
    }

    [Fact]
    public async Task InspectSignaturesAsync_MatchesTheByteArrayOverload()
    {
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Slide 1") });

        var expected = PresentationEditor.InspectSignatures(pptx);
        using var source = new MemoryStream(pptx, writable: false);
        var actual = await PresentationEditor.InspectSignaturesAsync(source);

        Assert.Equal(expected.HasSignatures, actual.HasSignatures);
        Assert.Equal(expected.SignatureCount, actual.SignatureCount);
    }

    [Fact]
    public void ValidateSignatures_ReportsAnUnsignedDeckCleanly()
    {
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Slide 1") });

        var report = PresentationEditor.ValidateSignatures(pptx);

        Assert.False(report.HasSignatures);
        Assert.False(report.IsCryptographicallyValid);
        Assert.Empty(report.Signatures);
    }

    [Fact]
    public async Task ValidateSignaturesAsync_MatchesTheByteArrayOverload()
    {
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Slide 1") });

        var expected = PresentationEditor.ValidateSignatures(pptx);
        using var source = new MemoryStream(pptx, writable: false);
        var actual = await PresentationEditor.ValidateSignaturesAsync(source);

        Assert.Equal(expected.HasSignatures, actual.HasSignatures);
        Assert.Equal(expected.IsCryptographicallyValid, actual.IsCryptographicallyValid);
    }

    [Fact]
    public void InspectSignatures_NullPptx_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => PresentationEditor.InspectSignatures(null!));
    }

    [Fact]
    public void InspectSignatures_EmptyPptx_Throws()
    {
        var ex = Assert.Throws<ArgumentException>(() => PresentationEditor.InspectSignatures(Array.Empty<byte>()));
        Assert.Equal("pptx", ex.ParamName);
    }

    // =========================================================================================
    // Metadata
    // =========================================================================================

    private static byte[] FreshDeck() => PresentationEditor.Create(new[] { PptxSlide.Titled("SLIDE-MARKER") });

    [Fact]
    public void MetadataSurvivesARoundTrip()
    {
        var pptx = FreshDeck();

        var stamped = PresentationEditor.WithMetadata(pptx, new DocumentMetadata
        {
            Title = "Quarterly report",
            Creator = "Contoso Ltd",
            Subject = "Revenue",
            Keywords = "revenue, quarterly",
        });

        var read = PresentationEditor.ReadMetadata(stamped);

        Assert.Equal("Quarterly report", read.Title);
        Assert.Equal("Contoso Ltd", read.Creator);
        Assert.Equal("Revenue", read.Subject);
        Assert.Equal("revenue, quarterly", read.Keywords);
    }

    [Fact]
    public void MetadataNotSetReadsBackAsNullRatherThanEmpty()
    {
        // ReadMetadata never saves, so all four properties are null here - identically to DOCX and
        // XLSX. WithMetadata is a different story: see
        // WithMetadata_LeavingCreatorNull_StillEndsUpStampedByOfficeIMO below for why saving loses
        // that distinction for Creator specifically.
        var read = PresentationEditor.ReadMetadata(FreshDeck());

        Assert.Null(read.Title);
        Assert.Null(read.Creator);
        Assert.Null(read.Subject);
        Assert.Null(read.Keywords);
    }

    [Fact]
    public void WithMetadata_LeavingCreatorNull_StillEndsUpStampedByOfficeIMO()
    {
        // Pinned so a future OfficeIMO upgrade that changes this is caught by a red test rather
        // than a silent behaviour drift. See WithMetadata's own remarks for the full mechanism:
        // OfficeIMO.PowerPoint's Save() unconditionally stamps Creator when it is empty, on every
        // save, whether or not this method ever touches Creator.
        var stamped = PresentationEditor.WithMetadata(FreshDeck(), new DocumentMetadata { Title = "T" });

        Assert.Equal("OfficeIMO", PresentationEditor.ReadMetadata(stamped).Creator);
    }

    [Fact]
    public void WithMetadata_ANullPropertyLeavesTheExistingValueInPlace()
    {
        var withTitle = PresentationEditor.WithMetadata(FreshDeck(), new DocumentMetadata { Title = "Original title" });

        var stamped = PresentationEditor.WithMetadata(withTitle, new DocumentMetadata { Creator = "Later author" });

        var read = PresentationEditor.ReadMetadata(stamped);
        Assert.Equal("Original title", read.Title);
        Assert.Equal("Later author", read.Creator);
    }

    [Fact]
    public void WithMetadata_AnEmptyStringClearsAnExistingValue()
    {
        var withTitle = PresentationEditor.WithMetadata(FreshDeck(), new DocumentMetadata { Title = "Original title" });

        var cleared = PresentationEditor.WithMetadata(withTitle, new DocumentMetadata { Title = "" });

        Assert.Equal("", PresentationEditor.ReadMetadata(cleared).Title);
    }

    [Fact]
    public void WithMetadata_DoesNotDisturbTheSlideText()
    {
        var pptx = FreshDeck();

        var stamped = PresentationEditor.WithMetadata(pptx, new DocumentMetadata { Title = "T" });

        Assert.Contains(PresentationEditor.ExtractText(stamped), t => t.Contains("SLIDE-MARKER"));
    }

    [Fact]
    public void ReadMetadata_NullPptx_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => PresentationEditor.ReadMetadata(null!));
    }

    [Fact]
    public void ReadMetadata_EmptyPptx_Throws()
    {
        var ex = Assert.Throws<ArgumentException>(() => PresentationEditor.ReadMetadata(Array.Empty<byte>()));
        Assert.Equal("pptx", ex.ParamName);
    }

    [Fact]
    public void WithMetadata_NullPptx_Throws()
    {
        Assert.Throws<ArgumentNullException>(
            () => PresentationEditor.WithMetadata(null!, new DocumentMetadata()));
    }

    [Fact]
    public void WithMetadata_NullMetadata_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => PresentationEditor.WithMetadata(FreshDeck(), null!));
    }

    [Fact]
    public void WithMetadata_EmptyPptx_Throws()
    {
        var ex = Assert.Throws<ArgumentException>(
            () => PresentationEditor.WithMetadata(Array.Empty<byte>(), new DocumentMetadata()));
        Assert.Equal("pptx", ex.ParamName);
    }
}
