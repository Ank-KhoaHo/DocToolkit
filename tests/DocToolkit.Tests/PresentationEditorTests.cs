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
}
