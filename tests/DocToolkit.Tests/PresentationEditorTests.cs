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
        Path.Combine(AppContext.BaseDirectory, "assets", "sample.pptx");

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
        // Pins the fixture itself. Every ReplaceImage assertion below is computed FROM these
        // numbers, so a fixture that quietly placed the box elsewhere would make them all agree
        // with each other and with nothing real.
        var pptx = PptxFixtures.DeckWithPlaceholderBox("{{chart}}", 1000000, 2000000, 4000000, 3000000);
        Assert.Empty(PptxFixtures.Validate(pptx));

        using var ms = new MemoryStream(pptx);
        using var doc = PresentationDocument.Open(ms, false);

        var shape = doc.PresentationPart!.SlideParts.Single()
                       .Slide!.CommonSlideData!.ShapeTree!.Elements<P.Shape>().Single();
        var xfrm = shape.ShapeProperties!.Transform2D!;

        Assert.Equal(1000000L, xfrm.Offset!.X!.Value);
        Assert.Equal(2000000L, xfrm.Offset.Y!.Value);
        Assert.Equal(4000000L, xfrm.Extents!.Cx!.Value);
        Assert.Equal(3000000L, xfrm.Extents.Cy!.Value);
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
}
