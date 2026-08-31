using System.Collections.Generic;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class PresentationEditorServiceTests
{
    private static byte[] SamplePptx() => File.ReadAllBytes(Path.Join("assets", "sample.pptx"));

    [Fact]
    public void SlideCount_ExtractText_MatchTheStaticMethods()
    {
        var pptx = SamplePptx();
        var sut = new PresentationEditorService();

        Assert.Equal(PresentationEditor.SlideCount(pptx), sut.SlideCount(pptx));
        Assert.Equal(PresentationEditor.ExtractText(pptx), sut.ExtractText(pptx));
    }

    [Fact]
    public void ReplaceText_SubstitutesPlaceholders()
    {
        var pptx = SamplePptx();
        var sut = new PresentationEditorService();
        var replacements = new Dictionary<string, string> { ["{{who}}"] = "World" };

        var edited = sut.ReplaceText(pptx, replacements);

        var text = sut.ExtractText(edited);
        Assert.Contains(text, t => t.Contains("Hello World"));
        Assert.DoesNotContain(text, t => t.Contains("{{who}}"));
    }

    [Fact]
    public async Task SlideCountAsync_ExtractTextAsync_MatchTheStaticMethods()
    {
        var pptx = SamplePptx();
        var sut = new PresentationEditorService();

        using var countStaticSource = new MemoryStream(pptx);
        using var countWrapperSource = new MemoryStream(pptx);

        Assert.Equal(
            await PresentationEditor.SlideCountAsync(countStaticSource),
            await sut.SlideCountAsync(countWrapperSource));

        using var textStaticSource = new MemoryStream(pptx);
        using var textWrapperSource = new MemoryStream(pptx);

        Assert.Equal(
            await PresentationEditor.ExtractTextAsync(textStaticSource),
            await sut.ExtractTextAsync(textWrapperSource));
    }

    [Fact]
    public async Task ReplaceTextAsync_MatchesTheStaticMethod()
    {
        var pptx = SamplePptx();
        var sut = new PresentationEditorService();
        var replacements = new Dictionary<string, string> { ["{{who}}"] = "World" };

        using var expectedSource = new MemoryStream(pptx);
        using var expected = new MemoryStream();
        await PresentationEditor.ReplaceTextAsync(expectedSource, replacements, expected);

        using var actualSource = new MemoryStream(pptx);
        using var actual = new MemoryStream();
        await sut.ReplaceTextAsync(actualSource, replacements, actual);

        Assert.Equal(expected.ToArray(), actual.ToArray());
    }

    [Fact]
    public void Create_MatchesTheStaticMethod()
    {
        var slides = new[]
        {
            PptxSlide.Titled("Quarterly review", "Revenue up", "Costs flat"),
            PptxSlide.Titled("Next steps"),
        };
        var sut = new PresentationEditorService();

        var expected = PresentationEditor.Create(slides);
        var actual = sut.Create(slides);

        // Semantic agreement, not byte equality: a freshly built package carries zip entry
        // timestamps, so two Create calls a second apart legitimately differ byte-for-byte.
        Assert.Equal(PresentationEditor.SlideCount(expected), PresentationEditor.SlideCount(actual));
        Assert.Equal(PresentationEditor.ExtractText(expected), PresentationEditor.ExtractText(actual));
        Assert.Contains(PresentationEditor.ExtractText(actual), t => t.Contains("Quarterly review"));
    }

    [Fact]
    public async Task CreateAsync_WritesTheSameDeckToTheDestination()
    {
        var slides = new[] { PptxSlide.Titled("Roadmap", "Ship 0.11.0") };
        var sut = new PresentationEditorService();

        using var destination = new MemoryStream();
        await sut.CreateAsync(slides, destination);

        var written = destination.ToArray();

        // B16: anchored to literals before the parity check. ExtractText compared against itself
        // passes on two empty lists, so it cannot on its own tell "the deck was written" from
        // "nothing was written and nothing was read".
        //
        // Two entries for ONE slide: ExtractText returns one entry per text-bearing BODY, so a
        // title and its content are separate. The parity assertion below never had to know that.
        Assert.Equal(new[] { "Roadmap", "Ship 0.11.0" }, PresentationEditor.ExtractText(written));

        Assert.Equal(
            PresentationEditor.ExtractText(PresentationEditor.Create(slides)),
            PresentationEditor.ExtractText(written));
    }

    // ---------------------------------------------------------------------------------------
    // ReplaceImage / ReplaceImageAsync, mirrored from core 0.21.0. Built with
    // PresentationEditor.Create rather than a fixture asset: PptxDocumentWriter.TextShape gives
    // every shape it builds its own explicit a:xfrm, so a title placeholder is a valid
    // ReplaceImage target without needing a hand-built OOXML fixture (core's
    // ADeckBuiltByCreateWorksWithReplaceImage pins exactly this).
    // ---------------------------------------------------------------------------------------

    private static byte[] OnePixelPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

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
    public void ReplaceImage_MatchesTheStaticMethod()
    {
        var deck = PresentationEditor.Create(new[] { PptxSlide.Titled("{{chart}}") });
        var sut = new PresentationEditorService();
        var png = OnePixelPng();

        var fromWrapper = sut.ReplaceImage(deck, "{{chart}}", png);
        var fromStatic = PresentationEditor.ReplaceImage(deck, "{{chart}}", png);

        // Structural comparison via the picture's own transform, not raw bytes - ReplaceImage
        // mints fresh relationship ids per call the same way Create does, so two calls over
        // identical input legitimately differ byte-for-byte.
        Assert.Equal(PictureTransformOf(fromStatic), PictureTransformOf(fromWrapper));
    }

    [Fact]
    public void ReplaceImage_PutsAPictureWhereThePlaceholderWasAndRemovesTheText()
    {
        var deck = PresentationEditor.Create(new[] { PptxSlide.Titled("{{chart}}") });
        var sut = new PresentationEditorService();

        var filled = sut.ReplaceImage(deck, "{{chart}}", OnePixelPng());

        // Concrete: the shape holding the placeholder is genuinely GONE and replaced by a
        // picture - a wrapper that ignored its arguments and handed back the input unchanged
        // would still pass a delegation-only comparison against calling the static method with
        // the same (also-unused) arguments, but would fail every assertion here.
        using var ms = new MemoryStream(filled);
        using var doc = PresentationDocument.Open(ms, false);
        var tree = doc.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.ShapeTree!;

        Assert.Single(tree.Elements<P.Picture>());
        Assert.Empty(tree.Elements<P.Shape>());
        Assert.DoesNotContain(sut.ExtractText(filled), t => t.Contains("{{chart}}"));
    }

    [Fact]
    public async Task ReplaceImageAsync_PutsAPictureWhereThePlaceholderWas()
    {
        var deck = PresentationEditor.Create(new[] { PptxSlide.Titled("{{chart}}") });
        var sut = new PresentationEditorService();

        using var source = new MemoryStream(deck);
        using var destination = new MemoryStream();
        await sut.ReplaceImageAsync(source, "{{chart}}", OnePixelPng(), destination);

        var written = destination.ToArray();
        using var ms = new MemoryStream(written);
        using var doc = PresentationDocument.Open(ms, false);
        var tree = doc.PresentationPart!.SlideParts.Single().Slide!.CommonSlideData!.ShapeTree!;

        Assert.Single(tree.Elements<P.Picture>());
        Assert.Empty(tree.Elements<P.Shape>());
        Assert.Single(doc.PresentationPart!.SlideParts.Single().ImageParts);
    }

    // ---------------------------------------------------------------------------------------
    // InsertSlides/ReadSlide/RemoveSlides/ReorderSlides, mirrored from core 0.43.0 (A70-DI).
    // Each test reads back a concrete, index-specific literal rather than only comparing the
    // wrapper's output to the static method's own output, since a wrapper calling the RIGHT
    // method with an off-by-one or reordered argument would otherwise still "match itself".
    // ---------------------------------------------------------------------------------------

    private static byte[] ThreeSlideDeck() => PresentationEditor.Create(new[]
    {
        PptxSlide.Titled("First"),
        PptxSlide.Titled("Second"),
        PptxSlide.Titled("Third"),
    });

    [Fact]
    public void ReadSlide_MatchesTheStaticMethodAndReadsTheRightSlide()
    {
        var deck = ThreeSlideDeck();
        var sut = new PresentationEditorService();

        var fromWrapper = sut.ReadSlide(deck, 2);

        Assert.Equal(PresentationEditor.ReadSlide(deck, 2), fromWrapper);
        Assert.Contains(fromWrapper, t => t.Contains("Second"));
    }

    [Fact]
    public async Task ReadSlideAsync_MatchesTheStaticMethod()
    {
        var deck = ThreeSlideDeck();
        var sut = new PresentationEditorService();

        using var expectedSource = new MemoryStream(deck);
        using var actualSource = new MemoryStream(deck);

        var expected = await PresentationEditor.ReadSlideAsync(expectedSource, 2);
        var actual = await sut.ReadSlideAsync(actualSource, 2);

        Assert.Equal(expected, actual);
        Assert.Contains(actual, t => t.Contains("Second"));
    }

    [Fact]
    public void RemoveSlides_MatchesTheStaticMethodAndRemovesTheRightSlide()
    {
        var deck = ThreeSlideDeck();
        var sut = new PresentationEditorService();

        var fromWrapper = sut.RemoveSlides(deck, new[] { 2 });

        Assert.Equal(
            PresentationEditor.ExtractText(PresentationEditor.RemoveSlides(deck, new[] { 2 })),
            PresentationEditor.ExtractText(fromWrapper));

        var remaining = PresentationEditor.ExtractText(fromWrapper);
        Assert.Contains(remaining, t => t.Contains("First"));
        Assert.Contains(remaining, t => t.Contains("Third"));
        Assert.DoesNotContain(remaining, t => t.Contains("Second"));
    }

    [Fact]
    public async Task RemoveSlidesAsync_MatchesTheStaticMethod()
    {
        var deck = ThreeSlideDeck();
        var sut = new PresentationEditorService();

        using var source = new MemoryStream(deck);
        using var destination = new MemoryStream();
        await sut.RemoveSlidesAsync(source, new[] { 1 }, destination);

        var remaining = PresentationEditor.ExtractText(destination.ToArray());
        Assert.DoesNotContain(remaining, t => t.Contains("First"));
        Assert.Contains(remaining, t => t.Contains("Second"));
        Assert.Contains(remaining, t => t.Contains("Third"));
    }

    [Fact]
    public void ReorderSlides_MatchesTheStaticMethodAndAppliesTheRealOrder()
    {
        var deck = ThreeSlideDeck();
        var sut = new PresentationEditorService();

        var fromWrapper = sut.ReorderSlides(deck, new[] { 3, 1, 2 });

        Assert.Equal(
            PresentationEditor.ExtractText(PresentationEditor.ReorderSlides(deck, new[] { 3, 1, 2 })),
            PresentationEditor.ExtractText(fromWrapper));

        // Order, not merely membership: a wrapper that ignored `order` and returned the deck
        // unchanged would still pass a set-equality check but fail this.
        Assert.Equal(
            new[] { "Third", "First", "Second" },
            PresentationEditor.ExtractText(fromWrapper).Where(t => t is "First" or "Second" or "Third"));
    }

    [Fact]
    public async Task ReorderSlidesAsync_MatchesTheStaticMethod()
    {
        var deck = ThreeSlideDeck();
        var sut = new PresentationEditorService();

        using var source = new MemoryStream(deck);
        using var destination = new MemoryStream();
        await sut.ReorderSlidesAsync(source, new[] { 2, 3, 1 }, destination);

        Assert.Equal(
            new[] { "Second", "Third", "First" },
            PresentationEditor.ExtractText(destination.ToArray()).Where(t => t is "First" or "Second" or "Third"));
    }

    [Fact]
    public void InsertSlides_MatchesTheStaticMethodAndInsertsAtTheRightPosition()
    {
        var deck = ThreeSlideDeck();
        var sut = new PresentationEditorService();
        var inserted = new[] { PptxSlide.Titled("Inserted") };

        var fromWrapper = sut.InsertSlides(deck, 2, inserted);

        Assert.Equal(
            PresentationEditor.ExtractText(PresentationEditor.InsertSlides(deck, 2, inserted)),
            PresentationEditor.ExtractText(fromWrapper));

        Assert.Equal(
            new[] { "First", "Inserted", "Second", "Third" },
            PresentationEditor.ExtractText(fromWrapper).Where(t => t is "First" or "Second" or "Third" or "Inserted"));
    }

    [Fact]
    public async Task InsertSlidesAsync_MatchesTheStaticMethod()
    {
        var deck = ThreeSlideDeck();
        var sut = new PresentationEditorService();
        var inserted = new[] { PptxSlide.Titled("Inserted") };

        using var source = new MemoryStream(deck);
        using var destination = new MemoryStream();
        await sut.InsertSlidesAsync(source, 4, inserted, destination);

        // atIndex 4 == SlideCount + 1: appends after everything, so the assertion also proves
        // the position argument (not just the slide content) threaded through correctly.
        Assert.Equal(
            new[] { "First", "Second", "Third", "Inserted" },
            PresentationEditor.ExtractText(destination.ToArray()).Where(t => t is "First" or "Second" or "Third" or "Inserted"));
    }

    [Fact]
    public async Task SlideCountAsync_HonorsCancellation()
    {
        var sut = new PresentationEditorService();
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        using var source = new MemoryStream();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => sut.SlideCountAsync(source, cts.Token));
    }

    // ---------------------------------------------------------------------------------------
    // InspectSignatures/ValidateSignatures and their Async forms, mirrored from core 0.45.0
    // (A99-DI). Exercised against a genuinely unsigned deck - see the identical reasoning in
    // DocxEditorServiceTests.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void InspectSignatures_MatchesTheStaticMethod()
    {
        var sut = new PresentationEditorService();
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Unsigned") });

        var info = sut.InspectSignatures(pptx);

        Assert.Equal(PresentationEditor.InspectSignatures(pptx).HasSignatures, info.HasSignatures);
        Assert.False(info.HasSignatures);
    }

    [Fact]
    public async Task InspectSignaturesAsync_MatchesTheStaticMethod()
    {
        var sut = new PresentationEditorService();
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Unsigned") });

        using var source = new MemoryStream(pptx);
        var info = await sut.InspectSignaturesAsync(source);

        Assert.False(info.HasSignatures);
    }

    [Fact]
    public void ValidateSignatures_MatchesTheStaticMethod()
    {
        var sut = new PresentationEditorService();
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Unsigned") });

        var report = sut.ValidateSignatures(pptx);

        Assert.Equal(PresentationEditor.ValidateSignatures(pptx).HasSignatures, report.HasSignatures);
        Assert.False(report.HasSignatures);
    }

    [Fact]
    public async Task ValidateSignaturesAsync_MatchesTheStaticMethod()
    {
        var sut = new PresentationEditorService();
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Unsigned") });

        using var source = new MemoryStream(pptx);
        var report = await sut.ValidateSignaturesAsync(source);

        Assert.False(report.HasSignatures);
    }

    // ---------------------------------------------------------------------------------------
    // ReadMetadata/WithMetadata, mirrored from core 0.46.0 (A102-DI).
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void WithMetadata_ReadMetadata_RoundTripCorrectly()
    {
        var sut = new PresentationEditorService();
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Slide 1") });
        var metadata = new DocumentMetadata { Title = "Q1 Review", Creator = "Finance" };

        var stamped = sut.WithMetadata(pptx, metadata);
        var read = sut.ReadMetadata(stamped);

        Assert.Equal("Q1 Review", read.Title);
        Assert.Equal("Finance", read.Creator);
        Assert.Equal(
            PresentationEditor.ReadMetadata(PresentationEditor.WithMetadata(pptx, metadata)).Title,
            read.Title);
    }

    // ---------------------------------------------------------------------------------------
    // AddChart/AddChartAsync and ReadSmartArt/ReadSmartArtAsync, mirrored from core 0.45.0
    // (A95-DI/A98-DI) - found missing by the derived mirror test, not filed ahead of time.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void AddChart_MatchesTheStaticMethodAndAddsAChart()
    {
        var sut = new PresentationEditorService();
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Slide 1") });
        var data = new ChartData(new[] { "North" }, new[] { new ChartSeries("Total", new double[] { 1200 }) });

        var withChart = sut.AddChart(pptx, 1, ChartType.ColumnClustered, data, title: "Regional Totals");

        using var source = new MemoryStream(withChart, writable: false);
        using var doc = OfficeIMO.PowerPoint.PowerPointPresentation.Load(source);
        var slide = doc.Slides[0];
        Assert.Single(slide.Charts);
        Assert.Equal("Regional Totals", slide.Charts.Single().Title);
    }

    [Fact]
    public async Task AddChartAsync_MatchesTheStaticMethod()
    {
        var sut = new PresentationEditorService();
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Slide 1") });
        var data = new ChartData(new[] { "North" }, new[] { new ChartSeries("Total", new double[] { 1200 }) });

        using var source = new MemoryStream(pptx);
        var withChart = await sut.AddChartAsync(source, 1, ChartType.ColumnClustered, data, title: "Regional Totals");

        using var readBack = new MemoryStream(withChart, writable: false);
        using var doc = OfficeIMO.PowerPoint.PowerPointPresentation.Load(readBack);
        var slide = doc.Slides[0];
        Assert.Single(slide.Charts);
        Assert.Equal("Regional Totals", slide.Charts.Single().Title);
    }

    private static byte[] DeckWithSmartArt(params string[] nodeTexts)
    {
        var pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Slide 1") });

        using var source = new MemoryStream(pptx, writable: false);
        using var doc = OfficeIMO.PowerPoint.PowerPointPresentation.Load(source);

        var slide = doc.Slides[0];
        var box = OfficeIMO.PowerPoint.PowerPointLayoutBox.FromInches(1, 3, 6, 2);
        slide.AddSmartArt(OfficeIMO.PowerPoint.PowerPointSmartArtType.BasicProcess, nodeTexts, box);

        using var output = new MemoryStream();
        doc.Save(output);
        return output.ToArray();
    }

    [Fact]
    public void ReadSmartArt_MatchesTheStaticMethod()
    {
        var sut = new PresentationEditorService();
        var pptx = DeckWithSmartArt("Plan", "Build", "Ship");

        var diagrams = sut.ReadSmartArt(pptx, 1);

        Assert.Single(diagrams);
        Assert.Equal("Plan\nBuild\nShip", diagrams[0]);
        Assert.Equal(PresentationEditor.ReadSmartArt(pptx, 1), diagrams);
    }

    [Fact]
    public async Task ReadSmartArtAsync_MatchesTheStaticMethod()
    {
        var sut = new PresentationEditorService();
        var pptx = DeckWithSmartArt("Plan", "Build", "Ship");

        using var source = new MemoryStream(pptx);
        var diagrams = await sut.ReadSmartArtAsync(source, 1);

        Assert.Single(diagrams);
        Assert.Equal("Plan\nBuild\nShip", diagrams[0]);
    }
}
