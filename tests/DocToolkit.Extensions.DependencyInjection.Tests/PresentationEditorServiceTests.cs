using System.Collections.Generic;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class PresentationEditorServiceTests
{
    private static byte[] SamplePptx() => File.ReadAllBytes(Path.Combine("assets", "sample.pptx"));

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
}
