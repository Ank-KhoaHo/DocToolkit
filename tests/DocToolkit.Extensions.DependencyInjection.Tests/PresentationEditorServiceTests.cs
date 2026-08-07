using System.Collections.Generic;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Xunit;

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
