using System.Collections.Generic;
using System.IO;
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
}
