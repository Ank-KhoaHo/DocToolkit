using DocToolkit;
using ShapeCrawler;
using Xunit;

namespace DocToolkit.Tests;

public class PresentationEditorTests
{
    /// <summary>Builds a one-slide deck with a single text box reading "Hello {{who}}".</summary>
    private static byte[] SampleDeck()
    {
        using var pres = new Presentation();
        pres.Slides.Add(1);
        var slide = pres.Slide(1);
        // ShapeCrawler 0.79.4 exposes AddTextBox (not AddText, which does not resolve on
        // IUserSlideShapeCollection in this version) — same signature, same effect.
        slide.Shapes.AddTextBox(50, 50, 400, 100, "Hello {{who}}");

        using var ms = new MemoryStream();
        pres.Save(ms);
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
}
