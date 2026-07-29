using DocumentFormat.OpenXml.Packaging;
using DocToolkit;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

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
}
