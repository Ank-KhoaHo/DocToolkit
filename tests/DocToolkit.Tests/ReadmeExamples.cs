using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// The README code blocks, as tests. They are injected into the three READMEs by
/// scripts/gen-readme-snippets.py, so a block that stops compiling breaks the build and a
/// block that compiles but is WRONG fails its assertion.
///
/// Separate from DocumentationExamples.cs deliberately: that file feeds the DocFX guides
/// through &lt;code source&gt;, this one feeds markdown through a generator. One file serving
/// two inclusion mechanisms is how a change for one silently reshapes the other.
///
/// Setup ABOVE the region, assertions BELOW it - the reader sees only the capability.
/// </summary>
public class ReadmeExamples
{
    [Fact]
    public void PresentationReplaceImageExample()
    {
        byte[] pptx = PptxFixtures.DeckWithPlaceholderBox("{{chart}}");
        byte[] chartPngBytes = ImageFixtures.Png(40, 30);

        #region readme-pptx-replace-image
        byte[] filled = PresentationEditor.ReplaceImage(pptx, "{{chart}}", chartPngBytes);
        #endregion

        Assert.NotEmpty(filled);
        Assert.DoesNotContain("{{chart}}", string.Join(" ", PresentationEditor.ExtractText(filled)));
    }
}
