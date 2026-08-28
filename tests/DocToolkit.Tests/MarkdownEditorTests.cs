using DocToolkit;
using Xunit;

namespace DocToolkit.Tests;

public class MarkdownEditorTests
{
    // -----------------------------------------------------------------------------------------
    // A69: ReadFrontMatter
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void ReadFrontMatter_ReturnsEveryKeyWithItsValue()
    {
        const string markdown = """
            ---
            title: Release notes
            version: 3
            draft: false
            ---

            # Body
            """;

        var frontMatter = MarkdownEditor.ReadFrontMatter(markdown);

        Assert.Equal("Release notes", frontMatter["title"]);
        Assert.Equal(3.0, frontMatter["version"]);
        Assert.Equal(false, frontMatter["draft"]);
    }

    [Fact]
    public void ReadFrontMatter_OnADocumentWithNone_ReturnsAnEmptyDictionary()
    {
        const string markdown = "# Just a heading\n\nNo front matter here.\n";

        var frontMatter = MarkdownEditor.ReadFrontMatter(markdown);

        Assert.Empty(frontMatter);
    }

    [Fact]
    public void ReadFrontMatter_RejectsNullMarkdown()
    {
        Assert.Throws<ArgumentNullException>(() => MarkdownEditor.ReadFrontMatter(null!));
    }
}
