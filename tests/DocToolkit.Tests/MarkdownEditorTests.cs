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

    // -----------------------------------------------------------------------------------------
    // A69: FindHeading
    // -----------------------------------------------------------------------------------------

    [Fact]
    public void FindHeading_FindsAnExactMatch()
    {
        const string markdown = "# Overview\n\n## Changed\n\nBody.\n";

        var heading = MarkdownEditor.FindHeading(markdown, "Changed");

        Assert.NotNull(heading);
        Assert.Equal(2, heading!.Level);
        Assert.Equal("Changed", heading.Text);
        Assert.Equal("changed", heading.Anchor);
    }

    [Fact]
    public void FindHeading_MatchesCaseInsensitivelyByDefault()
    {
        const string markdown = "# Overview\n\n## Changed\n\nBody.\n";

        var heading = MarkdownEditor.FindHeading(markdown, "CHANGED");

        Assert.NotNull(heading);
        Assert.Equal("Changed", heading!.Text);
    }

    [Fact]
    public void FindHeading_OnACleanMiss_ReturnsNull()
    {
        const string markdown = "# Overview\n\nBody.\n";

        Assert.Null(MarkdownEditor.FindHeading(markdown, "Does Not Exist"));
    }

    [Fact]
    public void FindHeading_OnDuplicateHeadingText_ReturnsTheFirstMatch()
    {
        const string markdown = "# Notes\n\n## Changed\n\nFirst.\n\n## Changed\n\nSecond.\n";

        var heading = MarkdownEditor.FindHeading(markdown, "Changed");

        Assert.NotNull(heading);
    }

    [Fact]
    public void FindHeading_RejectsNullArguments()
    {
        Assert.Throws<ArgumentNullException>(() => MarkdownEditor.FindHeading(null!, "x"));
        Assert.Throws<ArgumentNullException>(() => MarkdownEditor.FindHeading("# x", null!));
    }
}
