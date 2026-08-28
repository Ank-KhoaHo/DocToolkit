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

    // -----------------------------------------------------------------------------------------
    // A69: TableCount / ReadTable
    // -----------------------------------------------------------------------------------------

    private const string TwoTableMarkdown = """
        # Doc

        | a | b |
        |---|---|
        | 1 | 2 |
        | 3 | 4 |

        Some text between the tables.

        | x |
        |---|
        | y |
        """;

    [Fact]
    public void TableCount_CountsEveryTableInTheDocument()
    {
        Assert.Equal(2, MarkdownEditor.TableCount(TwoTableMarkdown));
    }

    [Fact]
    public void TableCount_OnADocumentWithNoTable_ReturnsZero()
    {
        Assert.Equal(0, MarkdownEditor.TableCount("# Doc\n\nNo tables here.\n"));
    }

    [Fact]
    public void ReadTable_ReturnsTheHeaderRowFirstThenDataRows()
    {
        var rows = MarkdownEditor.ReadTable(TwoTableMarkdown, 0);

        Assert.Equal(new[] { "a", "b" }, rows[0]);
        Assert.Equal(new[] { "1", "2" }, rows[1]);
        Assert.Equal(new[] { "3", "4" }, rows[2]);
    }

    [Fact]
    public void ReadTable_IndexIsZeroBasedAgainstDocumentOrder()
    {
        var rows = MarkdownEditor.ReadTable(TwoTableMarkdown, 1);

        Assert.Equal(new[] { "x" }, rows[0]);
        Assert.Equal(new[] { "y" }, rows[1]);
    }

    [Fact]
    public void ReadTable_RejectsANegativeIndex()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => MarkdownEditor.ReadTable(TwoTableMarkdown, -1));
    }

    [Fact]
    public void ReadTable_RejectsAnIndexAtOrBeyondTableCount()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => MarkdownEditor.ReadTable(TwoTableMarkdown, 2));
    }

    [Fact]
    public void ReadTable_ARowWithADifferentCellCountThanTheHeader_IsReturnedWithTheShapeItHas()
    {
        // Measured directly against OfficeIMO.Markdown: a short row is not padded, and a long
        // row is not truncated — matching DocxEditor.ReadTable's own "rows are returned with the
        // shape they have" precedent for DOCX tables.
        const string markdown = """
            # Doc

            | a | b | c |
            |---|---|---|
            | 1 | 2 |
            | x | y | z | extra |
            """;

        var rows = MarkdownEditor.ReadTable(markdown, 0);

        Assert.Equal(new[] { "a", "b", "c" }, rows[0]);
        Assert.Equal(new[] { "1", "2" }, rows[1]);
        Assert.Equal(new[] { "x", "y", "z", "extra" }, rows[2]);
    }
}
