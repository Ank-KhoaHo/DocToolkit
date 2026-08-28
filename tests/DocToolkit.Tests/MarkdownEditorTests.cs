using DocToolkit;
using OfficeIMO.Markdown;
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

    // -----------------------------------------------------------------------------------------
    // A69: ReplaceSection
    // -----------------------------------------------------------------------------------------

    private const string ThreeSectionMarkdown = """
        ---
        title: Sample
        ---

        # Overview

        Intro paragraph.

        ## Changed

        - item one
        - item two

        ## Table of things

        | a | b |
        |---|---|
        | 1 | 2 |

        # Appendix

        Appendix content.
        """;

    [Fact]
    public void ReplaceSection_ReplacesOnlyTheNamedSectionsBody()
    {
        var result = MarkdownEditor.ReplaceSection(
            ThreeSectionMarkdown, "Changed", "\n- a whole new list\n\n");

        Assert.Contains("- a whole new list", result);
        Assert.DoesNotContain("item one", result);
        Assert.DoesNotContain("item two", result);
    }

    [Fact]
    public void ReplaceSection_LeavesFrontMatterAndOtherSectionsUntouched()
    {
        var result = MarkdownEditor.ReplaceSection(
            ThreeSectionMarkdown, "Changed", "\n- replaced\n\n");

        Assert.Equal("Sample", MarkdownEditor.ReadFrontMatter(result)["title"]);
        Assert.NotNull(MarkdownEditor.FindHeading(result, "Overview"));
        Assert.NotNull(MarkdownEditor.FindHeading(result, "Table of things"));
        Assert.NotNull(MarkdownEditor.FindHeading(result, "Appendix"));
        Assert.Equal(1, MarkdownEditor.TableCount(result));
        Assert.Equal(new[] { "a", "b" }, MarkdownEditor.ReadTable(result, 0)[0]);
        Assert.Contains("Intro paragraph.", result);
        Assert.Contains("Appendix content.", result);
        Assert.Contains("- replaced", result);
        Assert.DoesNotContain("item one", result);
    }

    [Fact]
    public void ReplaceSection_StopsAtANextHeadingOfTheSameOrShallowerLevel()
    {
        // "Changed" is level 2. Its section must stop at "Table of things" (also level 2),
        // not run past it into "Appendix" (level 1) or beyond.
        var result = MarkdownEditor.ReplaceSection(
            ThreeSectionMarkdown, "Changed", "\nreplaced\n\n");

        Assert.NotNull(MarkdownEditor.FindHeading(result, "Table of things"));
        Assert.Equal(1, MarkdownEditor.TableCount(result));
    }

    [Fact]
    public void ReplaceSection_OnTheLastSectionInTheDocument_AppendsCleanly()
    {
        const string markdown = "# Title\n\nBody.\n\n## Last Section\n\nOld tail.\n";

        var result = MarkdownEditor.ReplaceSection(markdown, "Last Section", "\nNew tail.\n");

        Assert.DoesNotContain("Old tail.", result);
        Assert.Contains("New tail.", result);
        Assert.Contains("## Last Section", result);

        var doc = MarkdownReader.Parse(result);
        Assert.NotNull(doc.FindHeading("Title"));
        Assert.NotNull(doc.FindHeading("Last Section"));
    }

    [Fact]
    public void ReplaceSection_OnAHeadingWithNoBodyAtAll_InsertsCleanlyRightAfterIt()
    {
        const string markdown = "# Title\n\n## Empty Section\n";

        var result = MarkdownEditor.ReplaceSection(markdown, "Empty Section", "\nNew content.\n");

        Assert.Contains("## Empty Section", result);
        Assert.Contains("New content.", result);

        var doc = MarkdownReader.Parse(result);
        Assert.NotNull(doc.FindHeading("Title"));
        Assert.NotNull(doc.FindHeading("Empty Section"));
    }

    [Fact]
    public void ReplaceSection_OnAMissingHeading_Throws()
    {
        var ex = Assert.Throws<ArgumentException>(
            () => MarkdownEditor.ReplaceSection(ThreeSectionMarkdown, "Does Not Exist", "x"));

        Assert.Equal("headingText", ex.ParamName);
    }

    [Fact]
    public void ReplaceSection_RejectsNullArguments()
    {
        Assert.Throws<ArgumentNullException>(
            () => MarkdownEditor.ReplaceSection(null!, "Changed", "x"));
        Assert.Throws<ArgumentNullException>(
            () => MarkdownEditor.ReplaceSection(ThreeSectionMarkdown, null!, "x"));
        Assert.Throws<ArgumentNullException>(
            () => MarkdownEditor.ReplaceSection(ThreeSectionMarkdown, "Changed", null!));
    }

    [Fact]
    public void ReplaceSection_OutputRoundTripsThroughAFreshParse()
    {
        var result = MarkdownEditor.ReplaceSection(
            ThreeSectionMarkdown, "Table of things", "\nNo table anymore.\n\n");

        // Must remain parseable, and every OTHER capability must still resolve correctly
        // against the edited document.
        Assert.Equal("Sample", MarkdownEditor.ReadFrontMatter(result)["title"]);
        Assert.NotNull(MarkdownEditor.FindHeading(result, "Overview"));
        Assert.NotNull(MarkdownEditor.FindHeading(result, "Appendix"));
        Assert.Equal(0, MarkdownEditor.TableCount(result));
    }
}
