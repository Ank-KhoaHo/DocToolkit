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

    /// <summary>
    /// Normalises a raw string literal's line endings, and is load-bearing on every expected
    /// value below rather than cosmetic.
    /// </summary>
    /// <remarks>
    /// <b>A C# raw string literal keeps the line endings of the .cs file it is written in</b> —
    /// measured directly, not assumed: the same literal is 8 characters in a CRLF-saved file and 6
    /// in an LF-saved one. This repository's <c>.gitattributes</c> sets <c>* text=auto</c>, so on
    /// a Windows checkout every source file (and therefore <see cref="ThreeSectionMarkdown"/> and
    /// each expected document below) arrives with <c>\r\n</c>, while on the Linux and macOS
    /// runners it arrives with <c>\n</c>.
    /// <para>
    /// <see cref="MarkdownEditor.ReplaceSection"/> always returns <c>\n</c>, by documented design.
    /// So an expected value left un-normalised would compare equal on Linux and unequal on
    /// Windows — a test whose verdict depends on which runner picked it up. Normalising only the
    /// EXPECTED side is deliberate: the fixtures are handed to the method exactly as the checkout
    /// produced them, so on Windows these tests genuinely drive the CRLF path end to end and still
    /// demand the same answer.
    /// </para>
    /// </remarks>
    private static string Lf(string s) => s.Replace("\r\n", "\n", StringComparison.Ordinal);

    [Fact]
    public void ReplaceSection_ReplacesOnlyTheNamedSectionsBody()
    {
        // Asserted as the WHOLE document, not by containment. `Assert.Contains` cannot see an
        // off-by-one that leaves the right substrings present in a slightly wrong document, which
        // is how a mutated `bodyStart`/`sectionEnd` survived this suite.
        //
        // Derived by hand from the fixture before being run: the kept prefix ends where the next
        // block begins, which is after `## Changed` and the blank line following it, so the
        // leading `\n` of newContent lands as a SECOND blank line; the suffix resumes at the next
        // heading of the same or a shallower level, which is `## Table of things`.
        var expected = Lf("""
            ---
            title: Sample
            ---

            # Overview

            Intro paragraph.

            ## Changed


            - a whole new list

            ## Table of things

            | a | b |
            |---|---|
            | 1 | 2 |

            # Appendix

            Appendix content.
            """);

        var result = MarkdownEditor.ReplaceSection(
            ThreeSectionMarkdown, "Changed", "\n- a whole new list\n\n");

        Assert.Equal(expected, result);
    }

    [Fact]
    public void ReplaceSection_LeavesFrontMatterAndOtherSectionsUntouched()
    {
        var expected = Lf("""
            ---
            title: Sample
            ---

            # Overview

            Intro paragraph.

            ## Changed


            - replaced

            ## Table of things

            | a | b |
            |---|---|
            | 1 | 2 |

            # Appendix

            Appendix content.
            """);

        var result = MarkdownEditor.ReplaceSection(
            ThreeSectionMarkdown, "Changed", "\n- replaced\n\n");

        Assert.Equal(expected, result);

        // Kept beside the exact comparison rather than replaced by it: these read the result back
        // through the OTHER MarkdownEditor methods, so they would catch a document that matched
        // character for character and still failed to parse the way the readers expect.
        Assert.Equal("Sample", MarkdownEditor.ReadFrontMatter(result)["title"]);
        Assert.NotNull(MarkdownEditor.FindHeading(result, "Overview"));
        Assert.NotNull(MarkdownEditor.FindHeading(result, "Table of things"));
        Assert.NotNull(MarkdownEditor.FindHeading(result, "Appendix"));
        Assert.Equal(1, MarkdownEditor.TableCount(result));
        Assert.Equal(new[] { "a", "b" }, MarkdownEditor.ReadTable(result, 0)[0]);
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

        // No same-or-shallower heading follows, so the section runs to the end of the document and
        // the suffix is empty — the whole tail after `## Last Section`'s blank line is replaced,
        // and the document's own trailing newline is newContent's, not the original's.
        Assert.Equal("# Title\n\nBody.\n\n## Last Section\n\n\nNew tail.\n", result);

        var doc = MarkdownReader.Parse(result);
        Assert.NotNull(doc.FindHeading("Title"));
        Assert.NotNull(doc.FindHeading("Last Section"));
    }

    [Fact]
    public void ReplaceSection_OnAHeadingWithNoBodyAtAll_InsertsCleanlyRightAfterIt()
    {
        const string markdown = "# Title\n\n## Empty Section\n";

        var result = MarkdownEditor.ReplaceSection(markdown, "Empty Section", "\nNew content.\n");

        // This is the case the plan called highest-risk. The heading has no NextSibling, so
        // bodyStart falls back to the document's LENGTH and the heading line keeps its own
        // trailing newline. The tempting `heading.SourceSpan.EndOffset + 1` fallback is off by
        // one and would produce "## Empty Section\nNew content.\n" — still containing both
        // strings, which is exactly why this is asserted as an exact document.
        Assert.Equal("# Title\n\n## Empty Section\n\nNew content.\n", result);

        var doc = MarkdownReader.Parse(result);
        Assert.NotNull(doc.FindHeading("Title"));
        Assert.NotNull(doc.FindHeading("Empty Section"));
    }

    [Fact]
    public void ReplaceSection_OnCrlfInput_NormalisesRatherThanCorruptingTheDocument()
    {
        // Built with EXPLICIT \r\n rather than as a raw string literal, on purpose: a raw literal
        // carries whatever line endings the checkout gave this .cs file, so on Linux it would not
        // exercise the CRLF path at all and this regression test would pass vacuously on three of
        // the four CI platforms.
        //
        // OfficeIMO.Markdown measures every SourceSpan against LF-normalised text. Splicing those
        // offsets into the original CRLF string shifts each one left by the number of preceding
        // line breaks, which truncates the heading and corrupts an unrelated line further down
        // with no exception raised.
        var markdown = string.Join("\r\n", new[]
        {
            "# Release Notes",
            "",
            "Intro that must survive.",
            "",
            "## Changed",
            "",
            "- old item",
            "",
            "## Fixed",
            "",
            "- a fix that must survive intact",
            "",
        });

        var result = MarkdownEditor.ReplaceSection(markdown, "Changed", "\n- new item\n\n");

        Assert.Equal(
            "# Release Notes\n\nIntro that must survive.\n\n## Changed\n\n\n- new item\n\n"
            + "## Fixed\n\n- a fix that must survive intact\n",
            result);

        // Named individually so a failure says WHICH corruption came back, rather than only that
        // two long strings differ: the heading complete rather than truncated, the untouched
        // paragraph and the untouched later section whole rather than merged with a neighbour,
        // and no stray \r left anywhere.
        Assert.Contains("\n## Changed\n", result, StringComparison.Ordinal);
        Assert.Contains("\nIntro that must survive.\n", result, StringComparison.Ordinal);
        Assert.Contains("\n## Fixed\n", result, StringComparison.Ordinal);
        Assert.Contains("\n- a fix that must survive intact\n", result, StringComparison.Ordinal);
        Assert.DoesNotContain("\r", result, StringComparison.Ordinal);
    }

    [Fact]
    public void ReplaceSection_OnAHeadingNestedInABlockquote_Throws()
    {
        // FindHeading matches this heading, but its IndexInParent counts within the QuoteBlock,
        // not within the document — so the boundary walk would index the wrong list. Measured
        // before the guard existed: the document's tail came back duplicated, silently.
        var ex = Assert.Throws<ArgumentException>(() => MarkdownEditor.ReplaceSection(
            "> ## Deep\n>\n> p3\n\n## Boundary\n\ntail\n", "Deep", "x"));

        Assert.Equal("headingText", ex.ParamName);
        Assert.Contains("nested inside another block", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ReplaceSection_OnAHeadingNestedInAListItem_Throws()
    {
        // The same mismatch through the other container. This shape's measured symptom was the
        // opposite one — the edit silently did nothing at all — which is why both are pinned.
        var ex = Assert.Throws<ArgumentException>(() => MarkdownEditor.ReplaceSection(
            "- # Nested\n\n# Boundary\n\ntail\n", "Nested", "x"));

        Assert.Equal("headingText", ex.ParamName);
        Assert.Contains("nested inside another block", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ReplaceSection_OnATopLevelHeading_IsNotRejectedByTheNestingGuard()
    {
        // The positive control for the two tests above. A guard that rejected everything would
        // satisfy them both, so this asserts the ordinary case still edits — and asserts it
        // against an exact document, so "did not throw" is not mistaken for "did the right thing".
        var result = MarkdownEditor.ReplaceSection(
            "## Boundary\n\nold\n", "Boundary", "\nnew\n");

        Assert.Equal("## Boundary\n\n\nnew\n", result);
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
