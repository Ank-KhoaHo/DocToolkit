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

    [Fact]
    public void ReadFrontMatter_OnASequenceOrANestedMapping_ReturnsTheMeasuredShape()
    {
        // Pins the three non-scalar shapes the <remarks> now documents. Each was measured against
        // the pinned OfficeIMO.Markdown 3.2.6 before being written down, and each is surprising
        // enough that a doc comment alone would drift: two of the three LOSE data, and one
        // introduces a fourth runtime type the summary of "string, double or bool" did not cover.

        // 1. A BLOCK sequence is not read at all. `tags` is present and empty; alpha and beta are
        //    gone, with nothing raised. Asserted as the exact value, because "the items are
        //    missing" and "the key is missing" are different bugs.
        var block = MarkdownEditor.ReadFrontMatter(
            "---\ntags:\n  - alpha\n  - beta\n---\n\n# Body\n");
        Assert.Equal(string.Empty, Assert.IsType<string>(block["tags"]));

        // 2. An INLINE sequence is read, as a List<string> — the fourth runtime type. Its items are
        //    always strings, so a numeric one does NOT become a double the way a bare scalar does.
        var inline = MarkdownEditor.ReadFrontMatter(
            "---\ntags: [alpha, beta]\nnums: [1, 2]\n---\n\n# Body\n");
        Assert.Equal(new[] { "alpha", "beta" }, Assert.IsType<List<string>>(inline["tags"]));
        Assert.Equal(new[] { "1", "2" }, Assert.IsType<List<string>>(inline["nums"]));

        // 3. A NESTED mapping FLATTENS rather than nesting: the parent maps to an empty string and
        //    its children become top-level keys of the returned dictionary.
        var nested = MarkdownEditor.ReadFrontMatter(
            "---\nauthor:\n  name: Ada\n  email: a@b.c\n---\n\n# Body\n");
        Assert.Equal(3, nested.Count);
        Assert.Equal(string.Empty, nested["author"]);
        Assert.Equal("Ada", nested["name"]);
        Assert.Equal("a@b.c", nested["email"]);
    }

    [Fact]
    public void ReadFrontMatter_WhenAFlattenedKeyCollides_TheLaterValueSilentlyWins()
    {
        // The consequence of the flattening above that actually loses a caller's data, kept as its
        // own test because it is the one worth being warned about: a top-level `name` and a nested
        // `author.name` are the SAME key by the time this returns, and only one value survives.
        var frontMatter = MarkdownEditor.ReadFrontMatter(
            "---\nname: Outer\nauthor:\n  name: Ada\n---\n\n# Body\n");

        Assert.Equal(2, frontMatter.Count);
        Assert.Equal("Ada", frontMatter["name"]);
        Assert.Equal(string.Empty, frontMatter["author"]);
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
        // The two headings deliberately differ in LEVEL. With both at `##` this test asserted only
        // that something came back, so a "return the LAST match" implementation passed it too —
        // and the distinction is load-bearing rather than pedantic: a changelog repeating a section
        // title once per version is this capability's own motivating example, and
        // ReplaceSection resolves exactly this shape by taking the first top-level match.
        const string markdown = "# Notes\n\n## Changed\n\nFirst.\n\n### Changed\n\nSecond.\n";

        var heading = MarkdownEditor.FindHeading(markdown, "Changed");

        Assert.NotNull(heading);
        Assert.Equal(2, heading!.Level);
    }

    [Fact]
    public void FindHeading_FindsANestedHeading_UnlikeReplaceSection()
    {
        // Pins the asymmetry both doc comments now name. FindHeading searches every heading;
        // ReplaceSection considers only top-level ones, because it has to index the document's own
        // block list to find a section's boundaries. Asserted from both sides in one test, so the
        // two claims cannot drift apart independently.
        const string markdown = "> ## Quoted\n>\n> body\n\n# Real\n\ntail\n";

        var heading = MarkdownEditor.FindHeading(markdown, "Quoted");
        Assert.NotNull(heading);
        Assert.Equal(2, heading!.Level);

        Assert.Throws<ArgumentException>(
            () => MarkdownEditor.ReplaceSection(markdown, "Quoted", "x"));
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
    public void TableCount_RejectsNullMarkdown()
    {
        Assert.Throws<ArgumentNullException>(() => MarkdownEditor.TableCount(null!));
    }

    [Fact]
    public void ReadTable_RejectsNullMarkdown()
    {
        Assert.Throws<ArgumentNullException>(() => MarkdownEditor.ReadTable(null!, 0));
    }

    [Fact]
    public void ReadTable_ReturnsRowsTheCallerCannotMutate()
    {
        // Measured against the pinned OfficeIMO.Markdown 3.2.6: the parser's own header row is a
        // List<string>, and each data row is a string[]. Returning either straight back behind an
        // IReadOnlyList<string> would let a caller cast the reference and edit the parsed document
        // — not what read-only means — and would diverge from DocxEditor.ReadTableCore, which
        // builds a fresh list per row.
        //
        // Both source shapes are checked by their OWN real mutation vector, not the same one for
        // both: an unwrapped List<string> would ACCEPT .Add (proving the wrapper), while a
        // fixed-size array already rejects .Add regardless of wrapping — so .Add alone would pass
        // for an unwrapped array and prove nothing for it; only the indexer write does, since an
        // unwrapped array accepts that silently. The genuine read-only wrapper must refuse both.
        var rows = MarkdownEditor.ReadTable(TwoTableMarkdown, 0);

        foreach (var row in rows)
        {
            Assert.IsNotType<List<string>>(row);
            Assert.IsNotType<string[]>(row);
            Assert.Throws<NotSupportedException>(() => ((IList<string>)row).Add("injected"));
            Assert.Throws<NotSupportedException>(() => ((IList<string>)row)[0] = "MUTATED");
        }

        // Positive control: the values are unchanged by the copying, so this is a type change and
        // not a quiet behaviour change.
        Assert.Equal(new[] { "a", "b" }, rows[0]);
        Assert.Equal(new[] { "1", "2" }, rows[1]);
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
    public void ReplaceSection_NormalisesOnlyWhatCameFromMarkdown_AndOnlyCrLf()
    {
        // The <remarks> used to claim, flatly, that line endings in the result are normalised to
        // \n. Measured false in two directions — neither of them a bug, both of them a documented
        // claim that was wider than the code — so this pins the narrowed statement.

        // 1. newContent is spliced in VERBATIM. A \r\n written there survives.
        var fromNewContent = MarkdownEditor.ReplaceSection("## S\n\nold\n", "S", "\r\nnew\r\n");
        Assert.Equal("## S\n\n\r\nnew\r\n", fromNewContent);

        // 2. Only \r\n is recognised as a line ending. A LONE \r in a part of `markdown` the edit
        //    keeps is left exactly as it is.
        var loneCarriageReturn = MarkdownEditor.ReplaceSection(
            "## S\n\nbody\n\n## T\n\nkeep\rme\n", "S", "\nnew\n");
        Assert.Equal("## S\n\n\nnew\n## T\n\nkeep\rme\n", loneCarriageReturn);
    }

    [Fact]
    public void ReplaceSection_OnAHeadingNestedInABlockquote_ReportsNoMatch()
    {
        // ReplaceSection searches only the document's own top-level blocks, so this heading is
        // never a candidate — not found and refused, simply not looked at. The reason it must not
        // be a candidate: its IndexInParent counts within the QuoteBlock rather than within the
        // document, so the boundary walk would index an unrelated list. Measured before any guard
        // existed, on this exact fixture, the document's tail came back DUPLICATED, silently.
        var ex = Assert.Throws<ArgumentException>(() => MarkdownEditor.ReplaceSection(
            "> ## Deep\n>\n> p3\n\n## Boundary\n\ntail\n", "Deep", "x"));

        Assert.Equal("headingText", ex.ParamName);
        Assert.Contains("No heading matching 'Deep' was found", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ReplaceSection_OnAHeadingNestedInAListItem_ReportsNoMatch()
    {
        // The same mismatch through the other container, pinned separately because the containers
        // are different code paths in the parser. Both shapes were measured to produce the SAME
        // symptom without the guard — a duplicated document tail — rather than one duplicating and
        // the other silently doing nothing.
        var ex = Assert.Throws<ArgumentException>(() => MarkdownEditor.ReplaceSection(
            "- # Nested\n\n# Boundary\n\ntail\n", "Nested", "x"));

        Assert.Equal("headingText", ex.ParamName);
        Assert.Contains("No heading matching 'Nested' was found", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ReplaceSection_WhenANestedHeadingShadowsARealTopLevelOne_EditsTheTopLevelOne()
    {
        // The regression this rewrite exists to close. Searching the whole document and then
        // rejecting a nested hit matched the BLOCKQUOTE's "Changed" first — it is earlier in
        // document order — and refused the whole call, making the perfectly good top-level
        // "Changed" below it uneditable with no workaround available to a caller.
        //
        // Asserted as the exact document: the blockquote must come back byte-for-byte, and the
        // edit must land under the top-level heading.
        var result = MarkdownEditor.ReplaceSection(
            "> ## Changed\n>\n> quoted\n\n## Changed\n\nreal\n", "Changed", "\nnew\n");

        Assert.Equal("> ## Changed\n>\n> quoted\n\n## Changed\n\n\nnew\n", result);
    }

    [Fact]
    public void ReplaceSection_OnATopLevelHeading_IsStillFoundAndEdited()
    {
        // The positive control for the three tests above. A search that found nothing at all would
        // satisfy all three, so this asserts the ordinary case still edits — and asserts it against
        // an exact document, so "did not throw" is not mistaken for "did the right thing".
        var result = MarkdownEditor.ReplaceSection(
            "## Boundary\n\nold\n", "Boundary", "\nnew\n");

        Assert.Equal("## Boundary\n\n\nnew\n", result);
    }

    [Fact]
    public void ReplaceSection_ALinkReferenceDefinitionRightAfterTheHeading_SurvivesTheReplacement()
    {
        // Pins the one limitation the <remarks> names, so the documented claim cannot drift into
        // being false. A link reference definition is consumed into the parser's link registry and
        // leaves no block, so `bodyStart` — the next BLOCK's start offset — lands past it and it
        // ends up in the kept prefix.
        var kept = MarkdownEditor.ReplaceSection(
            "## Changed\n\n[x]: http://example.com\n\nbody\n\n## Next\n\ntail\n", "Changed", "\nnew\n");

        Assert.Equal(
            "## Changed\n\n[x]: http://example.com\n\n\nnew\n## Next\n\ntail\n", kept);

        // The counterpart, which is what makes the behaviour position-dependent rather than simply
        // "definitions are preserved": the same definition after other content is inside the
        // replaced range and goes.
        var removed = MarkdownEditor.ReplaceSection(
            "## Changed\n\nbody\n\n[x]: http://example.com\n\n## Next\n\ntail\n", "Changed", "\nnew\n");

        Assert.Equal("## Changed\n\n\nnew\n## Next\n\ntail\n", removed);
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
