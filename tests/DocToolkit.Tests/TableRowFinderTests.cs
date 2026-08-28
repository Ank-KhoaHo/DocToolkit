using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit.Tests;

/// <summary>
/// Discovery is tested on its own, not merely through <c>FillRows</c>, because its one dangerous
/// behaviour — refusing to reach into nested tables — fails silently when it regresses. An
/// end-to-end test would show a wrong document; these show which rows were chosen.
/// </summary>
public class TableRowFinderTests
{
    private static IReadOnlyList<TableRow> FindIn(byte[] docx, string marker)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        return TableRowFinder.Find(doc.MainDocumentPart!.Document!.Body!, marker);
    }

    [Fact]
    public void Find_ReturnsOnlyTheRowsHoldingTheMarker()
    {
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.Row(DocxFixtures.R("Description")),
            DocxFixtures.Row(DocxFixtures.R("{{item.Desc}}")),
            DocxFixtures.Row(DocxFixtures.R("Total"))));

        var found = FindIn(docx, "{{item.");

        Assert.Single(found);
        Assert.Contains("{{item.Desc}}", found[0].InnerText);
    }

    [Fact]
    public void Find_ReturnsNothingWhenNoRowHoldsTheMarker()
    {
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.Row(DocxFixtures.R("Description")),
            DocxFixtures.Row(DocxFixtures.R("{{payment.Total}}"))));

        Assert.Empty(FindIn(docx, "{{item."));
    }

    [Fact]
    public void Find_DoesNotTreatAContainerRowAsATemplateRow()
    {
        // The outer row's own text has no marker; only the nested table's row does.
        // Descendants<TableRow>() would return both, and the outer row would then be
        // cloned per record with the whole inner table inside it.
        var inner = DocxFixtures.Tbl(DocxFixtures.Row(DocxFixtures.R("{{item.Desc}}")));
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.RowOf(DocxFixtures.P(DocxFixtures.R("container")), inner)));

        var found = FindIn(docx, "{{item.");

        Assert.Single(found);
        Assert.DoesNotContain("container", found[0].InnerText);
    }

    [Fact]
    public void Find_ReturnsNestedRowsBeforeTheRowsContainingThem()
    {
        // Both the inner row and the outer row carry the marker. Innermost must come first, so a
        // nested template row is expanded before anything clones the row it sits in.
        var inner = DocxFixtures.Tbl(DocxFixtures.Row(DocxFixtures.R("{{item.Inner}}")));
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.RowOf(DocxFixtures.P(DocxFixtures.R("{{item.Outer}}")), inner)));

        var found = FindIn(docx, "{{item.");

        Assert.Equal(2, found.Count);
        Assert.Contains("{{item.Inner}}", found[0].InnerText);
        Assert.Contains("{{item.Outer}}", found[1].InnerText);
    }

    [Fact]
    public void Find_SpansSeveralTables()
    {
        var docx = DocxFixtures.Build(
            DocxFixtures.Tbl(DocxFixtures.Row(DocxFixtures.R("{{item.A}}"))),
            DocxFixtures.P(DocxFixtures.R("between the tables")),
            DocxFixtures.Tbl(DocxFixtures.Row(DocxFixtures.R("{{item.B}}"))));

        var found = FindIn(docx, "{{item.");

        Assert.Equal(2, found.Count);
    }

    /// <summary>
    /// Found 2026-08-28, while re-measuring B30's own fix: <c>OwnsMarker</c>'s inner loop over a
    /// cell's paragraphs is an <c>Any()</c>, and collapsing it to <c>All()</c> survived every other
    /// test in this file. The reason is <c>DocxFixtures.Row</c>/<c>RowOf</c> always build one
    /// paragraph per cell (<c>P(children)</c> wraps everything into a single <see
    /// cref="Paragraph"/>), so no fixture here had ever put two paragraphs in the same cell - `Any`
    /// and `All` agree on a sequence of one. Confirmed by hand: applying the mutation left all 1822
    /// tests green (this file's own suite included) until this test was added.
    ///
    /// The sibling <c>Any(cell =&gt; ...)</c> at the outer, per-CELL level looked identical in a
    /// Stryker report but is NOT a gap - <c>DocxTableContentControlTests.FillRows_...</c> already
    /// builds genuinely multi-cell rows through <c>FillRows</c> and catches it; that mutant's
    /// "Survived" reading was this repository's own documented mutation-coverage corruption (see
    /// `CLAUDE.md`, B30), not a second real gap. Hand-verified separately before writing this: the
    /// outer mutation alone still fails 2 tests; the inner one, alone, failed none.
    /// </summary>
    [Fact]
    public void Find_MatchesACellWhoseMarkerIsInOnlyOneOfSeveralParagraphs()
    {
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.RowOf(
                DocxFixtures.P(DocxFixtures.R("a heading paragraph, no marker here")),
                DocxFixtures.P(DocxFixtures.R("{{item.Desc}}")))));

        var found = FindIn(docx, "{{item.");

        Assert.Single(found);
    }
}
