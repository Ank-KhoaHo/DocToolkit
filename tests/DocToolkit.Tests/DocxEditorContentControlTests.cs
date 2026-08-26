using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// <see cref="DocxEditor.ExtractText(byte[])"/> against Word content controls (<c>w:sdt</c>).
///
/// <b>Every one of these used to return nothing.</b> Measured 2026-08-26: a document whose text sat
/// inside a block-level content control extracted to the empty string, while the same document
/// rendered to PDF carried the text perfectly — so the content was visibly in the document, survived
/// a render, and was absent from the one API whose job is reading a document's text. A caller
/// indexing a corpus lost every content-control field, which in template-driven documents is often
/// the only part that varies.
///
/// <b>The cause was structural rather than a typo.</b> <c>BlockText</c> switched on
/// <see cref="Paragraph"/> and <see cref="Table"/>, and a content control is a THIRD block-level
/// child — so it fell through and its whole subtree went with it. The same hole existed one and two
/// levels further in, at <see cref="SdtRow"/> and <see cref="SdtCell"/>.
///
/// <b>Each fixture carries a sibling paragraph</b>, so "the token is missing" can be told apart from
/// "extraction returned nothing at all" — without it, a broken extractor and a lost control look
/// identical.
/// </summary>
public class DocxEditorContentControlTests
{
    private const string Sibling = "SIBLINGOK";

    public static TheoryData<string, string> EveryPosition() => new()
    {
        { "a block control at body level", "BLOCKTOKEN" },
        { "a control nested in a control", "NESTTOKEN" },
        { "a control holding a table", "INTBLTOKEN" },
        { "a block control inside a table cell", "CELLTOKEN" },
        { "a control wrapping a row", "ROWTOKEN" },
        { "a control wrapping a cell", "CELLCTRLTOKEN" },
    };

    [Theory]
    [MemberData(nameof(EveryPosition))]
    public void ExtractText_ReadsAContentControlWhereverItSits(string shape, string token)
    {
        string text = DocxEditor.ExtractText(Fixture(shape));

        Assert.Contains(Sibling, text, StringComparison.Ordinal);
        Assert.Contains(token, text, StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_KeepsRowAndCellSeparatorsWhenAControlWrapsThem()
    {
        // Not just "the text is somewhere". A wrapped row that came back merged into its neighbour,
        // or a wrapped cell that lost its tab, would satisfy the theory above while producing text
        // that no longer describes the table - which is the defect ExtractText's block separators
        // exist to prevent, arriving by a different route.
        Assert.Contains($"PLAINROW\nROWTOKEN", DocxEditor.ExtractText(Fixture("a control wrapping a row")),
            StringComparison.Ordinal);
        Assert.Contains($"PLAINCELL\tCELLCTRLTOKEN", DocxEditor.ExtractText(Fixture("a control wrapping a cell")),
            StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractText_OnAnEmptyControl_ReturnsTheRestRatherThanThrowing()
    {
        // A control with no content element at all is legal, and a caller's corpus will contain one.
        byte[] docx = DocxFixtures.Build(new SdtBlock(), Paragraph(Sibling));

        Assert.Equal(Sibling, DocxEditor.ExtractText(docx));
    }

    [Fact]
    public void ExtractText_StillSeparatesOrdinaryBlocks()
    {
        // The control that stops the change above being a regression dressed as a fix: the
        // documented block separators are exactly what 0.21.0's Migrating entry promises, and this
        // asserts the same literal that entry does.
        byte[] docx = DocxFixtures.Build(Paragraph("Title"), Paragraph("Body text."));

        Assert.Equal("Title\nBody text.", DocxEditor.ExtractText(docx));
    }

    [Fact]
    public void ExtractText_DoesNotFlattenATableNestedInACell()
    {
        // The Elements-not-Descendants rule this fix had to preserve. Unwrapping controls level by
        // level keeps it; a Descendants-based fix would have passed every test above and quietly
        // flattened this one.
        var inner = DocxFixtures.Tbl(DocxFixtures.Row(DocxFixtures.R("INNER")));
        byte[] docx = DocxFixtures.Build(
            DocxFixtures.Tbl(DocxFixtures.RowOf(inner, DocxFixtures.P(DocxFixtures.R("OUTER")))),
            Paragraph(Sibling));

        string text = DocxEditor.ExtractText(docx);

        Assert.Contains("INNER", text, StringComparison.Ordinal);
        Assert.Contains("OUTER", text, StringComparison.Ordinal);
    }

    // ---- fixtures ------------------------------------------------------------------------------

    private static Paragraph Paragraph(string text) => DocxFixtures.P(DocxFixtures.R(text));

    private static SdtProperties Props(string alias)
        => new(new SdtAlias { Val = alias }, new Tag { Val = alias });

    private static SdtBlock Block(string alias, params OpenXmlElement[] content)
        => new(Props(alias), new SdtContentBlock(content));

    private static byte[] Fixture(string shape) => shape switch
    {
        "a block control at body level" =>
            DocxFixtures.Build(Block("a", Paragraph("BLOCKTOKEN")), Paragraph(Sibling)),

        "a control nested in a control" =>
            DocxFixtures.Build(Block("o", Block("i", Paragraph("NESTTOKEN"))), Paragraph(Sibling)),

        "a control holding a table" =>
            DocxFixtures.Build(
                Block("t", DocxFixtures.Tbl(DocxFixtures.Row(DocxFixtures.R("INTBLTOKEN")))),
                Paragraph(Sibling)),

        "a block control inside a table cell" =>
            DocxFixtures.Build(
                DocxFixtures.Tbl(DocxFixtures.RowOf(Block("c", Paragraph("CELLTOKEN")), Paragraph(""))),
                Paragraph(Sibling)),

        "a control wrapping a row" =>
            DocxFixtures.Build(
                WithRows(DocxFixtures.Tbl(DocxFixtures.Row(DocxFixtures.R("PLAINROW"))),
                    new SdtRow(Props("r"),
                        new SdtContentRow(DocxFixtures.Row(DocxFixtures.R("ROWTOKEN"))))),
                Paragraph(Sibling)),

        "a control wrapping a cell" =>
            DocxFixtures.Build(
                WithRows(DocxFixtures.Tbl(),
                    new TableRow(
                        new TableCell(Paragraph("PLAINCELL")),
                        new SdtCell(Props("c"),
                            new SdtContentCell(new TableCell(Paragraph("CELLCTRLTOKEN")))))),
                Paragraph(Sibling)),

        _ => throw new ArgumentOutOfRangeException(nameof(shape), shape, "no such fixture"),
    };

    /// <summary>
    /// Appends rows a <c>Tbl</c> helper cannot take, because its parameter is <c>TableRow[]</c> and
    /// an <see cref="SdtRow"/> is not one.
    /// </summary>
    private static Table WithRows(Table table, params OpenXmlElement[] rows)
    {
        foreach (OpenXmlElement row in rows) table.Append(row);
        return table;
    }
}
