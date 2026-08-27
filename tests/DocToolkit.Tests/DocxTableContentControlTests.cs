using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit.Tests;

/// <summary>
/// A77 — <see cref="DocxEditor.TableCount"/>, <see cref="DocxEditor.ReadTable"/> and
/// <see cref="DocxEditor.FillRows"/> must see a table, row or cell wrapped in a content control.
/// </summary>
/// <remarks>
/// <b>The defect this closes was a DISAGREEMENT, not a single wrong number.</b> <c>ExtractText</c>
/// was fixed in #393 and reads all of these; <c>ReadTable</c> read none of them. Two readers
/// answering differently about the same document is worse than either answer alone, so several
/// tests here assert the two AGREE rather than asserting a value twice.
///
/// <para>Every fixture is checked by <c>OpenXmlValidator</c> and carries a sibling paragraph. The
/// sibling is load-bearing: without it "the token is missing" and "the whole read came back empty"
/// are the same observation, and an earlier measurement in this family was exactly that.</para>
/// </remarks>
public class DocxTableContentControlTests
{
    // ---- fixtures ------------------------------------------------------------------------------

    private static Paragraph P(string text) =>
        new(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

    private static TableRow Row(params string[] cells)
    {
        var row = new TableRow();
        foreach (string cell in cells) row.Append(new TableCell(P(cell)));
        return row;
    }

    /// <summary>
    /// A table with its <c>w:tblGrid</c>. Without one the validator rejects the first <c>w:tr</c>
    /// as an unexpected child — an earlier version of this repository's helper built schema-invalid
    /// tables that every test happily passed against.
    /// </summary>
    private static Table Tbl(params OpenXmlElement[] rows)
    {
        // OpenXmlElement rather than TableRow: a w:sdt at row level is an SdtRow, which is not a
        // TableRow at all - which is the entire shape of the defect being closed here.
        var table = new Table(new TableProperties(),
                              new TableGrid(new GridColumn(), new GridColumn()));
        foreach (var row in rows) table.Append(row);
        return table;
    }

    private static SdtBlock BlockControl(OpenXmlElement inner) => new(
        new SdtProperties(new SdtAlias { Val = "c" }, new Tag { Val = "c" }),
        new SdtContentBlock(inner));

    private static SdtRow RowControl(TableRow inner) => new(
        new SdtProperties(new SdtAlias { Val = "r" }, new Tag { Val = "r" }),
        new SdtContentRow(inner));

    private static SdtCell CellControl(TableCell inner) => new(
        new SdtProperties(new SdtAlias { Val = "x" }, new Tag { Val = "x" }),
        new SdtContentCell(inner));

    private static byte[] Doc(params OpenXmlElement[] blocks)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var body = new Body();
            foreach (var block in blocks) body.Append(block);
            doc.AddMainDocumentPart().Document = new Document(body);
        }

        return ms.ToArray();
    }

    private static void AssertSchemaValid(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Empty(new OpenXmlValidator().Validate(doc).Select(e => e.Description).Take(1));
    }

    private static string Flat(IReadOnlyList<IReadOnlyList<string>> table) =>
        string.Join(" / ", table.Select(row => string.Join("|", row)));

    // ---- the fixtures are what they claim to be ------------------------------------------------

    [Fact]
    public void EveryFixtureHereIsSchemaValid()
    {
        // The control on the whole file. A fixture Word would reject proves nothing about Word,
        // and this repository has already had a family of measurements invalidated by a fixture
        // that was not what it claimed - see CLAUDE.md on hand-built SdtBlock content controls.
        AssertSchemaValid(Doc(BlockControl(Tbl(Row("W"))), P("after")));
        AssertSchemaValid(Doc(Tbl(Row("Name"), RowControl(Row("W"))), P("after")));
        AssertSchemaValid(Doc(Tbl(new TableRow(CellControl(new TableCell(P("W"))))), P("after")));
    }

    // ---- TableCount ----------------------------------------------------------------------------

    [Fact]
    public void TableCount_SeesATableWrappedInAContentControl()
    {
        // Measured before the fix: 0.
        Assert.Equal(1, DocxEditor.TableCount(Doc(BlockControl(Tbl(Row("W"))), P("after"))));
    }

    [Fact]
    public void TableCount_CountsWrappedAndOrdinaryTablesTogether()
    {
        byte[] docx = Doc(Tbl(Row("A")), BlockControl(Tbl(Row("W"))), P("after"));

        Assert.Equal(2, DocxEditor.TableCount(docx));
    }

    [Fact]
    public void TableCount_StillIgnoresATableNestedInsideACell()
    {
        // The half that must NOT change. A nested table is part of its cell's text, not an entry of
        // its own - so unwrapping controls must not become a Descendants walk. CLAUDE.md records
        // what Descendants cost twice: deleted text-box content, and a nested row swept into its
        // container's expansion.
        var outer = Tbl(new TableRow(new TableCell(P("cell"), Tbl(Row("inner")))));

        Assert.Equal(1, DocxEditor.TableCount(Doc(outer, P("after"))));
    }

    [Fact]
    public void TableCount_SeesATableWrappedInTwoNestedControls()
    {
        // Word nests controls, so one unwrap is not enough.
        byte[] docx = Doc(BlockControl(BlockControl(Tbl(Row("W")))), P("after"));

        Assert.Equal(1, DocxEditor.TableCount(docx));
    }

    // ---- ReadTable ------------------------------------------------------------------------------

    [Fact]
    public void ReadTable_ReadsATableWrappedInAContentControl()
    {
        // Measured before the fix: ArgumentOutOfRangeException, "The document has 0 table(s)".
        byte[] docx = Doc(BlockControl(Tbl(Row("W", "9"))), P("after"));

        Assert.Equal("W|9", Flat(DocxEditor.ReadTable(docx, 0)));
    }

    [Fact]
    public void ReadTable_IndexZeroIsTheFIRSTTable_EvenWhenItIsWrapped()
    {
        // THE case that makes this breaking, and the one worth reading twice. Before the fix this
        // returned "A" - the table that is physically SECOND - because the wrapped one was
        // invisible and the index silently slid past it. A confidently wrong answer, not an error.
        byte[] docx = Doc(BlockControl(Tbl(Row("W"))), Tbl(Row("A")), P("after"));

        Assert.Equal("W", Flat(DocxEditor.ReadTable(docx, 0)));
        Assert.Equal("A", Flat(DocxEditor.ReadTable(docx, 1)));
    }

    [Fact]
    public void ReadTable_KeepsARowWrappedInAContentControl()
    {
        // NOT in the filed row - found by measuring it. This is the worse defect of the two: a
        // TableCount of 0 is visibly wrong, while a table that comes back one row short looks
        // exactly like data. Measured before the fix: 2 rows, the wrapped one silently dropped.
        byte[] docx = Doc(
            Tbl(Row("Name", "Qty"), RowControl(Row("WRAPPED", "9")), Row("plain", "1")),
            P("after"));

        Assert.Equal("Name|Qty / WRAPPED|9 / plain|1", Flat(DocxEditor.ReadTable(docx, 0)));
    }

    [Fact]
    public void ReadTable_KeepsACellWrappedInAContentControl()
    {
        // One level further in. A dropped cell also shifts every cell beside it, so the columns
        // stop lining up - which is why the assertion is the exact row rather than a count.
        var row = new TableRow(
            new TableCell(P("a")),
            CellControl(new TableCell(P("WRAPPED"))),
            new TableCell(P("c")));

        Assert.Equal("a|WRAPPED|c", Flat(DocxEditor.ReadTable(Doc(Tbl(row), P("after")), 0)));
    }

    [Fact]
    public void ReadTable_StillDoesNotFlattenATableNestedInACell()
    {
        // The nesting guarantee again, this time on content rather than on the count: the inner
        // table's text belongs to the cell, and the outer table still has exactly one row.
        var outer = Tbl(new TableRow(new TableCell(P("cell"), Tbl(Row("inner")))));
        var read = DocxEditor.ReadTable(Doc(outer, P("after")), 0);

        Assert.Single(read);
        Assert.Contains("inner", read[0][0], StringComparison.Ordinal);
    }

    [Fact]
    public void ReadTable_OutOfRangeStillNamesTheRealCount()
    {
        byte[] docx = Doc(BlockControl(Tbl(Row("W"))), P("after"));

        var ex = Assert.Throws<ArgumentOutOfRangeException>(() => DocxEditor.ReadTable(docx, 5));
        Assert.Contains("1 table(s)", ex.Message, StringComparison.Ordinal);
    }

    // ---- the disagreement that was the real defect ----------------------------------------------

    [Theory]
    [InlineData("wrapped table")]
    [InlineData("wrapped row")]
    [InlineData("wrapped cell")]
    public void ExtractTextAndReadTable_AgreeAboutWhatTheDocumentHolds(string shape)
    {
        // ExtractText was fixed in #393 and read all three; ReadTable read none. Asserting the two
        // AGREE is stronger than asserting each value, because it cannot be satisfied by fixing one
        // reader and forgetting the other - which is exactly how this gap was created.
        byte[] docx = shape switch
        {
            "wrapped table" => Doc(BlockControl(Tbl(Row("TOKEN"))), P("after")),
            "wrapped row" => Doc(Tbl(Row("head"), RowControl(Row("TOKEN"))), P("after")),
            _ => Doc(Tbl(new TableRow(CellControl(new TableCell(P("TOKEN"))))), P("after")),
        };

        Assert.Contains("TOKEN", DocxEditor.ExtractText(docx), StringComparison.Ordinal);
        Assert.Contains("TOKEN", Flat(DocxEditor.ReadTable(docx, 0)), StringComparison.Ordinal);
    }

    // ---- FillRows --------------------------------------------------------------------------------

    private static readonly Dictionary<string, string>[] TwoRecords =
    [
        new() { ["Name"] = "Widget", ["Qty"] = "2" },
        new() { ["Name"] = "Gadget", ["Qty"] = "5" },
    ];

    [Fact]
    public void FillRows_ExpandsATemplateRowInsideAWrappedTable()
    {
        // Measured before the fix: DocumentConversionException saying "the marker must appear
        // inside a table cell" - which it did. The message sent the caller to fix something they
        // had already done correctly, which is worse than the refusal itself.
        byte[] docx = Doc(
            BlockControl(Tbl(Row("Name", "Qty"), Row("{{item.Name}}", "{{item.Qty}}"))),
            P("after"));

        string text = DocxEditor.ExtractText(DocxEditor.FillRows(docx, "item", TwoRecords));

        Assert.Contains("Widget", text, StringComparison.Ordinal);
        Assert.Contains("Gadget", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{item.", text, StringComparison.Ordinal);
    }

    [Fact]
    public void FillRows_ExpandsATemplateRowThatIsItselfWrapped()
    {
        byte[] docx = Doc(
            Tbl(Row("Name", "Qty"), RowControl(Row("{{item.Name}}", "{{item.Qty}}"))),
            P("after"));

        string text = DocxEditor.ExtractText(DocxEditor.FillRows(docx, "item", TwoRecords));

        Assert.Contains("Widget", text, StringComparison.Ordinal);
        Assert.Contains("Gadget", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{item.", text, StringComparison.Ordinal);
    }

    [Fact]
    public void FillRows_StillRefusesWhenTheMarkerIsGenuinelyNotInATable()
    {
        // The control. Without it, a fix that found rows ANYWHERE would pass every test above while
        // destroying the refusal - and that refusal is the one telling a caller their template is
        // wrong. The message must still be the honest one, because now it is true.
        byte[] docx = Doc(P("{{item.Name}} sitting in a bare paragraph"));

        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxEditor.FillRows(docx, "item", TwoRecords));
        Assert.Contains("inside a table cell", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void FillRows_StillExpandsTheInnermostRowFirstWhenTablesNest()
    {
        // TableRowFinder returns rows innermost-first so a nested template row is expanded before
        // anything clones the row it sits in. Unwrapping controls must not disturb that ordering.
        var inner = Tbl(Row("{{item.Name}}"));
        var outer = Tbl(new TableRow(new TableCell(P("outer"), inner)));

        string text = DocxEditor.ExtractText(DocxEditor.FillRows(Doc(outer, P("after")), "item", TwoRecords));

        Assert.Contains("Widget", text, StringComparison.Ordinal);
        Assert.Contains("Gadget", text, StringComparison.Ordinal);
    }
    [Fact]
    public void FillRows_FindsAMarkerSittingInAWrappedCELL()
    {
        // Found by sabotage, not by design: reverting OwnsMarker to a ChildElements walk survived
        // the whole suite, because every other FillRows test wraps the TABLE or the ROW and leaves
        // the cells ordinary. A marker inside a cell-level control is the third place a template
        // can hide, and it was the only one nothing measured.
        // The marker sits ONLY in the wrapped cell. A first version put {{item.Qty}} in an
        // ordinary cell of the same row, and OwnsMarker found the row through THAT - so the test
        // passed against the reverted walk and proved nothing. The sibling cell carries plain text.
        var templateRow = new TableRow(
            CellControl(new TableCell(P("{{item.Name}}"))),
            new TableCell(P("fixed")));

        byte[] docx = Doc(Tbl(Row("Name", "Qty"), templateRow), P("after"));

        string text = DocxEditor.ExtractText(DocxEditor.FillRows(docx, "item", TwoRecords));

        Assert.Contains("Widget", text, StringComparison.Ordinal);
        Assert.Contains("Gadget", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{item.", text, StringComparison.Ordinal);
    }
}
