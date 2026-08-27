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

    // Delegated to DocxFixtures rather than redeclared. The only builders that live here are
    // the three w:sdt wrappers, which no other test file needs.
    private static Paragraph P(string text) =>
        DocxFixtures.P(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));

    private static TableRow Row(params string[] cells) => DocxFixtures.RowOfText(cells);

    private static Table Tbl(params OpenXmlElement[] rows) => DocxFixtures.Tbl(rows);

    private static SdtBlock BlockControl(OpenXmlElement inner) => new(
        new SdtProperties(new SdtAlias { Val = "c" }, new Tag { Val = "c" }),
        new SdtContentBlock(inner));

    private static SdtRow RowControl(TableRow inner) => new(
        new SdtProperties(new SdtAlias { Val = "r" }, new Tag { Val = "r" }),
        new SdtContentRow(inner));

    private static SdtCell CellControl(TableCell inner) => new(
        new SdtProperties(new SdtAlias { Val = "x" }, new Tag { Val = "x" }),
        new SdtContentCell(inner));

    private static byte[] Doc(params OpenXmlElement[] blocks) => DocxFixtures.Build(blocks);

    private static void AssertSchemaValid(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Empty(new OpenXmlValidator().Validate(doc).Select(e => e.Description).Take(1));
    }

    private static string Flat(IReadOnlyList<IReadOnlyList<string>> table) =>
        string.Join(" / ", table.Select(row => string.Join("|", row)));

    // ---- the fixtures are what they claim to be ------------------------------------------------

    /// <summary>Every distinct document SHAPE this file builds, named once.</summary>
    /// <remarks>
    /// A code review pointed out that the class comment claimed every fixture was validated while
    /// this test checked three of about a dozen — the double-nested control, both nested-table
    /// shapes and every FillRows fixture went unchecked. A claim nothing verifies is the defect
    /// this repository keeps recording, so the shapes are enumerated here and the validator runs
    /// over all of them.
    ///
    /// <para><b>The theory passes a NAME and builds the document inside the test</b>, rather than
    /// passing the bytes through <c>MemberData</c>. A <c>byte[]</c> in a theory case puts the whole
    /// array into the case's display name, and Stryker's coverage analysis then cannot match those
    /// cases to the tests it ran — it warned <i>"mutation tests may be inaccurate"</i> 90 times in
    /// one run, which quietly degrades the one tool that measures whether these tests discriminate
    /// at all.</para>
    /// </remarks>
    public static TheoryData<string> EveryShapeName() =>
    [
        "wrapped table",
        "wrapped table, twice nested",
        "wrapped row",
        "wrapped cell",
        "paragraph wrapped in a cell",
        "table nested in a cell",
        "wrapped and ordinary together",
        "template row, wrapped table",
        "template row, wrapped row",
        "template marker in a wrapped cell",
    ];

    private static byte[] ShapeNamed(string shape) => shape switch
    {
        "wrapped table" => Doc(BlockControl(Tbl(Row("W"))), P("after")),

        "wrapped table, twice nested" => Doc(BlockControl(BlockControl(Tbl(Row("W")))), P("after")),

        "wrapped row" => Doc(Tbl(Row("Name"), RowControl(Row("W"))), P("after")),

        "wrapped cell" => Doc(Tbl(new TableRow(CellControl(new TableCell(P("W"))))), P("after")),

        "paragraph wrapped in a cell" => Doc(Tbl(new TableRow(new TableCell(BlockControl(P("W"))))), P("after")),

        "table nested in a cell" => Doc(Tbl(new TableRow(new TableCell(P("cell"), Tbl(Row("inner"))))), P("after")),

        "wrapped and ordinary together" => Doc(BlockControl(Tbl(Row("W"))), Tbl(Row("A")), P("after")),

        "template row, wrapped table" => Doc(BlockControl(Tbl(Row("Name", "Qty"), Row("{{item.Name}}", "{{item.Qty}}"))), P("after")),

        "template row, wrapped row" => Doc(Tbl(Row("Name", "Qty"), RowControl(Row("{{item.Name}}", "{{item.Qty}}"))), P("after")),

        _ => Doc(Tbl(Row("Name", "Qty"), new TableRow(CellControl(new TableCell(P("{{item.Name}}"))), new TableCell(P("fixed")))), P("after")),
    };

    [Theory]
    [MemberData(nameof(EveryShapeName))]
    public void EveryFixtureShapeIsSchemaValid(string shape)
    {
        byte[] docx = ShapeNamed(shape);
        // A fixture Word would reject proves nothing about Word, and this repository has already
        // had a whole family of measurements invalidated by a fixture that was not what it claimed
        // - see CLAUDE.md on hand-built SdtBlock content controls.
        Assert.NotEmpty(shape);
        AssertSchemaValid(docx);
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
    public void FillRows_ReachesATemplateRowInsideANestedTable()
    {
        // RENAMED after a code review measured what it actually checks. It used to be called
        // ...ExpandsTheInnermostRowFirstWhenTablesNest and its comment claimed the ordering
        // guarantee - but the marker sits only in the INNER table, so TableRowFinder returns
        // exactly one row, and a one-element list has no order. Adding found.Reverse() left it
        // green. What it really proves is that a nested template row is reachable through a cell,
        // which is worth keeping under a name that says so.
        //
        // The ordering guarantee is checked by the sibling below, and by the pre-existing
        // TableRowFinderTests.Find_ReturnsNestedRowsBeforeTheRowsContainingThem.
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
    // ---- found by code review, measured before fixing --------------------------------------------

    [Fact]
    public void FillRows_FindsAMarkerInAParagraphWrappedInsideACell()
    {
        // The FOURTH wrapper position, and the one the first pass missed: w:tc > w:sdt > w:p.
        // Collect unwrapped SdtBlock when looking for nested TABLES in a cell, but OwnsMarker still
        // read the cell's direct-child paragraphs - so this refused with the same misleading
        // "the marker must appear inside a table cell" message A77 was filed against.
        //
        // It is the shape Word writes when a Rich Text control is inserted into an empty cell.
        var wrappedMarker = new TableCell(BlockControl(P("{{item.Name}}")));
        var templateRow = new TableRow(wrappedMarker, new TableCell(P("fixed")));

        byte[] docx = Doc(Tbl(Row("Name", "Qty"), templateRow), P("after"));

        // ExtractText already reads it, which is the disagreement this whole row exists to close.
        Assert.Contains("{{item.Name}}", DocxEditor.ExtractText(docx), StringComparison.Ordinal);

        string text = DocxEditor.ExtractText(DocxEditor.FillRows(docx, "item", TwoRecords));
        Assert.Contains("Widget", text, StringComparison.Ordinal);
        Assert.Contains("Gadget", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{item.", text, StringComparison.Ordinal);
    }

    [Fact]
    public void FillRows_WithNoRecords_StillRemovesTheTable_EvenWhenTheTemplateRowIsWrapped()
    {
        // ExpandRow took template.Parent to find the table to remove. For a wrapped row that parent
        // is SdtContentRow, not Table, so the removal silently stopped happening on the very path
        // this change opened - leaving an empty frame behind.
        //
        // This is not a preference: DocxEditor's shipped XML doc states the removal as behaviour
        // ("an empty frame left on the page reads worse than rendering nothing") and that text
        // renders on the API site. The change would have made a published claim false.
        byte[] wrappedRow = Doc(Tbl(RowControl(Row("{{item.Name}}"))), P("after"));
        byte[] plainRow = Doc(Tbl(Row("{{item.Name}}")), P("after"));

        Assert.Equal(0, DocxEditor.TableCount(DocxEditor.FillRows(plainRow, "item", [])));
        Assert.Equal(0, DocxEditor.TableCount(DocxEditor.FillRows(wrappedRow, "item", [])));
    }

    [Fact]
    public void FillRows_WithNoRecords_KeepsATableThatStillHasOtherRows()
    {
        // The control on the test above. Without it, "remove the table" could be implemented as
        // "always remove the table" and both assertions there would still pass.
        byte[] docx = Doc(Tbl(Row("Name"), RowControl(Row("{{item.Name}}"))), P("after"));

        byte[] emptied = DocxEditor.FillRows(docx, "item", []);

        Assert.Equal(1, DocxEditor.TableCount(emptied));
        Assert.Equal("Name", Flat(DocxEditor.ReadTable(emptied, 0)));
    }

    [Fact]
    public void FillRows_WithNoRecords_KeepsATableWhoseOnlyRemainingRowIsItselfWRAPPED()
    {
        // Found by sabotage, not by design. The test above leaves an ORDINARY row behind, so a
        // direct-child row count still finds it and asking emptiness the old way passed.
        //
        // Here the survivor is wrapped: a direct-child count sees ZERO rows, calls the table empty
        // and deletes a row the caller can plainly see. That is silent data loss introduced by the
        // very fix that removes empty frames, and only asking through ContentControls avoids it.
        byte[] docx = Doc(Tbl(RowControl(Row("Kept")), Row("{{item.Name}}")), P("after"));

        byte[] emptied = DocxEditor.FillRows(docx, "item", []);

        Assert.Equal(1, DocxEditor.TableCount(emptied));
        Assert.Equal("Kept", Flat(DocxEditor.ReadTable(emptied, 0)));
    }

    [Fact]
    public void TableRowFinder_ReturnsTheInnermostRowFirst_WithBothRowsOwningTheMarker()
    {
        // The ordering test in this file could not discriminate: its fixture put the marker only in
        // the inner table, so Find returned ONE row and a one-element list has no order. Reversing
        // the finder left it green. Both rows own a marker here, so the order is observable.
        //
        // Order matters because a nested template row must be expanded BEFORE anything clones the
        // row it sits in - otherwise the clones carry an unexpanded template.
        var inner = Tbl(Row("{{item.Name}}"));
        var outer = Tbl(new TableRow(new TableCell(P("outer {{item.Qty}}"), inner)));

        byte[] docx = Doc(outer, P("after"));
        string text = DocxEditor.ExtractText(DocxEditor.FillRows(docx, "item", TwoRecords));

        // Two records against two nested template rows: every record must appear, and no template
        // may survive anywhere - which is exactly what expanding in the wrong order breaks.
        Assert.Contains("Widget", text, StringComparison.Ordinal);
        Assert.Contains("Gadget", text, StringComparison.Ordinal);
        Assert.DoesNotContain("{{item.", text, StringComparison.Ordinal);
    }
    [Fact]
    public void FillRows_KeepsTheGeneratedRowsInsideTheControlTheTemplateWasIn()
    {
        // The STRUCTURE, not just the text. Both wrapped-row FillRows tests above assert only that
        // "Widget" and "Gadget" appear, which any structure preserving the text satisfies - the
        // presence-only shape CLAUDE.md flags as a repeat failure here. A code review measured what
        // actually comes out and it is worth pinning deliberately rather than leaving to accident.
        //
        // Measured: the table's children are tblPr, tblGrid, tr, sdt - and that ONE w:sdt holds
        // BOTH generated rows. Schema-valid, and ReadTable reads all three rows in order.
        //
        // The alternative - cloning the whole w:sdt once per record - would give each row its own
        // control carrying the SAME tag and alias, which is its own ambiguity. Which one Word
        // prefers is NOT verified here: there is no Word on this machine, and this repository's
        // rule is that an unverified claim is worse than an absent one. What is verified is that
        // the output is schema-valid, reads back correctly, and stays inside the control the
        // author put the row in rather than escaping to the table. If a real Word document ever
        // shows this is wrong, this test is the record of what was chosen and why.
        byte[] docx = Doc(Tbl(Row("Name", "Qty"), RowControl(Row("{{item.Name}}", "{{item.Qty}}"))), P("after"));

        byte[] filled = DocxEditor.FillRows(docx, "item", TwoRecords);

        AssertSchemaValid(filled);
        Assert.Equal("Name|Qty / Widget|2 / Gadget|5", Flat(DocxEditor.ReadTable(filled, 0)));

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single();

        // No row escaped the control: the table still has exactly one direct-child w:tr, the header.
        Assert.Single(table.Elements<TableRow>());

        var control = Assert.Single(table.Elements<SdtRow>());
        Assert.Equal(2, control.SdtContentRow!.Elements<TableRow>().Count());
    }
}
