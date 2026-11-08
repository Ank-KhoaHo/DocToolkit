using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// Covers <see cref="DocxToPdfPreflight"/> — the inventory of constructs a document carries that
/// <see cref="DocxToPdfConverter"/> may not represent.
///
/// <b>Three kinds of test here, and the third is the one that keeps this honest.</b>
///
/// <list type="number">
/// <item>the construct is present and IS reported;</item>
/// <item>the construct is absent and is NOT reported — an inventory that fires on clean input is
/// worse than none, because a caller learns to ignore it;</item>
/// <item><b>the loss is real.</b> Each <c>Known</c> finding has a test asserting the content really
/// is missing from the rendered PDF. If a future OfficeIMO learns to render nested tables, that test
/// FAILS — which is the signal to demote the finding or drop it.</item>
/// </list>
///
/// Since OfficeIMO 3.5 the nested table is the only <c>Known</c> finding: footnotes and content
/// controls render, so their tests here assert that they render and are NOT reported.
///
/// That third kind is the whole staleness defence. The construct list cannot be derived from
/// anything — measured 2026-08-25, all three candidate sources in the graph describe consuming a PDF
/// rather than producing one — so it is hand-written, and this repository has watched six
/// hand-maintained lists go stale. A list whose entries each carry a failing-on-improvement test
/// cannot drift into over-reporting unnoticed.
///
/// <b>Every fixture carries a sibling paragraph.</b> Without it, "the token is missing from the PDF"
/// cannot be told apart from "the whole render came out empty", and an earlier pass of this
/// measurement was exactly that ambiguity.
/// </summary>
public class DocxToPdfPreflightTests
{
    private const string Sibling = "SIBLINGOK";

    // ---- reported when present ---------------------------------------------------------------

    [Fact]
    public void Inspect_ReportsNestedTables()
    {
        var report = DocxToPdfPreflight.Inspect(WithNestedTable());

        Assert.True(report.HasFindings);
        var finding = Assert.Single(report.Findings, f => f.Code == "NestedTable");
        Assert.Equal(1, finding.Count);
        Assert.Equal(DocxToPdfRisk.Known, finding.Risk);
    }

    // ---- NOT reported when absent ------------------------------------------------------------

    [Fact]
    public void Inspect_OnAPlainDocument_ReportsNothing()
    {
        // The control that stops this suite passing vacuously, and the property that keeps the
        // report worth reading: it must be silent on a document with none of these constructs.
        var report = DocxToPdfPreflight.Inspect(
            DocxEditor.Create([DocxBlock.Paragraph("nothing special here")]));

        Assert.Empty(report.Findings);
        Assert.False(report.HasFindings);
    }

    [Fact]
    public void Inspect_DoesNotReportAnOrdinaryTable()
    {
        // A table is fine. Only a table INSIDE a cell is the reported case, and confusing the two
        // would fire on a large share of real documents.
        byte[] docx = DocxEditor.Create([DocxBlock.Paragraph("before")]);
        var report = DocxToPdfPreflight.Inspect(WithFlatTable());

        Assert.DoesNotContain(report.Findings, f => f.Code == "NestedTable");
        Assert.Empty(DocxToPdfPreflight.Inspect(docx).Findings);
    }

    [Fact]
    public void Inspect_DoesNotReportABodyLevelContentControlOrATextBox()
    {
        // Both were MEASURED to survive the render. Reporting them would be crying wolf, and
        // `w:sdt` is named as at-risk in the issue this feature came from - so this test pins the
        // measurement against the assumption.
        Assert.Empty(DocxToPdfPreflight.Inspect(WithContentControl()).Findings);
        Assert.Empty(DocxToPdfPreflight.Inspect(WithTextBox()).Findings);
    }

    // ---- the loss is real, and these FAIL IF THE RENDERER IMPROVES ---------------------------

    [Fact]
    public void AFootnoteNowRenders_AndTheFindingWasRemovedBecauseOfIt()
    {
        // 3.5 renders a footnote whose reference run lacks the style; 3.2.6 lost it.
        string pdf = string.Join(" ", PdfEditor.ExtractText(DocxToPdfConverter.Convert(WithFootnote())));

        Assert.Contains(Sibling, pdf, StringComparison.Ordinal);
        Assert.Contains("FOOTTOKEN", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void TheNestedTableLossIsReal_AndThisFailsIfTheRendererStartsCarryingIt()
    {
        byte[] docx = WithNestedTable();
        string pdf = string.Join(" ", PdfEditor.ExtractText(DocxToPdfConverter.Convert(docx)));

        Assert.Contains(Sibling, pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("NESTTOKEN", pdf, StringComparison.Ordinal);
    }

    /// <summary>
    /// The mirror of the two tests above, and the reason the EXCLUSIONS are falsifiable too.
    ///
    /// Content controls and text boxes are deliberately not reported, on the grounds that they were
    /// measured to survive the render. That is a claim, and a claim nothing re-checks is how a list
    /// goes stale in the other direction: if OfficeIMO ever stops carrying them, the exclusion
    /// becomes wrong and this fails - which is the signal to promote them to findings.
    /// </summary>
    [Fact]
    public void TheExcludedConstructsReallyDoSurviveTheRender()
    {
        string sdt = string.Join(" ", PdfEditor.ExtractText(DocxToPdfConverter.Convert(WithContentControl())));
        string box = string.Join(" ", PdfEditor.ExtractText(DocxToPdfConverter.Convert(WithTextBox())));

        Assert.Contains(Sibling, sdt, StringComparison.Ordinal);
        Assert.Contains("SDTTOKEN", sdt, StringComparison.Ordinal);

        Assert.Contains(Sibling, box, StringComparison.Ordinal);
        Assert.Contains("BOXTOKEN", box, StringComparison.Ordinal);
    }

    [Fact]
    public void AStyledFootnoteSurvivesTheRender()
    {
        byte[] docx = WithStyledFootnote();
        string pdf = string.Join(" ", PdfEditor.ExtractText(DocxToPdfConverter.Convert(docx)));

        Assert.Contains(Sibling, pdf, StringComparison.Ordinal);
        Assert.Contains("STYLEDFOOTTOKEN", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ADefinitionOnlyStyledFootnoteNowRenders()
    {
        string pdf = string.Join(" ", PdfEditor.ExtractText(
            DocxToPdfConverter.Convert(WithStyledFootnoteDefinitionOnly())));

        Assert.Contains(Sibling, pdf, StringComparison.Ordinal);
        Assert.Contains("DEFONLYTOKEN", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void AReferenceOnlyStyledFootnoteSurvivesTheRender()
    {
        byte[] docx = WithStyledFootnoteReferenceOnly();
        string pdf = string.Join(" ", PdfEditor.ExtractText(DocxToPdfConverter.Convert(docx)));

        Assert.Contains(Sibling, pdf, StringComparison.Ordinal);
        Assert.Contains("REFONLYTOKEN", pdf, StringComparison.Ordinal);
    }

    // ---- guards -------------------------------------------------------------------------------

    [Fact]
    public void Inspect_RefusesNullAndEmptyByTheParameterItDeclares()
    {
        Assert.Equal("docx", Assert.Throws<ArgumentNullException>(
            () => DocxToPdfPreflight.Inspect(null!)).ParamName);
        Assert.Equal("docx", Assert.Throws<ArgumentException>(
            () => DocxToPdfPreflight.Inspect([])).ParamName);
    }

    [Fact]
    public async Task InspectAsync_ReadsTheSameStateAndLeavesTheStreamOpen()
    {
        using var source = new MemoryStream(WithNestedTable());

        var report = await DocxToPdfPreflight.InspectAsync(source);

        Assert.Single(report.Findings, f => f.Code == "NestedTable");
        source.Position = 0;
        Assert.True(source.ReadByte() >= 0, "the caller's stream must not be closed");
    }

    [Fact]
    public void Inspect_WrapsAnUnreadableDocument()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxToPdfPreflight.Inspect([1, 2, 3, 4]));
        Assert.NotNull(ex.InnerException);
    }

    // ---- degenerate documents: report nothing, raise nothing ---------------------------------

    /// <summary>
    /// Four documents that are openable but missing a part this scanner reaches through. All four
    /// are constructible - measured, not assumed - and a preflight that THREW on one would be worse
    /// than useless: the whole point is to run it over documents somebody else authored, which is
    /// exactly the population that contains malformed files.
    /// </summary>
    [Theory]
    [InlineData("no main document part")]
    [InlineData("a main part with no Document")]
    [InlineData("a Document with no Body")]
    [InlineData("a FootnotesPart holding no Footnotes")]
    public void Inspect_OnADegenerateDocument_ReportsNothingRatherThanThrowing(string shape)
    {
        var report = DocxToPdfPreflight.Inspect(Degenerate(shape));

        Assert.Empty(report.Findings);
        Assert.False(report.HasFindings);
    }

    [Fact]
    public void EveryFinding_CarriesTextAConsumerCanActOn()
    {
        // Deliberately NOT pinning the wording - a message is prose and a test asserting its exact
        // words fails on every rewrite while proving nothing. What must hold is that a caller
        // rendering the report to a human gets something in every field, and that Count is not
        // silently disagreeing with the message it appears in.
        var finding = Assert.Single(DocxToPdfPreflight.Inspect(WithNestedTable()).Findings);

        Assert.Equal("NestedTable", finding.Code);
        Assert.False(string.IsNullOrWhiteSpace(finding.Construct));
        Assert.False(string.IsNullOrWhiteSpace(finding.Message));
        Assert.Contains(finding.Count.ToString(), finding.Message, StringComparison.Ordinal);
    }

    // ---- fixtures ------------------------------------------------------------------------------

    private static byte[] Build(Action<MainDocumentPart, Body> build)
    {
        using var ms = new MemoryStream();
        using (var d = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = d.AddMainDocumentPart();
            var body = new Body();
            build(main, body);
            body.Append(new Paragraph(new Run(new Text(Sibling))));
            main.Document = new Document(body);
            main.Document.Save();
        }
        return ms.ToArray();
    }

    /// <summary>
    /// A properly formed footnote: separator, continuation separator and a reference mark.
    ///
    /// <b>The separators matter.</b> A first version of this fixture omitted them, and Word may
    /// legitimately ignore a footnotes part without them - so "absent from the PDF" would have said
    /// nothing about the renderer. Measured both ways: lost either way.
    /// </summary>
    private static byte[] WithFootnote() => Build((main, body) =>
    {
        var part = main.AddNewPart<FootnotesPart>();
        part.Footnotes = new Footnotes(
            new Footnote(new Paragraph(new Run(new SeparatorMark())))
            { Id = -1, Type = FootnoteEndnoteValues.Separator },
            new Footnote(new Paragraph(new Run(new ContinuationSeparatorMark())))
            { Id = 0, Type = FootnoteEndnoteValues.ContinuationSeparator },
            new Footnote(new Paragraph(
                new Run(new FootnoteReferenceMark()),
                new Run(new Text(" FOOTTOKEN"))))
            { Id = 1 });
        part.Footnotes.Save();

        body.Append(new Paragraph(
            new Run(new Text("body ")),
            new Run(new FootnoteReference { Id = 1 })));
    });

    /// <summary>
    /// A footnote authored the way <see cref="DocxEditor.AddFootnote"/> (and Word itself) actually
    /// shape one — unlike <see cref="WithFootnote"/> above, which is deliberately the UNSTYLED
    /// shape neither of those produces. Built through the real API rather than hand-rolled XML, on
    /// purpose: A94 found <see cref="WithFootnote"/>'s own lack of
    /// <c>RunStyle="FootnoteReference"</c> mattered to the renderer (3.2.6 lost such a footnote;
    /// OfficeIMO 3.5 renders both), and a second hand-built fixture would only re-risk the same
    /// mistake.
    /// </summary>
    private static byte[] WithStyledFootnote() =>
        DocxEditor.AddFootnote(
            DocxEditor.Create([DocxBlock.Paragraph($"body {{{{note}}}} here {Sibling}")]),
            "{{note}}", "STYLEDFOOTTOKEN");

    /// <summary>
    /// The footnote's own DEFINITION carries <c>ParagraphStyleId="FootnoteText"</c>; the BODY's
    /// reference run carries no style at all. A94's own probe (Case E) measured this shape lost —
    /// this pins that the finding condition reads the body reference run, not the definition, since
    /// a predicate reading the wrong half of the document would pass every other fixture in this
    /// file and only this one would catch it.
    /// </summary>
    private static byte[] WithStyledFootnoteDefinitionOnly() => Build((main, body) =>
    {
        var part = main.AddNewPart<FootnotesPart>();
        part.Footnotes = new Footnotes(
            new Footnote(new Paragraph(new Run(new SeparatorMark())))
            { Id = -1, Type = FootnoteEndnoteValues.Separator },
            new Footnote(new Paragraph(new Run(new ContinuationSeparatorMark())))
            { Id = 0, Type = FootnoteEndnoteValues.ContinuationSeparator },
            new Footnote(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "FootnoteText" }),
                new Run(new FootnoteReferenceMark()),
                new Run(new Text(" DEFONLYTOKEN"))))
            { Id = 1 });
        part.Footnotes.Save();

        body.Append(new Paragraph(
            new Run(new Text("body ")),
            new Run(new FootnoteReference { Id = 1 })));
    });

    /// <summary>
    /// The mirror of <see cref="WithStyledFootnoteDefinitionOnly"/>: the footnote's own definition
    /// carries no styling at all; only the BODY's reference run carries
    /// <c>RunStyle="FootnoteReference"</c>. A94's own probe (Case G) measured this shape survives —
    /// the definition's own styling turned out never to matter.
    /// </summary>
    private static byte[] WithStyledFootnoteReferenceOnly() => Build((main, body) =>
    {
        var part = main.AddNewPart<FootnotesPart>();
        part.Footnotes = new Footnotes(
            new Footnote(new Paragraph(new Run(new SeparatorMark())))
            { Id = -1, Type = FootnoteEndnoteValues.Separator },
            new Footnote(new Paragraph(new Run(new ContinuationSeparatorMark())))
            { Id = 0, Type = FootnoteEndnoteValues.ContinuationSeparator },
            new Footnote(new Paragraph(
                new Run(new FootnoteReferenceMark()),
                new Run(new Text(" REFONLYTOKEN"))))
            { Id = 1 });
        part.Footnotes.Save();

        body.Append(new Paragraph(
            new Run(new Text("body ")),
            new Run(
                new RunProperties(new RunStyle { Val = "FootnoteReference" }),
                new FootnoteReference { Id = 1 })));
    });

    /// <summary>The separators alone — what Word writes into a document with no author footnotes.</summary>
    private static byte[] WithSeparatorsOnly() => Build((main, body) =>
    {
        var part = main.AddNewPart<FootnotesPart>();
        part.Footnotes = new Footnotes(
            new Footnote(new Paragraph(new Run(new SeparatorMark())))
            { Id = -1, Type = FootnoteEndnoteValues.Separator },
            new Footnote(new Paragraph(new Run(new ContinuationSeparatorMark())))
            { Id = 0, Type = FootnoteEndnoteValues.ContinuationSeparator });
        part.Footnotes.Save();

        body.Append(new Paragraph(new Run(new Text("no footnotes of its own"))));
    });

    private static Table Table(params OpenXmlElement[] cellContent) => new(
        new TableProperties(new TableWidth { Type = TableWidthUnitValues.Auto }),
        new TableGrid(new GridColumn()),
        new TableRow(new TableCell(cellContent)));

    private static byte[] WithNestedTable() => Build((main, body) =>
        body.Append(Table(Table(new Paragraph(new Run(new Text("NESTTOKEN")))), new Paragraph())));

    private static byte[] WithFlatTable() => Build((main, body) =>
        body.Append(Table(new Paragraph(new Run(new Text("FLATTOKEN"))))));

    private static byte[] WithContentControl() => Build((main, body) =>
        body.Append(new SdtBlock(
            new SdtProperties(new SdtAlias { Val = "field" }),
            new SdtContentBlock(new Paragraph(new Run(new Text("SDTTOKEN")))))));

    private static byte[] WithTextBox() => Build((main, body) =>
    {
        var textbox = new DocumentFormat.OpenXml.Vml.TextBox(
            new TextBoxContent(new Paragraph(new Run(new Text("BOXTOKEN")))));
        var shape = new DocumentFormat.OpenXml.Vml.Shape(textbox)
        { Style = "width:200pt;height:50pt", Id = "s1" };
        body.Append(new Paragraph(new Run(new Picture(shape))));
    });

    private static byte[] Degenerate(string shape)
    {
        using var ms = new MemoryStream();
        using (var d = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            if (shape == "no main document part") { }
            else
            {
                var main = d.AddMainDocumentPart();
                if (shape == "a Document with no Body")
                {
                    main.Document = new Document();
                    main.Document.Save();
                }
                else if (shape == "a FootnotesPart holding no Footnotes")
                {
                    main.Document = new Document(new Body(new Paragraph(new Run(new Text("x")))));
                    main.AddNewPart<FootnotesPart>();
                    main.Document.Save();
                }
            }
        }
        return ms.ToArray();
    }
    // ---- A76: a content control INSIDE a table cell -------------------------------------------

    private static byte[] WithControlInACell()
    {
        var cell = new TableCell(ControlHolding(DocxFixtures.P(Text_("INCELLTOKEN"))));
        return DocxFixtures.Build(
            DocxFixtures.Tbl(new TableRow(cell)),
            DocxFixtures.P(Text_(Sibling)));
    }

    private static Run Text_(string text) =>
        new(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

    private static SdtBlock ControlHolding(OpenXmlElement inner) => new(
        new SdtProperties(new SdtAlias { Val = "c" }, new Tag { Val = "c" }),
        new SdtContentBlock(inner));

    [Fact]
    public void AControlInsideATableCellNowRenders()
    {
        string pdf = string.Join(" ", PdfEditor.ExtractText(
            DocxToPdfConverter.Convert(WithControlInACell())));

        Assert.Contains(Sibling, pdf, StringComparison.Ordinal);
        Assert.Contains("INCELLTOKEN", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void AFootnoteAndAControlInACellAreNotReported_BecauseBothRenderOn35()
    {
        // Every shape the two removed findings keyed on, and the shapes they excluded.
        // "Not reported" alone would also pass for a broken detector, so the last line proves
        // the report still speaks.
        byte[][] fixtures =
        [
            WithFootnote(), WithStyledFootnote(), WithStyledFootnoteDefinitionOnly(),
            WithStyledFootnoteReferenceOnly(), WithSeparatorsOnly(), WithControlInACell(),
            CellWrappedInControl(), RowWrappedInControl(),
        ];

        foreach (var docx in fixtures)
        {
            Assert.DoesNotContain(DocxToPdfPreflight.Inspect(docx).Findings,
                f => f.Code is "Footnote" or "ControlInCell");
            Assert.Empty(DocxToPdfPreflight.Inspect(docx).Findings);
        }

        Assert.Single(DocxToPdfPreflight.Inspect(WithNestedTable()).Findings);
    }

    [Fact]
    public void ABodyLevelControlIsStillNotReported_BecauseItStillRenders()
    {
        // The boundary. A67 excluded content controls after measuring a body-level one surviving,
        // and that measurement still holds - so the new finding must be about the CELL case
        // specifically. Reporting every control would make the preflight cry wolf on the common
        // shape, which is how a report stops being read.
        var report = DocxToPdfPreflight.Inspect(WithContentControl());

        Assert.Empty(report.Findings);
    }

    [Fact]
    public void ANestedTableWrappedInAControlIsStillCountedAsNested()
    {
        // Found by the code review of A77: the nested-table walk read c.Elements<Table>(), so a
        // table wrapped in a w:sdt was invisible to it. Measured before the fix - the same shape
        // reported 1 finding unwrapped and 0 wrapped.
        //
        // Both fixtures are asserted here rather than only the wrapped one, because a walk that
        // reported ZERO for both would satisfy a single-sided test.
        // Named locals rather than one nested expression: the first version of this fixture was
        // seven parentheses deep and did not compile, which is a fair warning about how readable
        // it would have been.
        static Table InnerTable() =>
            DocxFixtures.Tbl(new TableRow(new TableCell(DocxFixtures.P(Text_("inner")))));

        static byte[] OuterHolding(OpenXmlElement nested) => DocxFixtures.Build(
            DocxFixtures.Tbl(new TableRow(new TableCell(DocxFixtures.P(Text_("outer")), nested))));

        byte[] wrapped = OuterHolding(ControlHolding(InnerTable()));
        byte[] plain = OuterHolding(InnerTable());

        Assert.Equal(1, DocxToPdfPreflight.Inspect(plain).Findings
            .Single(f => f.Code == "NestedTable").Count);
        Assert.Equal(1, DocxToPdfPreflight.Inspect(wrapped).Findings
            .Single(f => f.Code == "NestedTable").Count);
    }
    // ---- the shapes a code review found nothing covered ----------------------------------------

    private static SdtCell CellControl(TableCell inner) =>
        new(new SdtProperties(new Tag { Val = "c" }), new SdtContentCell(inner));

    private static SdtRow RowControl(TableRow inner) =>
        new(new SdtProperties(new Tag { Val = "r" }), new SdtContentRow(inner));

    private static Paragraph TextBoxHolding(OpenXmlElement inner) =>
        new(new Run(new DocumentFormat.OpenXml.Wordprocessing.Picture(
            new DocumentFormat.OpenXml.Vml.Shape(
                new DocumentFormat.OpenXml.Vml.TextBox(new TextBoxContent(inner)))
            {
                Id = "TextBox1",
                Type = "#_x0000_t202",
                Style = "position:absolute;width:200pt;height:80pt",
            })));

    [Fact]
    public void AControlInATextBoxTableCellIsNOTReported_BecauseItRenders()
    {
        // THE false positive. Descendants<TableCell> walks into w:txbxContent, so a control in a
        // table inside a text box was reported as lost - while its text reaches the PDF perfectly.
        // Measured both halves here, because "not reported" alone would also be satisfied by a
        // counter that had stopped working.
        byte[] docx = DocxFixtures.Build(
            TextBoxHolding(DocxFixtures.Tbl(new TableRow(new TableCell(
                ControlHolding(DocxFixtures.P(Text_("BOXCTL"))))))),
            DocxFixtures.P(Text_(Sibling)));

        Assert.Empty(DocxToPdfPreflight.Inspect(docx).Findings);

        string pdf = string.Join(" ", PdfEditor.ExtractText(DocxToPdfConverter.Convert(docx)));
        Assert.Contains(Sibling, pdf, StringComparison.Ordinal);
        Assert.Contains("BOXCTL", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void ANestedTableInATextBoxIsNOTReported_BecauseItRenders()
    {
        // The same false positive on the OLDER finding, one line away in the source. Pre-dates the
        // content-control work; fixed with it because both walks now share the same cell filter.
        var inner = DocxFixtures.Tbl(new TableRow(new TableCell(DocxFixtures.P(Text_("BOXNEST")))));
        var outer = DocxFixtures.Tbl(new TableRow(new TableCell(DocxFixtures.P(Text_("o")), inner)));

        byte[] docx = DocxFixtures.Build(TextBoxHolding(outer), DocxFixtures.P(Text_(Sibling)));

        Assert.Empty(DocxToPdfPreflight.Inspect(docx).Findings);

        string pdf = string.Join(" ", PdfEditor.ExtractText(DocxToPdfConverter.Convert(docx)));
        Assert.Contains(Sibling, pdf, StringComparison.Ordinal);
        Assert.Contains("BOXNEST", pdf, StringComparison.Ordinal);
    }

    private static byte[] CellWrappedInControl() => DocxFixtures.Build(
        DocxFixtures.Tbl(new TableRow(
            new TableCell(DocxFixtures.P(Text_("LEFTOK"))),
            CellControl(new TableCell(DocxFixtures.P(Text_("RIGHTNOW")))))),
        DocxFixtures.P(Text_(Sibling)));

    private static byte[] RowWrappedInControl() => DocxFixtures.Build(
        DocxFixtures.Tbl(
            new TableRow(new TableCell(DocxFixtures.P(Text_("LEFTOK")))),
            RowControl(new TableRow(new TableCell(DocxFixtures.P(Text_("RIGHTNOW")))))),
        DocxFixtures.P(Text_(Sibling)));

    [Theory]
    [InlineData("cell")]
    [InlineData("row")]
    public void ACellOrRowWrappedInAControlNowRenders(string position)
    {
        byte[] docx = position == "cell" ? CellWrappedInControl() : RowWrappedInControl();

        string pdf = string.Join(" ", PdfEditor.ExtractText(DocxToPdfConverter.Convert(docx)));
        Assert.Contains(Sibling, pdf, StringComparison.Ordinal);
        Assert.Contains("LEFTOK", pdf, StringComparison.Ordinal);
        Assert.Contains("RIGHTNOW", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void NestedTablesAreCountedByDEPTH_NotBySweepingEveryTable()
    {
        // Kills `SelectMany(ContentControls.Tables)` -> `SelectMany(c => c.Descendants<Table>())`,
        // the tempting shorthand, which survived because every other fixture is only two levels
        // deep - where the two walks agree. Three levels separates them: the cell-scoped
        // Descendants sweep counts the innermost table twice, once from each ancestor cell.
        var l3 = DocxFixtures.Tbl(new TableRow(new TableCell(DocxFixtures.P(Text_("L3")))));
        var l2 = DocxFixtures.Tbl(new TableRow(new TableCell(DocxFixtures.P(Text_("L2")), l3)));
        var l1 = DocxFixtures.Tbl(new TableRow(new TableCell(DocxFixtures.P(Text_("L1")), l2)));

        byte[] docx = DocxFixtures.Build(l1, DocxFixtures.P(Text_(Sibling)));

        Assert.Equal(2, Assert.Single(DocxToPdfPreflight.Inspect(docx).Findings).Count);
    }
}
