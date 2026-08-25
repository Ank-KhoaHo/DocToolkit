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
    public void Inspect_ReportsFootnotes()
    {
        var report = DocxToPdfPreflight.Inspect(WithFootnote());

        var finding = Assert.Single(report.Findings, f => f.Code == "Footnote");
        Assert.Equal(1, finding.Count);
        Assert.Equal(DocxToPdfRisk.Known, finding.Risk);
        Assert.True(report.HasFindings);
    }

    [Fact]
    public void Inspect_ReportsNestedTables()
    {
        var report = DocxToPdfPreflight.Inspect(WithNestedTable());

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
    public void Inspect_DoesNotReportAFootnotesPartThatHoldsOnlySeparators()
    {
        // Word writes the separator (id -1) and continuation separator (id 0) into any document it
        // has touched, whether or not the author wrote a footnote. Counting every Footnote element
        // would report almost every real document - the loudest possible false positive.
        var report = DocxToPdfPreflight.Inspect(WithSeparatorsOnly());

        Assert.DoesNotContain(report.Findings, f => f.Code == "Footnote");
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
    public void Inspect_DoesNotReportContentControlsOrTextBoxes()
    {
        // Both were MEASURED to survive the render. Reporting them would be crying wolf, and
        // `w:sdt` is named as at-risk in the issue this feature came from - so this test pins the
        // measurement against the assumption.
        Assert.Empty(DocxToPdfPreflight.Inspect(WithContentControl()).Findings);
        Assert.Empty(DocxToPdfPreflight.Inspect(WithTextBox()).Findings);
    }

    // ---- the loss is real, and these FAIL IF THE RENDERER IMPROVES ---------------------------

    [Fact]
    public void TheFootnoteLossIsReal_AndThisFailsIfTheRendererStartsCarryingThem()
    {
        byte[] docx = WithFootnote();
        string pdf = string.Join(" ", PdfEditor.ExtractText(DocxToPdfConverter.Convert(docx)));

        // The sibling proves the document rendered at all. Without it, an empty PDF would satisfy
        // the assertion below and this test would "pass" while proving nothing.
        Assert.Contains(Sibling, pdf, StringComparison.Ordinal);
        Assert.DoesNotContain("FOOTTOKEN", pdf, StringComparison.Ordinal);
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
        using var source = new MemoryStream(WithFootnote());

        var report = await DocxToPdfPreflight.InspectAsync(source);

        Assert.Single(report.Findings, f => f.Code == "Footnote");
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
        var finding = Assert.Single(DocxToPdfPreflight.Inspect(WithFootnote()).Findings);

        Assert.Equal("Footnote", finding.Code);
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
}
