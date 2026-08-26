using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

/// <summary>
/// The three services that landed one release behind their core classes, because the extensions
/// package builds against the PUBLISHED core: <see cref="DocxFormService"/>,
/// <see cref="DocxMailMergeService"/> and <see cref="DocxToPdfPreflightService"/>.
///
/// <b>These are pure delegation, so the risk is not wrong logic — it is a member wired to the wrong
/// thing, or silently doing nothing.</b> A test asserting only "it did not throw" misses both, and
/// this package is held at 100% coverage precisely because an uncovered member here is one nobody
/// checked was connected to anything.
///
/// So every assertion below is one a passthrough would fail. Where a pair of members can be played
/// off against each other — a strict merge against a lenient one, a document that HAS a construct
/// against one that does not — they are, because no single stub satisfies both.
/// </summary>
public class DocxTemplateServiceTests
{
    // ---- DocxFormService -------------------------------------------------------------------

    [Fact]
    public void Form_Inspect_ReadsTheControlsRatherThanReportingNone()
    {
        // A stub returning an empty report passes a no-controls test. It fails here.
        DocxFormReport report = new DocxFormService().Inspect(Form());

        Assert.Equal(2, report.Fields.Count);
        Assert.Equal("Khoa Ho", Assert.Single(report.Fields, f => f.Key == "FullName").Value.Text);
    }

    [Fact]
    public async Task Form_InspectAsync_MatchesItsByteArrayTwin()
    {
        using var source = new MemoryStream(Form());

        DocxFormReport streamed = await new DocxFormService().InspectAsync(source);

        Assert.Equal(
            new DocxFormService().Inspect(Form()).Fields.Select(f => f.Key),
            streamed.Fields.Select(f => f.Key));
    }

    [Fact]
    public void Form_Validate_ReportsIssuesRatherThanApprovingAnything()
    {
        // Asserted on BOTH outcomes, so a member hard-coded to either one fails.
        IDocxForm service = new DocxFormService();

        DocxFormValidation bad = service.Validate(Form(), new Dictionary<string, DocxFormValue>
        {
            ["FullName"] = DocxFormValue.FromText("x"),
            ["Nonexistent"] = DocxFormValue.FromText("spare"),
        });
        Assert.False(bad.IsValid);
        Assert.Contains(bad.Issues, i => i.Kind == DocxFormIssueKind.UnusedValue);

        DocxFormValidation good = service.Validate(Form(), new Dictionary<string, DocxFormValue>
        {
            ["FullName"] = DocxFormValue.FromText("x"),
            ["Plan"] = DocxFormValue.FromChoice("Team"),
        });
        Assert.True(good.IsValid);
    }

    [Fact]
    public async Task Form_ValidateAsync_MatchesItsByteArrayTwin()
    {
        var values = new Dictionary<string, DocxFormValue> { ["FullName"] = DocxFormValue.FromText("x") };
        using var source = new MemoryStream(Form());

        DocxFormValidation streamed = await new DocxFormService().ValidateAsync(source, values);

        Assert.Equal(new DocxFormService().Validate(Form(), values).IsValid, streamed.IsValid);
        Assert.False(streamed.IsValid);
    }

    [Fact]
    public void Form_Fill_WritesTheValueRatherThanReturningItsInput()
    {
        byte[] filled = new DocxFormService().Fill(Form(),
            new Dictionary<string, DocxFormValue> { ["FullName"] = DocxFormValue.FromText("Someone Else") });

        // A member returning its input unchanged passes any "did not throw" test and fails this.
        Assert.Equal("Someone Else",
            Assert.Single(new DocxFormService().Inspect(filled).Fields, f => f.Key == "FullName").Value.Text);
    }

    [Fact]
    public async Task Form_FillAsync_WritesToTheDestination()
    {
        using var source = new MemoryStream(Form());
        using var destination = new MemoryStream();

        await new DocxFormService().FillAsync(source, destination,
            new Dictionary<string, DocxFormValue> { ["FullName"] = DocxFormValue.FromText("streamed") });

        Assert.Equal("streamed", Assert.Single(
            new DocxFormService().Inspect(destination.ToArray()).Fields, f => f.Key == "FullName").Value.Text);
    }

    [Fact]
    public void Form_KeyModeReachesTheCoreCall()
    {
        // The one argument a delegating member can silently drop. Alias returns a different key, so
        // a service that ignored the parameter would return "FullName" here.
        Assert.Contains(new DocxFormService().Inspect(Form(), DocxFormKey.Alias).Fields,
            f => f.Key == "Full name");
    }

    // ---- DocxMailMergeService --------------------------------------------------------------

    [Fact]
    public void MailMerge_InspectTemplate_ReadsTheFieldNames()
    {
        DocxMailMergeTemplate template = new DocxMailMergeService().InspectTemplate(MergeTemplate());

        Assert.True(template.IsValid);
        Assert.Equal("FirstName", Assert.Single(template.FieldNames));
    }

    [Fact]
    public async Task MailMerge_InspectTemplateAsync_MatchesItsByteArrayTwin()
    {
        using var source = new MemoryStream(MergeTemplate());

        DocxMailMergeTemplate streamed = await new DocxMailMergeService().InspectTemplateAsync(source);

        Assert.Equal(
            new DocxMailMergeService().InspectTemplate(MergeTemplate()).FieldNames, streamed.FieldNames);
    }

    [Fact]
    public void MailMerge_MergeIsStrictWhileMergeWithReportIsNot()
    {
        // OPPOSITE outcomes on the same input: no single stub satisfies both, and neither does a
        // member that returns its argument.
        IDocxMailMerge service = new DocxMailMergeService();
        var nothing = new Dictionary<string, string>();

        Assert.Throws<DocumentConversionException>(() => service.Merge(MergeTemplate(), nothing));

        DocxMailMergeResult lenient = service.MergeWithReport(MergeTemplate(), nothing);
        Assert.False(lenient.Report.IsComplete);
        Assert.NotEmpty(lenient.Document);

        byte[] merged = service.Merge(MergeTemplate(),
            new Dictionary<string, string> { ["FirstName"] = "Khoa" });
        Assert.Contains("Khoa", DocxEditor.ExtractText(merged), StringComparison.Ordinal);
    }

    [Fact]
    public async Task MailMerge_TheStreamFormsMatchTheirByteArrayTwins()
    {
        var values = new Dictionary<string, string> { ["FirstName"] = "Khoa" };

        using var source = new MemoryStream(MergeTemplate());
        using var destination = new MemoryStream();
        await new DocxMailMergeService().MergeAsync(source, destination, values);
        Assert.Contains("Khoa", DocxEditor.ExtractText(destination.ToArray()), StringComparison.Ordinal);

        using var reportSource = new MemoryStream(MergeTemplate());
        using var reportDestination = new MemoryStream();
        DocxMailMergeReport report = await new DocxMailMergeService()
            .MergeWithReportAsync(reportSource, reportDestination, new Dictionary<string, string>());

        Assert.False(report.IsComplete);
        Assert.NotEmpty(reportDestination.ToArray());
    }

    // ---- DocxToPdfPreflightService ---------------------------------------------------------

    [Fact]
    public void Preflight_ReportsAConstructAndStaysSilentWithoutOne()
    {
        // Both directions. A member hard-coded to "no findings" passes the clean case and fails the
        // first; one hard-coded to a finding fails the second.
        IDocxToPdfPreflight service = new DocxToPdfPreflightService();

        DocxToPdfPreflightReport found = service.Inspect(WithANestedTable());
        Assert.Equal("NestedTable", Assert.Single(found.Findings).Code);
        Assert.True(found.HasFindings);

        DocxToPdfPreflightReport clean = service.Inspect(DocxEditor.Create([DocxBlock.Paragraph("plain")]));
        Assert.False(clean.HasFindings);
    }

    [Fact]
    public async Task Preflight_InspectAsync_MatchesItsByteArrayTwin()
    {
        using var source = new MemoryStream(WithANestedTable());

        DocxToPdfPreflightReport streamed = await new DocxToPdfPreflightService().InspectAsync(source);

        Assert.Equal(
            new DocxToPdfPreflightService().Inspect(WithANestedTable()).Findings.Select(f => f.Code),
            streamed.Findings.Select(f => f.Code));
    }

    // ---- fixtures ----------------------------------------------------------------------------

    /// <summary>
    /// A form authored through OfficeIMO's own API — hand-built <c>SdtBlock</c> markup is not a
    /// TYPED control, and measuring against one is how a whole design came to rest on an artefact.
    /// </summary>
    private static byte[] Form()
    {
        using var buffer = new MemoryStream();
        using (WordDocument document = WordDocument.Create(buffer))
        {
            document.AddParagraph().AddStructuredDocumentTag("Khoa Ho", "Full name", "FullName");
            document.AddParagraph().AddDropDownList(["Free", "Pro", "Team"], "Plan", "Plan");
            document.Save();
        }
        return buffer.ToArray();
    }

    /// <summary>A template carrying one <c>MERGEFIELD</c>, in the simple on-disk form.</summary>
    private static byte[] MergeTemplate() => Build(body => body.Append(new Paragraph(
        new Run(new Text("Dear ")),
        new SimpleField(new Run(new Text("<<FirstName>>")))
        { Instruction = " MERGEFIELD FirstName \\* MERGEFORMAT " })));

    /// <summary>A table inside a table cell — one of the preflight's two measured findings.</summary>
    private static byte[] WithANestedTable() => Build(body =>
    {
        static Table Tbl(params OpenXmlElement[] cellContent) => new(
            new TableProperties(new TableWidth { Type = TableWidthUnitValues.Auto }),
            new TableGrid(new GridColumn()),
            new TableRow(new TableCell(cellContent)));

        body.Append(Tbl(Tbl(new Paragraph(new Run(new Text("inner")))), new Paragraph()));
    });

    private static byte[] Build(Action<Body> fill)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            MainDocumentPart main = document.AddMainDocumentPart();
            var body = new Body();
            fill(body);
            body.Append(new Paragraph(new Run(new Text("sibling"))));
            main.Document = new Document(body);
            main.Document.Save();
        }
        return ms.ToArray();
    }
}
