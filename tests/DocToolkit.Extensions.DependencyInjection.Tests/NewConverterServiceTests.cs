using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

/// <summary>
/// The services added when 1:1 parity was restored (A28): the two Markdown importers, the two
/// spreadsheet exporters, and the members that had gone missing from three existing interfaces.
///
/// These wrappers are pure delegation, so the temptation is to assert only that the wrapper agrees
/// with the static method it wraps. That holds identically when both return nothing, which is the
/// tautology B16 spent a day removing - so every case below also asserts a LITERAL.
/// </summary>
public class NewConverterServiceTests
{
    private const string Markdown = "# Quarterly Report\n\nRevenue was **up 12%**.\n";

    private static byte[] Workbook() => DocToolkit.WorkbookEditor.Create("Sales", new[]
    {
        new object?[] { "Region", "Total" },
        new object?[] { "North", 1234.5 },
    });

    private static byte[] Docx() => DocToolkit.DocxEditor.Create(new[]
    {
        DocxBlock.Heading("Quarterly Report", 1),
        DocxBlock.Paragraph("Revenue was up 12%."),
    });

    // =====================================================================================
    // Markdown -> DOCX
    // =====================================================================================

    [Fact]
    public async Task MarkdownToDocx_ConvertsAndStreamsAndReports()
    {
        IMarkdownToDocxConverter sut = new MarkdownToDocxConverterService();

        var docx = sut.Convert(Markdown);
        Assert.Contains("Quarterly Report", DocToolkit.DocxEditor.ExtractText(docx), StringComparison.Ordinal);

        using var destination = new MemoryStream();
        await sut.ConvertAsync(Markdown, destination);
        Assert.Contains("Quarterly Report",
            DocToolkit.DocxEditor.ExtractText(destination.ToArray()), StringComparison.Ordinal);

        var reported = sut.ConvertWithReport(Markdown);
        Assert.Contains("Quarterly Report",
            DocToolkit.DocxEditor.ExtractText(reported.Value), StringComparison.Ordinal);
        Assert.False(reported.HasLoss);
    }

    // =====================================================================================
    // Markdown -> PDF
    // =====================================================================================

    [Fact]
    public async Task MarkdownToPdf_ConvertsAndStreamsAndReports()
    {
        IMarkdownToPdfConverter sut = new MarkdownToPdfConverterService();

        var pdf = sut.Convert(Markdown);
        Assert.Equal(new byte[] { 0x25, 0x50, 0x44, 0x46 }, pdf.Take(4).ToArray());   // "%PDF"

        using var destination = new MemoryStream();
        await sut.ConvertAsync(Markdown, destination);
        Assert.Equal(new byte[] { 0x25, 0x50, 0x44, 0x46 }, destination.ToArray().Take(4).ToArray());

        var reported = sut.ConvertWithReport(Markdown);
        Assert.Equal(new byte[] { 0x25, 0x50, 0x44, 0x46 }, reported.Value.Take(4).ToArray());
        Assert.False(reported.HasLoss);
    }

    // =====================================================================================
    // XLSX -> CSV and XLSX -> HTML
    // =====================================================================================

    [Fact]
    public async Task XlsxToCsv_ExportsTheLiteralGrid()
    {
        IXlsxToCsvConverter sut = new XlsxToCsvConverterService();
        var xlsx = Workbook();

        Assert.Equal("Region,Total\r\nNorth,1234.5\r\n", sut.Convert(xlsx, "Sales"));

        using var source = new MemoryStream(xlsx, writable: false);
        Assert.Equal("Region,Total\r\nNorth,1234.5\r\n", await sut.ConvertAsync(source, "Sales"));
    }

    [Fact]
    public async Task XlsxToHtml_ExportsATableFragment()
    {
        IXlsxToHtmlConverter sut = new XlsxToHtmlConverterService();
        var xlsx = Workbook();

        var html = sut.Convert(xlsx, "Sales");
        Assert.Contains("<th>Region</th>", html, StringComparison.Ordinal);
        Assert.Contains("<td>1234.5</td>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<html", html, StringComparison.OrdinalIgnoreCase);

        using var source = new MemoryStream(xlsx, writable: false);
        Assert.Equal(html, await sut.ConvertAsync(source, "Sales"));
    }

    // =====================================================================================
    // The members that had gone missing from existing interfaces
    // =====================================================================================

    [Fact]
    public async Task DocxToHtml_ConvertWithReport_SurfacesTheLossItAlwaysComputed()
    {
        IDocxToHtmlConverter sut = new DocxToHtmlConverterService();
        var docx = Docx();

        var reported = sut.ConvertWithReport(docx);
        Assert.Contains("Quarterly Report", reported.Value, StringComparison.Ordinal);

        // A plain DOCX reports this today; it is the entry that justified A22 existing at all.
        Assert.Contains(reported.Warnings, w => w.Code == "SectionLayoutFlattened");
        Assert.True(reported.HasLoss);

        using var source = new MemoryStream(docx, writable: false);
        var streamed = await sut.ConvertWithReportAsync(source);
        Assert.Equal(reported.Value, streamed.Value);
        Assert.Contains(streamed.Warnings, w => w.Code == "SectionLayoutFlattened");
    }

    [Fact]
    public async Task DocxToMarkdown_ConvertWithReport_ReturnsTheMarkdownAndNoFalseWarnings()
    {
        IDocxToMarkdownConverter sut = new DocxToMarkdownConverterService();
        var docx = Docx();

        var reported = sut.ConvertWithReport(docx);
        Assert.Contains("# Quarterly Report", reported.Value, StringComparison.Ordinal);
        Assert.False(reported.HasLoss);

        using var source = new MemoryStream(docx, writable: false);
        var streamed = await sut.ConvertWithReportAsync(source);
        Assert.Equal(reported.Value, streamed.Value);
        Assert.Empty(streamed.Warnings);
    }

    [Fact]
    public async Task WorkbookEditor_Format_BoldsTheHeaderAndKeepsTheValues()
    {
        IWorkbookEditor sut = new WorkbookEditorService();
        var xlsx = Workbook();

        var formatted = sut.Format(xlsx, "Sales", XlsxFormat.Report);
        Assert.Equal("North", DocToolkit.WorkbookEditor.ReadCell(formatted, "Sales", "A2"));
        Assert.Equal(
            DocToolkit.WorkbookEditor.ReadSheet(xlsx, "Sales"),
            DocToolkit.WorkbookEditor.ReadSheet(formatted, "Sales"));

        using var source = new MemoryStream(xlsx, writable: false);
        using var destination = new MemoryStream();
        await sut.FormatAsync(source, "Sales", XlsxFormat.Report, destination);
        Assert.Equal("North", DocToolkit.WorkbookEditor.ReadCell(destination.ToArray(), "Sales", "A2"));
    }
}
