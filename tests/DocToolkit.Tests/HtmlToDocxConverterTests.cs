using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocToolkit;
using Xunit;

namespace DocToolkit.Tests;

public class HtmlToDocxConverterTests
{
    private const string Html = """
        <h1>Quarterly Report</h1>
        <p>Revenue was <strong>up 12%</strong> and costs were <em>flat</em>.</p>
        <table border="1"><tr><th>Region</th><th>Total</th></tr>
        <tr><td>North</td><td>1200</td></tr></table>
        <ul><li>First</li><li>Second</li></ul>
        <p><a href="https://example.com/report">Full report</a></p>
        """;

    [Fact]
    public async Task ConvertAsync_ProducesAValidDocxPackage()
    {
        var bytes = await HtmlToDocxConverter.ConvertAsync(Html);

        Assert.NotEmpty(bytes);
        // A .docx is a ZIP: it must start with the local file header magic "PK\x03\x04".
        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, bytes.Take(4).ToArray());
    }

    [Fact]
    public async Task ConvertAsync_PreservesStructureAndFormatting()
    {
        var bytes = await HtmlToDocxConverter.ConvertAsync(Html);

        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var body = doc.MainDocumentPart!.Document.Body!;

        Assert.True(body.Descendants<Paragraph>().Count() >= 4);
        Assert.Single(body.Descendants<Table>());
        Assert.Equal(2, body.Descendants<TableRow>().Count());
        Assert.NotEmpty(body.Descendants<Bold>());
        Assert.NotEmpty(body.Descendants<Italic>());
        Assert.NotEmpty(body.Descendants<Hyperlink>());
        Assert.Contains("Quarterly Report", body.InnerText);
    }

    [Fact]
    public async Task ConvertAsync_RejectsNullHtml()
    {
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToDocxConverter.ConvertAsync(null!));
    }
}
