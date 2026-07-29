using DocToolkit;
using Xunit;

namespace DocToolkit.Tests;

public class DocxToPdfConverterTests
{
    [Fact]
    public async Task Convert_ProducesAPdfContainingTheSourceText()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<h1>Invoice INV-42</h1><p>Total due: 18,100.00</p>");

        var pdf = DocxToPdfConverter.Convert(docx);

        Assert.True(PdfProbe.IsPdf(pdf));
        var text = PdfProbe.ExtractText(pdf);
        Assert.Contains("Invoice INV-42", text);
        Assert.Contains("18,100.00", text);
    }

    [Fact]
    public async Task Convert_PaginatesLongDocuments()
    {
        var rows = string.Concat(Enumerable.Range(1, 60).Select(i =>
            $"<tr><td>Line item {i} with a reasonably long description</td><td>{i * 950}</td></tr>"));
        var html = $"<h1>Big</h1><table border=\"1\">{rows}</table><p>END-MARKER</p>";

        var pdf = DocxToPdfConverter.Convert(await HtmlToDocxConverter.ConvertAsync(html));

        Assert.True(PdfProbe.PageCount(pdf) > 1, "expected the document to span multiple pages");
        Assert.Contains("END-MARKER", PdfProbe.ExtractText(pdf));
        Assert.DoesNotContain(PdfProbe.TextYPositions(pdf), y => y < 0);
    }

    [Fact]
    public void Convert_RejectsEmptyInput()
    {
        Assert.Throws<ArgumentException>(() => DocxToPdfConverter.Convert(Array.Empty<byte>()));
    }
}
