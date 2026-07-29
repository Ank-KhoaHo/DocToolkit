using DocToolkit;
using Xunit;

namespace DocToolkit.Tests;

public class HtmlToPdfConverterTests
{
    [Fact]
    public async Task ConvertAsync_ProducesAPdfFromHtmlInOneCall()
    {
        var pdf = await HtmlToPdfConverter.ConvertAsync(
            "<h1>Statement</h1><p>Balance: 4,250.00</p>");

        Assert.True(PdfProbe.IsPdf(pdf));
        var text = PdfProbe.ExtractText(pdf);
        Assert.Contains("Statement", text);
        Assert.Contains("4,250.00", text);
    }

    [Fact]
    public async Task ConvertAsync_MatchesTheTwoStepPipeline()
    {
        const string html = "<h1>Same Input</h1><p>Same output text.</p>";

        var direct = await HtmlToPdfConverter.ConvertAsync(html);
        var stepwise = DocxToPdfConverter.Convert(await HtmlToDocxConverter.ConvertAsync(html));

        // Byte equality is not guaranteed (timestamps/ids), but the rendered text must match.
        Assert.Equal(PdfProbe.ExtractText(stepwise), PdfProbe.ExtractText(direct));
    }

    [Fact]
    public async Task ConvertAsync_RejectsNullHtml()
    {
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToPdfConverter.ConvertAsync(null!));
    }
}
