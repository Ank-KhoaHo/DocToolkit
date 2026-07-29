using DocToolkit;
using Xunit;

namespace DocToolkit.Tests;

public class DocxEditorTests
{
    [Fact]
    public async Task ReplaceText_SubstitutesPlaceholders()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<p>Dear {{name}}, your balance is {{balance}}.</p>");

        var edited = DocxEditor.ReplaceText(docx, new Dictionary<string, string>
        {
            ["{{name}}"] = "Contoso Ltd",
            ["{{balance}}"] = "4,250.00",
        });

        var text = DocxEditor.ExtractText(edited);
        Assert.Contains("Contoso Ltd", text);
        Assert.Contains("4,250.00", text);
        Assert.DoesNotContain("{{name}}", text);
        Assert.DoesNotContain("{{balance}}", text);
    }

    [Fact]
    public async Task ReplaceText_LeavesTheDocumentOpenable()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync("<p>Hello {{who}}</p>");

        var edited = DocxEditor.ReplaceText(docx,
            new Dictionary<string, string> { ["{{who}}"] = "world" });

        // Still a valid package, and still renders.
        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, edited.Take(4).ToArray());
        Assert.Contains("world", PdfProbe.ExtractText(DocxToPdfConverter.Convert(edited)));
    }

    [Fact]
    public async Task ExtractText_ReturnsDocumentText()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync("<h1>Title</h1><p>Body copy.</p>");
        var text = DocxEditor.ExtractText(docx);

        Assert.Contains("Title", text);
        Assert.Contains("Body copy.", text);
    }
}
