using System.Collections.Generic;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class DocxEditorServiceTests
{
    [Fact]
    public async Task ReplaceText_SubstitutesPlaceholders()
    {
        var sut = new DocxEditorService();
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<p>Dear {{name}}, your balance is {{balance}}.</p>");

        var edited = sut.ReplaceText(docx, new Dictionary<string, string>
        {
            ["{{name}}"] = "Contoso Ltd",
            ["{{balance}}"] = "4,250.00",
        });

        var text = sut.ExtractText(edited);
        Assert.Contains("Contoso Ltd", text);
        Assert.Contains("4,250.00", text);
        Assert.DoesNotContain("{{name}}", text);
    }

    [Fact]
    public async Task ExtractText_WithHeadersAndFooters_MatchesTheStaticMethod()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync("<p>Body text.</p>");
        var sut = new DocxEditorService();

        Assert.Equal(
            DocxEditor.ExtractText(docx, includeHeadersAndFooters: true),
            sut.ExtractText(docx, includeHeadersAndFooters: true));
    }

    [Fact]
    public void ReplaceText_RejectsNullReplacements()
    {
        var sut = new DocxEditorService();

        Assert.Throws<ArgumentNullException>(() => sut.ReplaceText(Array.Empty<byte>(), null!));
    }
}
