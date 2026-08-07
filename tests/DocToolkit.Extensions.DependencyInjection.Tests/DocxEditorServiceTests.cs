using System.Collections.Generic;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
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
    public void ExtractText_WithHeadersAndFooters_MatchesTheStaticMethod()
    {
        var docx = DocxWithHeaderAndFooter("Body text.", "Page header", "Page footer");
        var sut = new DocxEditorService();

        Assert.Equal(
            DocxEditor.ExtractText(docx, includeHeadersAndFooters: true),
            sut.ExtractText(docx, includeHeadersAndFooters: true));

        // The old fixture (built via HtmlToDocxConverter) had zero header/footer parts, so this
        // assertion was previously impossible to fail even if the bool parameter were dropped or
        // inverted. This fixture genuinely has both, so these two lines are what actually prove the
        // parameter threads through correctly.
        Assert.Contains("Page header", sut.ExtractText(docx, includeHeadersAndFooters: true));
        Assert.DoesNotContain("Page header", sut.ExtractText(docx, includeHeadersAndFooters: false));
    }

    [Fact]
    public void ReplaceText_RejectsNullReplacements()
    {
        var sut = new DocxEditorService();

        Assert.Throws<ArgumentNullException>(() => sut.ReplaceText(Array.Empty<byte>(), null!));
    }

    [Fact]
    public async Task ExtractTextAsync_MatchesTheStaticMethod()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync("<p>Body copy.</p>");
        var sut = new DocxEditorService();

        using var expectedSource = new MemoryStream(docx);
        using var actualSource = new MemoryStream(docx);

        var expected = await DocxEditor.ExtractTextAsync(expectedSource);
        var actual = await sut.ExtractTextAsync(actualSource);

        Assert.Equal(expected, actual);
    }

    [Fact]
    public async Task ExtractTextAsync_WithHeadersAndFooters_MatchesTheStaticMethod()
    {
        var docx = DocxWithHeaderAndFooter("Body text.", "Page header", "Page footer");
        var sut = new DocxEditorService();

        using var expectedSource = new MemoryStream(docx);
        using var actualSource = new MemoryStream(docx);

        var expected = await DocxEditor.ExtractTextAsync(expectedSource, includeHeadersAndFooters: true);
        var actual = await sut.ExtractTextAsync(actualSource, includeHeadersAndFooters: true);

        Assert.Equal(expected, actual);
        Assert.Contains("Page header", actual);
    }

    [Fact]
    public async Task ReplaceTextAsync_MatchesTheStaticMethod()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync("<p>Dear {{name}}.</p>");
        var sut = new DocxEditorService();
        var replacements = new Dictionary<string, string> { ["{{name}}"] = "Contoso Ltd" };

        using var expectedSource = new MemoryStream(docx);
        using var expected = new MemoryStream();
        await DocxEditor.ReplaceTextAsync(expectedSource, replacements, expected);

        using var actualSource = new MemoryStream(docx);
        using var actual = new MemoryStream();
        await sut.ReplaceTextAsync(actualSource, replacements, actual);

        Assert.Equal(expected.ToArray(), actual.ToArray());
    }

    [Fact]
    public async Task ExtractTextAsync_HonorsCancellation()
    {
        var sut = new DocxEditorService();
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        using var source = new MemoryStream();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => sut.ExtractTextAsync(source, cts.Token));
    }

    private static byte[] DocxWithHeaderAndFooter(string bodyText, string headerText, string footerText)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var body = new Body(new Paragraph(new Run(new Text(bodyText))));
            mainPart.Document = new Document(body);

            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(new Paragraph(new Run(new Text(headerText))));
            var headerRelId = mainPart.GetIdOfPart(headerPart);

            var footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(new Paragraph(new Run(new Text(footerText))));
            var footerRelId = mainPart.GetIdOfPart(footerPart);

            body.Append(new SectionProperties(
                new HeaderReference { Type = HeaderFooterValues.Default, Id = headerRelId },
                new FooterReference { Type = HeaderFooterValues.Default, Id = footerRelId }));

            mainPart.Document.Save();
        }
        return ms.ToArray();
    }
}
