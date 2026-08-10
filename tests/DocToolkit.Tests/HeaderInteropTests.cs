using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// The two things headers touch that were not obviously safe, both measured by spike before this
/// feature was designed rather than assumed afterwards.
/// </summary>
public class HeaderInteropTests
{
    private static readonly DocxBlock[] Blocks = { DocxBlock.Paragraph("Body.") };

    /// <summary>
    /// ReplaceText walks header parts, and a generated header is an ordinary header part - so a
    /// letterhead can carry {{customer}}. But ReplaceText splices runs, and a field contains runs.
    /// </summary>
    [Fact]
    public void ReplaceTextFillsAHeaderTokenAndLeavesThePageFieldIntact()
    {
        var page = PageSetup.A4.WithHeader(DocxHeader.Of(
            HeaderAlignment.Left,
            DocxHeaderSegment.Text("Customer: {{customer}} - Page "),
            DocxHeaderSegment.PageNumber));

        var filled = DocxEditor.ReplaceText(
            DocxEditor.Create(Blocks, page),
            new Dictionary<string, string> { ["{{customer}}"] = "Contoso Ltd" });

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var header = doc.MainDocumentPart!.HeaderParts.Single().Header!;

        Assert.Contains("Contoso Ltd", header.InnerText, StringComparison.Ordinal);
        var field = Assert.Single(header.Descendants<SimpleField>());
        Assert.Equal(" PAGE ", field.Instruction!.Value);
    }

    /// <summary>
    /// HtmlToPdfConverter pivots through DOCX. If the renderer ignored headers, a caller would get
    /// a PDF quietly missing what they asked for.
    /// </summary>
    [Fact]
    public async Task AHeaderSurvivesConversionToPdf()
    {
        var page = PageSetup.A4.WithHeader(DocxHeader.Text("HEADERMARKER"));

        var pdf = await HtmlToPdfConverter.ConvertAsync("<p>Body</p>", page);

        Assert.Contains("HEADERMARKER", PdfProbe.ExtractText(pdf), StringComparison.Ordinal);
    }

    [Fact]
    public void ExtractTextSurfacesHeaderTextOnlyWhenAsked()
    {
        var page = PageSetup.A4.WithHeader(DocxHeader.Text("HEADERMARKER"));
        var docx = DocxEditor.Create(Blocks, page);

        Assert.DoesNotContain("HEADERMARKER", DocxEditor.ExtractText(docx), StringComparison.Ordinal);
        Assert.Contains("HEADERMARKER",
            DocxEditor.ExtractText(docx, includeHeadersAndFooters: true), StringComparison.Ordinal);
    }
}
