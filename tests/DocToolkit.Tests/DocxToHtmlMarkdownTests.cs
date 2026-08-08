using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// DOCX → HTML and DOCX → Markdown.
///
/// These assert <b>structure</b>, never merely that a non-empty string came back. A test that only
/// checked for non-emptiness would pass against <see cref="DocxEditor.ExtractText(byte[])"/>'s flat
/// output — which is precisely the gap these converters exist to close, so it is the one thing the
/// assertions must be able to tell apart.
/// </summary>
public class DocxToHtmlMarkdownTests
{
    /// <summary>A 10x10 PNG, inlined so no test fixture is borrowed.</summary>
    private static byte[] Png() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAAFUlEQVR42mNk+M9QzzCKRsEoGgWjAAB8"
        + "vwHxylsYIQAAAABJRU5ErkJggg==");

    private static byte[] Structured() => DocxEditor.Create(new[]
    {
        DocxBlock.Heading("Quarterly Report", 1),
        DocxBlock.Paragraph("Revenue was up 12%."),
        DocxBlock.Table(new[] { "Region", "Total" }, new[] { new object?[] { "North", 1200 } }),
    });

    // =====================================================================================
    // HTML
    // =====================================================================================

    [Fact]
    public void Html_KeepsTheHeadingParagraphAndTable()
    {
        string html = DocxToHtmlConverter.Convert(Structured());

        Assert.Contains("<h1", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Quarterly Report", html, StringComparison.Ordinal);
        Assert.Contains("<p", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<table", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("North", html, StringComparison.Ordinal);
    }

    // The flat-text oracle: ExtractText returns the same words with none of the markup. If a future
    // change made Convert delegate to it, every "contains the text" assertion above would still
    // pass and only this would fail.
    [Fact]
    public void Html_IsStructurallyRicherThanExtractText()
    {
        byte[] docx = Structured();

        string html = DocxToHtmlConverter.Convert(docx);
        string flat = DocxEditor.ExtractText(docx);

        Assert.DoesNotContain("<", flat, StringComparison.Ordinal);
        Assert.Contains("<h1", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Html_EmbedsImagesAsDataUris_SoTheOutputIsSelfContained()
    {
        byte[] docx = DocxEditor.Create(new[] { DocxBlock.Image(Png(), 40, 40, "logo") });

        string html = DocxToHtmlConverter.Convert(docx);

        Assert.Contains("data:image/png;base64,", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Html_RejectsNullAndEmptyUnwrapped()
    {
        Assert.Throws<ArgumentNullException>(() => DocxToHtmlConverter.Convert(null!));
        Assert.Throws<ArgumentException>(() => DocxToHtmlConverter.Convert(Array.Empty<byte>()));
    }

    [Fact]
    public void Html_WrapsAFailureInDocumentConversionException()
    {
        Assert.Throws<DocumentConversionException>(
            () => DocxToHtmlConverter.Convert(new byte[] { 1, 2, 3, 4 }));
    }

    [Fact]
    public async Task HtmlAsync_MatchesTheByteArrayOverload()
    {
        byte[] docx = Structured();
        using var source = new MemoryStream(docx);

        string fromStream = await DocxToHtmlConverter.ConvertAsync(source);

        Assert.Equal(DocxToHtmlConverter.Convert(docx), fromStream);
    }

    // =====================================================================================
    // Markdown
    // =====================================================================================

    [Fact]
    public void Markdown_KeepsTheHeadingParagraphAndTable()
    {
        string md = DocxToMarkdownConverter.Convert(Structured());

        Assert.Contains("# Quarterly Report", md, StringComparison.Ordinal);
        Assert.Contains("Revenue was up 12%.", md, StringComparison.Ordinal);
        Assert.Contains("| --- |", md, StringComparison.Ordinal);
        Assert.Contains("North", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_IsStructurallyRicherThanExtractText()
    {
        byte[] docx = Structured();

        string md = DocxToMarkdownConverter.Convert(docx);
        string flat = DocxEditor.ExtractText(docx);

        Assert.DoesNotContain("# ", flat, StringComparison.Ordinal);
        Assert.Contains("# Quarterly Report", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_EmbedsImagesAsDataUris_SoTheOutputIsSelfContained()
    {
        byte[] docx = DocxEditor.Create(new[] { DocxBlock.Image(Png(), 40, 40, "logo") });

        string md = DocxToMarkdownConverter.Convert(docx);

        Assert.Contains("data:image/png;base64,", md, StringComparison.Ordinal);
    }

    [Fact]
    public void Markdown_RejectsNullAndEmptyUnwrapped()
    {
        Assert.Throws<ArgumentNullException>(() => DocxToMarkdownConverter.Convert(null!));
        Assert.Throws<ArgumentException>(() => DocxToMarkdownConverter.Convert(Array.Empty<byte>()));
    }

    [Fact]
    public void Markdown_WrapsAFailureInDocumentConversionException()
    {
        Assert.Throws<DocumentConversionException>(
            () => DocxToMarkdownConverter.Convert(new byte[] { 1, 2, 3, 4 }));
    }

    [Fact]
    public async Task MarkdownAsync_MatchesTheByteArrayOverload()
    {
        byte[] docx = Structured();
        using var source = new MemoryStream(docx);

        string fromStream = await DocxToMarkdownConverter.ConvertAsync(source);

        Assert.Equal(DocxToMarkdownConverter.Convert(docx), fromStream);
    }

    // =====================================================================================
    // Self-containment: what actually guards it, and what cannot be tested.
    //
    // A test asserting that TextExportOptions sets EmbedImagesAsBase64 / ImageExportMode.Base64
    // was written here and DELETED, because it was vacuous. Removing both explicit assignments
    // and re-running left it passing: the upstream defaults are already Base64, so nothing
    // observable distinguishes "we set it" from "we inherited it". Mutation-verified, which is
    // the only reason it was caught - it was written specifically to close the gap the A6 work
    // found, and it did not.
    //
    // The real guard is the two data-URI tests above. If upstream ever flips its default AND the
    // explicit assignment has been removed, those fail. If the assignment is present, they keep
    // passing - which is exactly what setting it explicitly buys, and why TextExportOptions does.
    // Until such a flip happens the two are indistinguishable by any test, and pretending
    // otherwise with an assertion that cannot fail is worse than saying so here.
    // =====================================================================================
}
