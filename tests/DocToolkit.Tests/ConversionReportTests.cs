using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// Conversion warnings (A22).
///
/// The trap this file has to avoid is the one B16 spent a day on the same week: asserting
/// <c>Warnings.Count >= 0</c>, or comparing a report against itself, proves nothing. So every
/// positive case asserts a <b>literal</b> loss kind, and the negative case below is what stops a
/// mapper that reported everything as a warning from satisfying all of them.
/// </summary>
public class ConversionReportTests
{
    private const string Html = "<h1>Quarterly Report</h1><p>Revenue was <strong>up 12%</strong>.</p>";

    private static Task<byte[]> DocxAsync() => HtmlToDocxConverter.ConvertAsync(Html);

    // =====================================================================================
    // The value must not drift from the plain overload
    // =====================================================================================

    /// <summary>
    /// The report-carrying overload returns exactly what the plain one does.
    ///
    /// This is the assertion that stops the two from becoming two different products. It matters
    /// more than it looks for HTML, where <c>ConvertWithReport</c> goes through a DIFFERENT
    /// upstream entry point (<c>ToHtmlResult</c> rather than <c>ToHtml</c>) — measured
    /// byte-identical on 2026-08-13, and this is what reports it if that ever stops being true.
    /// </summary>
    [Fact]
    public async Task Html_ConvertWithReport_ReturnsTheSameHtmlAsConvert()
    {
        var docx = await DocxAsync();

        Assert.Equal(DocxToHtmlConverter.Convert(docx), DocxToHtmlConverter.ConvertWithReport(docx).Value);

        // ...and the value is real HTML, not an empty string that would satisfy the line above
        // however broken both sides were.
        var value = DocxToHtmlConverter.ConvertWithReport(docx).Value;
        Assert.Contains("Quarterly Report", value, StringComparison.Ordinal);
        Assert.Contains("<html", value, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Same for Markdown, where it guards something specific: <c>ConvertWithReport</c> deliberately
    /// runs the plain conversion for its value rather than rendering the report's own
    /// <c>MarkdownDoc</c>, because that render differs (line endings, trailing newline). If someone
    /// "simplifies" it into a single call, this fails.
    /// </summary>
    [Fact]
    public async Task Markdown_ConvertWithReport_ReturnsTheSameMarkdownAsConvert()
    {
        var docx = await DocxAsync();

        Assert.Equal(
            DocxToMarkdownConverter.Convert(docx),
            DocxToMarkdownConverter.ConvertWithReport(docx).Value);

        var value = DocxToMarkdownConverter.ConvertWithReport(docx).Value;
        Assert.Contains("# Quarterly Report", value, StringComparison.Ordinal);
    }

    // =====================================================================================
    // The report itself
    // =====================================================================================

    /// <summary>
    /// A DOCX → HTML conversion reports a real, named loss today.
    ///
    /// <c>SectionLayoutFlattened</c> is raised because this library exports page geometry without
    /// section metadata. It was computed on every call before A22 and thrown away — which is the
    /// whole justification for this feature, so it is asserted by its literal code rather than by
    /// a count.
    /// </summary>
    [Fact]
    public async Task Html_ReportsTheSectionLayoutApproximationItHasAlwaysComputed()
    {
        var result = DocxToHtmlConverter.ConvertWithReport(await DocxAsync());

        Assert.True(result.HasLoss, "expected the HTML conversion to report a loss");

        var warning = Assert.Single(result.Warnings, w => w.Code == "SectionLayoutFlattened");
        Assert.Equal(ConversionLossKind.Approximation, warning.Kind);
        Assert.NotEmpty(warning.Message);
    }

    /// <summary>
    /// <b>The negative case, and it is load-bearing.</b> Without it, a mapper that reported every
    /// element in the document as a warning would satisfy every other test in this file.
    ///
    /// Markdown is the honest place to assert it: measured 2026-08-13, this document converts to
    /// Markdown with no diagnostics at all, while the same document converting to HTML raises one.
    /// So the two together prove the report tracks the conversion rather than being constant.
    /// </summary>
    [Fact]
    public async Task Markdown_ReportsNoLossForADocumentThatLosesNothing()
    {
        var result = DocxToMarkdownConverter.ConvertWithReport(await DocxAsync());

        Assert.Empty(result.Warnings);
        Assert.False(result.HasLoss);
    }

    /// <summary>
    /// <see cref="ConversionResult{T}.HasLoss"/> is derived from the warnings the caller can see,
    /// so the two can never disagree. An informational entry does not count as loss.
    /// </summary>
    [Fact]
    public void HasLoss_IgnoresInformationalWarnings_AndIsTrueForAnyRealLoss()
    {
        Assert.False(new ConversionResult<string>("x", Array.Empty<ConversionWarning>()).HasLoss);

        Assert.False(new ConversionResult<string>("x", new[]
        {
            new ConversionWarning("INFO", "nothing lost", ConversionLossKind.None),
        }).HasLoss);

        foreach (var kind in new[]
                 {
                     ConversionLossKind.Approximation,
                     ConversionLossKind.Omission,
                     ConversionLossKind.Failure,
                 })
        {
            Assert.True(
                new ConversionResult<string>("x", new[]
                {
                    new ConversionWarning("INFO", "nothing lost", ConversionLossKind.None),
                    new ConversionWarning("C", "m", kind),
                }).HasLoss,
                $"{kind} should count as loss");
        }
    }

    // =====================================================================================
    // Stream overloads and guards
    // =====================================================================================

    [Fact]
    public async Task StreamOverloads_ReportTheSameWarnings_AndLeaveTheSourceOpen()
    {
        var docx = await DocxAsync();
        var expected = DocxToHtmlConverter.ConvertWithReport(docx);

        using var source = new MemoryStream(docx, writable: false);
        var streamed = await DocxToHtmlConverter.ConvertWithReportAsync(source);

        Assert.Equal(expected.Value, streamed.Value);
        Assert.Equal(
            expected.Warnings.Select(w => (w.Code, w.Kind)),
            streamed.Warnings.Select(w => (w.Code, w.Kind)));
        Assert.Contains(streamed.Warnings, w => w.Code == "SectionLayoutFlattened");

        Assert.True(source.CanRead, "the source stream was disposed by a method that does not own it");
    }

    [Fact]
    public async Task MarkdownStreamOverload_MatchesTheByteArrayForm()
    {
        var docx = await DocxAsync();

        using var source = new MemoryStream(docx, writable: false);
        var streamed = await DocxToMarkdownConverter.ConvertWithReportAsync(source);

        Assert.Equal(DocxToMarkdownConverter.Convert(docx), streamed.Value);
        Assert.Empty(streamed.Warnings);
    }

    [Fact]
    public async Task RejectsTheSameInputsThePlainOverloadsDo()
    {
        Assert.Throws<ArgumentNullException>(() => DocxToHtmlConverter.ConvertWithReport(null!));
        Assert.Throws<ArgumentNullException>(() => DocxToMarkdownConverter.ConvertWithReport(null!));

        Assert.Throws<ArgumentException>(() => DocxToHtmlConverter.ConvertWithReport(Array.Empty<byte>()));
        Assert.Throws<ArgumentException>(() => DocxToMarkdownConverter.ConvertWithReport(Array.Empty<byte>()));

        await Assert.ThrowsAsync<ArgumentNullException>(
            () => DocxToHtmlConverter.ConvertWithReportAsync(null!));

        using var empty = new MemoryStream();
        var ex = await Assert.ThrowsAsync<ArgumentException>(
            () => DocxToHtmlConverter.ConvertWithReportAsync(empty));
        Assert.Equal("source", ex.ParamName);
    }

    [Fact]
    public void ReportsRubbishInputAsAConversionFailure_RatherThanAWarning()
    {
        // The line between the two: content that cannot be converted at all THROWS. Warnings are
        // for a conversion that succeeded and lost something on the way.
        var rubbish = System.Text.Encoding.ASCII.GetBytes("this is not a docx");

        Assert.Throws<DocumentConversionException>(() => DocxToHtmlConverter.ConvertWithReport(rubbish));
        Assert.Throws<DocumentConversionException>(() => DocxToMarkdownConverter.ConvertWithReport(rubbish));
    }
}
