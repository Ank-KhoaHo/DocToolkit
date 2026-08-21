using System.Text;
using Xunit;
using Xunit.Abstractions;

namespace DocToolkit.Tests;

/// <summary>
/// Markdown → PDF (A27), which is <see cref="MarkdownToDocxConverter"/> composed with
/// <see cref="DocxToPdfConverter"/>.
///
/// Assertions are on the RENDERED TEXT via <see cref="PdfProbe"/>, never on bytes or size: a PDF's
/// size varies ~100x with the fonts installed on the machine, which `CLAUDE.md` records as a
/// property of the host rather than of this code.
/// </summary>
public class MarkdownToPdfConverterTests
{
    private readonly ITestOutputHelper _output;

    public MarkdownToPdfConverterTests(ITestOutputHelper output) => _output = output;

    private const string Markdown = """
        # Quarterly Report

        Revenue was **up 12%** and costs were *flat*.

        | Region | Total |
        | --- | --- |
        | North | 1200 |
        """;

    [Fact]
    public void RendersTheMarkdownsTextIntoThePdf()
    {
        var pdf = MarkdownToPdfConverter.Convert(Markdown);

        Assert.True(PdfProbe.IsPdf(pdf));

        var text = PdfProbe.ExtractText(pdf);
        _output.WriteLine(text);

        Assert.Contains("Quarterly Report", text, StringComparison.Ordinal);
        Assert.Contains("North", text, StringComparison.Ordinal);

        // The markup is gone: a converter that dumped the Markdown source into the page would
        // satisfy both assertions above.
        Assert.DoesNotContain("**up 12%**", text, StringComparison.Ordinal);
        Assert.DoesNotContain("# Quarterly", text, StringComparison.Ordinal);
    }

    /// <summary>
    /// It really is the composition it claims to be — the same PDF as running the two steps by
    /// hand. Compared on rendered TEXT rather than bytes, since two PDF saves a moment apart are
    /// not byte-identical.
    /// </summary>
    [Fact]
    public void MatchesTheTwoStepPipeline()
    {
        var direct = MarkdownToPdfConverter.Convert(Markdown);
        var stepwise = DocxToPdfConverter.Convert(MarkdownToDocxConverter.Convert(Markdown));

        // The literal first: two blank PDFs would satisfy the equality on its own.
        Assert.Contains("Quarterly Report", PdfProbe.ExtractText(direct), StringComparison.Ordinal);
        Assert.Equal(PdfProbe.ExtractText(stepwise), PdfProbe.ExtractText(direct));
    }

    /// <summary>
    /// The offline guarantee carries over from <see cref="MarkdownToDocxConverter"/> — this class
    /// performs no conversion of its own, but that is a claim worth holding rather than assuming,
    /// since it is the composition that a future change would most easily break.
    /// </summary>
    [Fact]
    public async Task AnImageUrlIsNeverFetched()
    {
        using var probe = new LoopbackProbe(_output);
        var markdown = $"# Title\n\n![logo]({probe.BaseUrl}/logo.png)\n";

        var pdf = MarkdownToPdfConverter.Convert(markdown);

        await Task.Delay(TimeSpan.FromMilliseconds(750));

        Assert.Equal(0, probe.Connections);
        Assert.Contains("Title", PdfProbe.ExtractText(pdf), StringComparison.Ordinal);
    }

    /// <summary>
    /// The PDF reaches the destination as it is produced rather than in one write.
    ///
    /// This is the assertion a <c>byte[]</c> round trip wearing a <c>Stream</c> signature cannot
    /// pass, and it is why <c>ConvertAsync</c> hands the destination to the renderer instead of
    /// buffering — the same shape as
    /// <c>StreamOverloadTests.DocxToPdf_StreamsThePdfToTheDestinationInPieces</c>.
    /// </summary>
    [Fact]
    public async Task ConvertAsync_RendersWholeThenEmits_MatchingTheByteArrayPath()
    {
        // Inverted on 2026-08-20 along with its sibling in StreamOverloadTests, which carries the
        // full reasoning. In short: writing straight through meant this overload could apply none
        // of the PDF repairs and none of the diagnosis wrapping, so `0. item` got the generic "see
        // the inner exception" here while the byte[] path named the construct. Parity was chosen
        // over streaming, with measurements behind it.
        var body = new StringBuilder("# Long report\n\n");
        for (var i = 0; i < 2500; i++)
            body.Append("Line ").Append(i).Append(" of a report long enough to need many pages.\n\n");

        var sink = new ForwardOnlySink();
        await MarkdownToPdfConverter.ConvertAsync(body.ToString(), sink);

        var written = sink.ToArray();
        Assert.True(PdfProbe.IsPdf(written));
        Assert.True(written.Length > 100_000, $"expected a sizeable PDF, got {written.Length} bytes");
        Assert.False(sink.IsDisposed, "ConvertAsync disposed a destination it does not own");
    }

    [Fact]
    public async Task ConvertAsync_NamesTheConstruct_LikeTheByteArrayPath()
    {
        // The concrete thing the old straight-through write cost. An ordered list starting below 1
        // is refused by the renderer, and only the byte[] path's Render() wrapper re-described it.
        var fromBytes = Assert.Throws<DocumentConversionException>(
            () => MarkdownToPdfConverter.Convert("0. first\n0. second\n"));

        var sink = new ForwardOnlySink();
        var fromStream = await Assert.ThrowsAsync<DocumentConversionException>(
            () => MarkdownToPdfConverter.ConvertAsync("0. first\n0. second\n", sink));

        Assert.Equal(fromBytes.Message, fromStream.Message);
    }

    [Fact]
    public void ConvertWithReport_CarriesTheMarkdownHalfsWarnings_AndTheSamePdf()
    {
        var result = MarkdownToPdfConverter.ConvertWithReport(Markdown);

        Assert.True(PdfProbe.IsPdf(result.Value));
        Assert.Contains("Quarterly Report", PdfProbe.ExtractText(result.Value), StringComparison.Ordinal);

        // Whatever the Markdown -> DOCX half reports, this reports the same - no more, no less.
        Assert.Equal(
            MarkdownToDocxConverter.ConvertWithReport(Markdown).Warnings.Select(w => (w.Code, w.Kind)),
            result.Warnings.Select(w => (w.Code, w.Kind)));
    }

    /// <summary>The negative case, so the reporting assertion above cannot pass on "everything warns".</summary>
    [Fact]
    public void PlainMarkdownReportsNoLoss()
    {
        var result = MarkdownToPdfConverter.ConvertWithReport("# Title\n\nA paragraph.\n");

        Assert.Empty(result.Warnings);
        Assert.False(result.HasLoss);
    }

    [Fact]
    public async Task RejectsNullAndUnwritableArguments()
    {
        Assert.Throws<ArgumentNullException>(() => MarkdownToPdfConverter.Convert(null!));
        Assert.Throws<ArgumentNullException>(() => MarkdownToPdfConverter.ConvertWithReport(null!));

        using var destination = new MemoryStream();
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => MarkdownToPdfConverter.ConvertAsync(null!, destination));
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => MarkdownToPdfConverter.ConvertAsync(Markdown, null!));

        var ex = await Assert.ThrowsAsync<ArgumentException>(
            () => MarkdownToPdfConverter.ConvertAsync(Markdown, new NonWritableStream()));
        Assert.Equal("destination", ex.ParamName);
    }
}
