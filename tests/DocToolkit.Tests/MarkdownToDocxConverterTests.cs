using Xunit;
using Xunit.Abstractions;

namespace DocToolkit.Tests;

/// <summary>
/// Markdown → DOCX (A19).
///
/// Two things this file is careful about, both learned the same day it was written:
/// every assertion is anchored to a LITERAL rather than to a count or to the converter compared
/// against itself (B16), and the air-gap assertions carry a POSITIVE CONTROL, because a
/// zero-connection result passes identically when the probe is broken.
/// </summary>
public class MarkdownToDocxConverterTests
{
    private readonly ITestOutputHelper _output;

    public MarkdownToDocxConverterTests(ITestOutputHelper output) => _output = output;

    private const string Markdown = """
        # Quarterly Report

        Revenue was **up 12%** and costs were *flat*.

        - North grew
        - South held

        | Region | Total |
        | --- | --- |
        | North | 1200 |
        """;

    // =====================================================================================
    // Content
    // =====================================================================================

    [Fact]
    public void ConvertsHeadingsEmphasisListsAndTables_AsLiteralText()
    {
        var docx = MarkdownToDocxConverter.Convert(Markdown);

        var text = DocxEditor.ExtractText(docx);
        _output.WriteLine(text);

        Assert.Contains("Quarterly Report", text, StringComparison.Ordinal);
        Assert.Contains("up 12%", text, StringComparison.Ordinal);
        Assert.Contains("North grew", text, StringComparison.Ordinal);
        Assert.Contains("South held", text, StringComparison.Ordinal);

        // The markup itself must be gone - a converter that emitted the Markdown source verbatim
        // into a paragraph would satisfy every Contains above.
        Assert.DoesNotContain("**up 12%**", text, StringComparison.Ordinal);
        Assert.DoesNotContain("# Quarterly", text, StringComparison.Ordinal);
    }

    /// <summary>
    /// A pipe table becomes a real <c>w:tbl</c>, not a paragraph of pipes. Read back through
    /// <c>ReadTable</c>, which only sees genuine tables.
    /// </summary>
    [Fact]
    public void APipeTableBecomesARealTable()
    {
        var docx = MarkdownToDocxConverter.Convert(Markdown);

        Assert.Equal(1, DocxEditor.TableCount(docx));

        var table = DocxEditor.ReadTable(docx, 0);
        Assert.Equal(new[] { "Region", "Total" }, table[0]);
        Assert.Equal(new[] { "North", "1200" }, table[^1]);
    }

    /// <summary>
    /// The round trip A19 exists to complete. Asserted on literals, not on
    /// <c>Convert(Convert(x)) == x</c>, which would hold if both directions were equally broken.
    /// </summary>
    [Fact]
    public void RoundTripsThroughDocxToMarkdown()
    {
        var backAgain = DocxToMarkdownConverter.Convert(MarkdownToDocxConverter.Convert(Markdown));
        _output.WriteLine(backAgain);

        Assert.Contains("# Quarterly Report", backAgain, StringComparison.Ordinal);
        Assert.Contains("up 12%", backAgain, StringComparison.Ordinal);
        Assert.Contains("| Region |", backAgain, StringComparison.Ordinal);
        Assert.Contains("North", backAgain, StringComparison.Ordinal);
    }

    // =====================================================================================
    // The offline guarantee
    // =====================================================================================

    /// <summary>
    /// Markdown naming a loopback URL as an image converts without connecting to it.
    ///
    /// Meaningless on its own — see the positive control below, without which this passes just as
    /// happily when the probe cannot observe anything at all.
    /// </summary>
    [Fact]
    public async Task AnImageUrlIsNeverFetched()
    {
        using var probe = new LoopbackProbe(_output);
        var markdown = $"# Title\n\n![logo]({probe.BaseUrl}/logo.png)\n";

        var docx = MarkdownToDocxConverter.Convert(markdown);

        await Task.Delay(TimeSpan.FromMilliseconds(750));

        Assert.Equal(0, probe.Connections);
        Assert.Contains("Title", DocxEditor.ExtractText(docx), StringComparison.Ordinal);
    }

    /// <summary>
    /// <b>The positive control, and it is load-bearing.</b> It proves this probe can see a fetch on
    /// this code path, so the zero above is a statement about the converter rather than about the
    /// probe. `CLAUDE.md` records this hazard biting twice; both times the wrong fix was deleting
    /// the positive test.
    ///
    /// It deliberately reaches around <see cref="MarkdownToDocxConverter"/> and drives the
    /// underlying converter with a resolver supplied, because there is no public way to enable
    /// fetching — which is the point of the design.
    /// </summary>
    [Fact]
    public async Task PositiveControl_TheProbeSeesAFetchWhenOneIsGenuinelyMade()
    {
        using var probe = new LoopbackProbe(_output);
        var markdown = $"# Title\n\n![logo]({probe.BaseUrl}/logo.png)\n";

        using var http = new HttpClient();
        var options = new OfficeIMO.Word.Markdown.MarkdownToWordOptions
        {
            RemoteImageResolver = uri =>
            {
                // The probe is a raw TcpListener and does not speak valid HTTP, so the response
                // legitimately fails to parse. Narrowed to HttpRequestException on purpose: the
                // CONNECTION is what this control measures, and anything else going wrong here
                // should fail the test loudly rather than be swallowed.
                try
                {
                    return http.GetByteArrayAsync(uri).GetAwaiter().GetResult();
                }
                catch (HttpRequestException)
                {
                    return null;
                }
            },
        };

        var parsed = OfficeIMO.Markdown.MarkdownReader.Parse(markdown);
        using var word = OfficeIMO.Word.Markdown.WordMarkdownConverterExtensions
            .ToWordDocument(parsed, options);

        Assert.True(
            await probe.WaitForConnectionAsync(TimeSpan.FromSeconds(5)),
            "the probe saw no connection even with a fetching resolver supplied - it can no longer "
            + "discriminate, which makes AnImageUrlIsNeverFetched vacuous");
    }

    /// <summary>
    /// A local file reference embeds nothing. Reading a path out of untrusted document content is
    /// a file-disclosure primitive: the document would be choosing which file lands in the output.
    /// </summary>
    [Fact]
    public void ALocalFileImageIsNotRead()
    {
        using var secret = new TempFile();
        File.WriteAllBytes(secret.Path, ImageFixtures.Png());

        var markdown = $"# Title\n\n![local](file:///{secret.Path.Replace('\\', '/')})\n";

        var docx = MarkdownToDocxConverter.Convert(markdown);

        // No image part means nothing off the disk reached the package.
        Assert.Empty(DocxFixtures.Read(docx, main => main.ImageParts.ToList()));
        Assert.Contains("Title", DocxEditor.ExtractText(docx), StringComparison.Ordinal);
    }

    /// <summary>
    /// <b>A RELATIVE path, which is the form the guard actually governs.</b>
    ///
    /// <see cref="ALocalFileImageIsNotRead"/> above uses a <c>file://</c> URI, and mutation (B14)
    /// showed it passes just as happily with <c>AllowLocalImages</c> flipped to <c>true</c> — that
    /// scheme is refused by a different mechanism entirely, so it proves nothing about this
    /// setting. A bare relative reference is the one that separates them: measured, it is read and
    /// embedded when the flag is flipped and blocked when it is not.
    ///
    /// The file is written to the working directory because that is what a relative reference
    /// resolves against, which is precisely the disclosure being prevented: a service converting
    /// Markdown somebody else wrote would otherwise let that document choose which file lands in
    /// the output.
    /// </summary>
    [Fact]
    public void ARelativeFileImageIsNotRead()
    {
        var name = $"doctoolkit-probe-{Guid.NewGuid():N}.png";
        var path = Path.Join(Directory.GetCurrentDirectory(), name);
        File.WriteAllBytes(path, ImageFixtures.Png());

        try
        {
            var docx = MarkdownToDocxConverter.Convert($"# Title\n\n![local]({name})\n");

            Assert.Empty(DocxFixtures.Read(docx, main => main.ImageParts.ToList()));
            Assert.Contains("Title", DocxEditor.ExtractText(docx), StringComparison.Ordinal);
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// A remote image is not merely dropped — it becomes a <b>hyperlink carrying its alt text</b>,
    /// so the reader can still follow it deliberately. That is why the warning is an
    /// <see cref="ConversionLossKind.Approximation"/> rather than an
    /// <see cref="ConversionLossKind.Omission"/>, and asserting the text back is what makes the
    /// label a measurement rather than a claim.
    /// </summary>
    [Fact]
    public void ARemoteImageBecomesAHyperlinkRatherThanVanishing()
    {
        var result = MarkdownToDocxConverter.ConvertWithReport(
            "# Release notes\n\n![build status](https://img.example.com/badge.svg)\n");

        var text = DocxEditor.ExtractText(result.Value);
        _output.WriteLine(text);

        Assert.Contains("build status", text, StringComparison.Ordinal);
        Assert.Empty(DocxFixtures.Read(result.Value, main => main.ImageParts.ToList()));

        var warning = Assert.Single(result.Warnings);
        Assert.Equal(ConversionLossKind.Approximation, warning.Kind);
    }

    /// <summary>
    /// A <c>data:</c> URI carries its own bytes, so honouring it costs no I/O and does not weaken
    /// the offline guarantee. It is bounded anyway, at 32 MB, so a document cannot make the
    /// process materialise an arbitrary amount of memory.
    ///
    /// One assertion, three mutants: this fails if <c>AllowDataUriImages</c> is flipped off, and
    /// also if the byte cap is mis-computed — <c>32 * 1024 * 1024</c> mutated to
    /// <c>32 * 1024 / 1024</c> or <c>32 / 1024</c> puts the limit below this image's size, and all
    /// three survived the suite before this test existed.
    /// </summary>
    [Fact]
    public void ADataUriImageIsEmbedded_BecauseItCostsNoIO()
    {
        var png = ImageFixtures.Png();
        var markdown = $"# Title\n\n![inline](data:image/png;base64,{Convert.ToBase64String(png)})\n";

        var docx = MarkdownToDocxConverter.Convert(markdown);

        var parts = DocxFixtures.Read(docx, main => main.ImageParts.Select(p =>
        {
            using var stream = p.GetStream();
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
        }).ToList());

        var embedded = Assert.Single(parts);
        Assert.Equal(png, embedded);
    }

    // =====================================================================================
    // Reporting, guards and overloads
    // =====================================================================================

    [Fact]
    public void ConvertWithReport_ReturnsTheSameDocumentAsConvert()
    {
        var plain = MarkdownToDocxConverter.Convert(Markdown);
        var reported = MarkdownToDocxConverter.ConvertWithReport(Markdown);

        // Same readable content, and real content rather than two empty documents.
        Assert.Equal(DocxEditor.ExtractText(plain), DocxEditor.ExtractText(reported.Value));
        Assert.Contains("Quarterly Report", DocxEditor.ExtractText(reported.Value), StringComparison.Ordinal);
    }

    /// <summary>
    /// The negative case. Without it, a mapper that reported every element as a warning would
    /// satisfy any positive assertion about reporting.
    /// </summary>
    [Fact]
    public void PlainMarkdownReportsNoLoss()
    {
        var result = MarkdownToDocxConverter.ConvertWithReport("# Title\n\nA paragraph.\n");

        Assert.Empty(result.Warnings);
        Assert.False(result.HasLoss);
    }

    [Fact]
    public async Task ConvertAsync_WritesAWholeDocumentToAForwardOnlySink_AndLeavesItOpen()
    {
        var sink = new ForwardOnlySink();

        await MarkdownToDocxConverter.ConvertAsync(Markdown, sink);

        var written = sink.ToArray();
        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, written.Take(4).ToArray());
        Assert.Contains("Quarterly Report", DocxEditor.ExtractText(written), StringComparison.Ordinal);
        Assert.False(sink.IsDisposed, "ConvertAsync disposed a destination it does not own");
    }

    [Fact]
    public async Task RejectsNullAndUnwritableArguments()
    {
        Assert.Throws<ArgumentNullException>(() => MarkdownToDocxConverter.Convert(null!));
        Assert.Throws<ArgumentNullException>(() => MarkdownToDocxConverter.ConvertWithReport(null!));

        using var destination = new MemoryStream();
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => MarkdownToDocxConverter.ConvertAsync(null!, destination));

        await Assert.ThrowsAsync<ArgumentNullException>(
            () => MarkdownToDocxConverter.ConvertAsync(Markdown, null!));

        var ex = await Assert.ThrowsAsync<ArgumentException>(
            () => MarkdownToDocxConverter.ConvertAsync(Markdown, new NonWritableStream()));
        Assert.Equal("destination", ex.ParamName);
    }

    /// <summary>
    /// Empty Markdown is a valid document, not an error — unlike an empty .docx, which is a
    /// truncated upload. Recorded as a decision rather than left to be inferred from its absence.
    /// </summary>
    [Fact]
    public void EmptyMarkdownProducesAnEmptyButValidDocument()
    {
        var docx = MarkdownToDocxConverter.Convert(string.Empty);

        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, docx.Take(4).ToArray());
        Assert.Equal(string.Empty, DocxEditor.ExtractText(docx).Trim());
    }
}
