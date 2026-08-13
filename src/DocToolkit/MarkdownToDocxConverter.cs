using OfficeIMO.Markdown;
using OfficeIMO.Word.Markdown;

namespace DocToolkit;

/// <summary>
/// Converts Markdown to a Word (.docx) package, completing the round trip
/// <see cref="DocxToMarkdownConverter"/> opened: a <c>#</c> becomes a heading, a pipe table becomes
/// a table.
/// </summary>
/// <remarks>
/// <b>Nothing here reaches the network or the disk.</b> A remote image reference becomes a
/// hyperlink rather than a fetch, and a local file reference is refused — reading a path out of
/// untrusted document content would let the document choose which file gets embedded in the output.
/// <c>data:</c> images are inlined, since they carry their own bytes. See
/// <see cref="MarkdownImportOptions"/>, which sets each of those explicitly rather than inheriting
/// it.
///
/// <b>No <c>Stream source</c> overload.</b> The input is a <see cref="string"/> rather than a
/// document, so a caller holding bytes decides their own encoding — the same reason
/// <see cref="HtmlToDocxConverter"/> has none.
/// </remarks>
public static class MarkdownToDocxConverter
{
    private const string FailureMessage = "Failed to convert Markdown to DOCX.";

    /// <summary>Converts <paramref name="markdown"/> to a .docx package.</summary>
    /// <param name="markdown">The Markdown to convert.</param>
    /// <exception cref="ArgumentNullException"><paramref name="markdown"/> is null.</exception>
    /// <exception cref="DocumentConversionException">The Markdown could not be converted.</exception>
    public static byte[] Convert(string markdown)
    {
        ArgumentNullException.ThrowIfNull(markdown);

        return ConvertCore(markdown, out _);
    }

    /// <summary>
    /// Converts <paramref name="markdown"/> to a .docx package, and reports what the conversion
    /// could not carry across.
    /// </summary>
    /// <remarks>
    /// <see cref="Convert(string)"/> produces the same document and is unchanged; this overload is
    /// opt-in. Markdown and DOCX do not express the same set of things, so a conversion between
    /// them can legitimately drop or approximate content — and the underlying converter reports
    /// that rather than leaving it to be discovered.
    ///
    /// The result is always usable: this reports loss, it does not refuse.
    /// </remarks>
    /// <param name="markdown">The Markdown to convert.</param>
    /// <exception cref="ArgumentNullException"><paramref name="markdown"/> is null.</exception>
    /// <exception cref="DocumentConversionException">The Markdown could not be converted.</exception>
    public static ConversionResult<byte[]> ConvertWithReport(string markdown)
    {
        ArgumentNullException.ThrowIfNull(markdown);

        var docx = ConvertCore(markdown, out var warnings);
        return new ConversionResult<byte[]>(docx, warnings);
    }

    /// <summary>
    /// Converts <paramref name="markdown"/> and writes the .docx package to
    /// <paramref name="destination"/>.
    ///
    /// <paramref name="destination"/> is <b>written</b> and is neither disposed, closed nor sought,
    /// so it may be forward-only — an HTTP response body, for instance.
    /// </summary>
    /// <param name="markdown">The Markdown to convert.</param>
    /// <param name="destination">The stream the .docx package is written to.</param>
    /// <param name="ct">Cancels the write.</param>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="destination"/> is not writable.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The Markdown could not be converted.</exception>
    public static async Task ConvertAsync(
        string markdown, Stream destination, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(markdown);
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        var docx = ConvertCore(markdown, out _);

        using var buffer = new MemoryStream(docx, writable: false);
        await StreamPipeline.EmitAsync(buffer, destination, FailureMessage, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// The one real implementation; every overload calls it so they cannot drift apart.
    /// </summary>
    /// <remarks>
    /// Uses <c>ToWordDocumentResult</c> rather than <c>ToWordDocument</c> on every path, including
    /// the ones that discard the report. The two are the same conversion, so taking the result
    /// form always keeps <see cref="Convert(string)"/> and <see cref="ConvertWithReport(string)"/>
    /// producing byte-identical documents by construction rather than by a test that has to notice
    /// them diverging.
    /// </remarks>
    private static byte[] ConvertCore(string markdown, out IReadOnlyList<ConversionWarning> warnings)
    {
        try
        {
            var parsed = MarkdownReader.Parse(markdown);
            var result = parsed.ToWordDocumentResult(MarkdownImportOptions.ForImport());

            var collected = new List<ConversionWarning>(result.Report.Diagnostics.Count);
            foreach (var d in result.Report.Diagnostics)
                collected.Add(ConversionDiagnostics.Warning(d.Code, d.Message, d.LossKind));
            warnings = collected;

            using var word = result.Value;
            using var buffer = new MemoryStream();
            word.Save(buffer);
            return buffer.ToArray();
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException(FailureMessage, ex);
        }
    }
}
