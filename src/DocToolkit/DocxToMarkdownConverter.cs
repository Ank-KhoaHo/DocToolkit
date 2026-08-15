using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

namespace DocToolkit;

/// <summary>
/// Converts a Word (.docx) package to Markdown, keeping the structure
/// <see cref="DocxEditor.ExtractText(byte[])"/> throws away: a heading becomes <c>#</c>, a table
/// becomes a pipe table. Use <c>ExtractText</c> when flat text is what you want.
///
/// <b>The result is self-contained.</b> Images in the source document are embedded as
/// <c>data:</c> URIs rather than written beside the output and referenced by path, so the string
/// this returns stands on its own.
/// </summary>
public static class DocxToMarkdownConverter
{
    private const string FailureMessage =
        "Failed to convert DOCX to Markdown. See the inner exception for details.";

    /// <summary>Converts the .docx in <paramref name="docx"/> to Markdown.</summary>
    /// <param name="docx">The document to convert.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The document could not be converted.</exception>
    /// <example>
    /// <code source="../../tests/DocToolkit.Tests/DocumentationExamples.cs" region="DocxToMarkdown"/>
    /// </example>
    public static string Convert(byte[] docx)
    {
        ArgumentNullException.ThrowIfNull(docx);
        if (docx.Length == 0)
            throw new ArgumentException("DOCX content was empty.", nameof(docx));

        try
        {
            // Expandable copy: the loader opens the package read/write, as DocxToPdfConverter does.
            using var input = new MemoryStream();
            input.Write(docx, 0, docx.Length);
            input.Position = 0;

            using var word = WordDocument.Load(input);
            return word.ToMarkdown(TextExportOptions.ForMarkdown());
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException(FailureMessage, ex);
        }
    }

    /// <summary>
    /// Converts the .docx in <paramref name="docx"/> to Markdown, and reports what the conversion
    /// could not carry across.
    /// </summary>
    /// <remarks>
    /// <see cref="Convert(byte[])"/> returns the same Markdown and is unchanged.
    ///
    /// <b>The conversion runs twice here, deliberately.</b> The report and the string come from two
    /// different upstream entry points, and the report-carrying one renders a
    /// <c>MarkdownDoc</c> whose own output is <i>not</i> byte-identical to what
    /// <see cref="Convert(byte[])"/> produces — measured 2026-08-13: different line endings and a
    /// trailing newline. Rendering from the report would therefore have changed this converter's
    /// output for anyone who switched overloads. Running the plain conversion for the value keeps
    /// the two overloads returning exactly the same string, and the extra pass is paid only by a
    /// caller who asked for the report. Calling both on one loaded document is measured safe: the
    /// second call returns what a fresh load returns.
    ///
    /// The result is always usable: this reports loss, it does not refuse.
    /// </remarks>
    /// <param name="docx">The document to convert.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The document could not be converted.</exception>
    public static ConversionResult<string> ConvertWithReport(byte[] docx)
    {
        ArgumentNullException.ThrowIfNull(docx);
        if (docx.Length == 0)
            throw new ArgumentException("DOCX content was empty.", nameof(docx));

        try
        {
            using var input = new MemoryStream();
            input.Write(docx, 0, docx.Length);
            input.Position = 0;

            using var word = WordDocument.Load(input);
            var options = TextExportOptions.ForMarkdown();

            var report = word.ToMarkdownDocumentResult(options);
            var markdown = word.ToMarkdown(options);

            var warnings = new List<ConversionWarning>(report.Report.Diagnostics.Count);
            foreach (var d in report.Report.Diagnostics)
                warnings.Add(ConversionDiagnostics.Warning(d.Code, d.Message, d.LossKind));

            return new ConversionResult<string>(markdown, warnings);
        }
        // No `when (ex is not ArgumentException)` filter here. Arguments are validated BEFORE
        // the try, so the filter could only ever catch an ArgumentException raised inside
        // the conversion itself - which is a conversion failure, not a caller mistake. With
        // the filter, this method let that escape raw while the sibling Convert wrapped it,
        // so a caller catching DocumentConversionException around both crashed on one.
        catch (Exception ex)
        {
            throw new DocumentConversionException(FailureMessage, ex);
        }
    }

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and converts it to Markdown, reporting what the
    /// conversion could not carry across.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and is not disposed, closed or sought.
    /// </summary>
    /// <inheritdoc cref="ConvertWithReport(byte[])"/>
    /// <param name="source">The stream the .docx package is read from.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    public static async Task<ConversionResult<string>> ConvertWithReportAsync(
        Stream source, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, "DOCX content was empty.", nameof(source), FailureMessage, ct)
            .ConfigureAwait(false);

        return ConvertWithReport(docx.ToArray());
    }

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and returns it as Markdown.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and is not disposed, closed or sought, so
    /// it may be forward-only — an HTTP request body, for instance.
    /// </summary>
    /// <param name="source">The stream the .docx package is read from.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be converted.</exception>
    public static async Task<string> ConvertAsync(Stream source, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, "DOCX content was empty.", nameof(source), FailureMessage, ct)
            .ConfigureAwait(false);

        return Convert(docx.ToArray());
    }
}
