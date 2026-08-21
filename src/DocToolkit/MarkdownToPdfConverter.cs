namespace DocToolkit;

/// <summary>
/// Converts Markdown to PDF by way of DOCX.
/// </summary>
/// <remarks>
/// <b>This is a composition, and must stay one.</b> It calls
/// <see cref="MarkdownToDocxConverter"/> and then <see cref="DocxToPdfConverter"/> — exactly as
/// <see cref="HtmlToPdfConverter"/> pivots through DOCX, and for the same reason: no
/// permissively-licensed, NuGet-only, Linux-safe library renders either format to PDF directly.
/// Do not reimplement conversion inside this class; a second rendering path is how the two come to
/// disagree about what a document looks like.
///
/// Everything <see cref="MarkdownToDocxConverter"/> guarantees carries over unchanged, because
/// this performs no conversion of its own: <b>nothing here reaches the network or the disk</b>, a
/// remote image reference becomes a hyperlink rather than a fetch, and a local file reference is
/// refused.
///
/// The fidelity caveats of DOCX → PDF apply — see <see cref="DocxToPdfConverter"/>.
/// </remarks>
public static class MarkdownToPdfConverter
{
    private const string FailureMessage = "Failed to convert Markdown to PDF.";

    /// <summary>Converts <paramref name="markdown"/> to a PDF.</summary>
    /// <param name="markdown">The Markdown to convert.</param>
    /// <exception cref="ArgumentNullException"><paramref name="markdown"/> is null.</exception>
    /// <exception cref="DocumentConversionException">The Markdown could not be converted.</exception>
    public static byte[] Convert(string markdown)
    {
        ArgumentNullException.ThrowIfNull(markdown);

        return Render(markdown, () => DocxToPdfConverter.Convert(MarkdownToDocxConverter.Convert(markdown)));
    }

    /// <summary>
    /// Runs a Markdown-to-PDF render and re-describes the one failure the PDF stage cannot explain.
    /// </summary>
    /// <remarks>
    /// <b>An ordered list starting below 1 is refused by the PDF renderer, and the message a caller
    /// gets says only that a DOCX could not be rendered.</b> That is accurate and useless: they
    /// wrote Markdown, the DOCX is an implementation detail of this path, and the construct at fault
    /// is <c>0. item</c> - which converts to DOCX perfectly well, so it is genuinely a PDF-only
    /// limit rather than something wrong with their document.
    ///
    /// <see cref="DocxToPdfConverter"/> cannot say any of that: it has never seen the Markdown. So
    /// it is said here, where the source is still in scope, and only for a cause that can be told
    /// apart - everything else propagates exactly as it did.
    /// </remarks>
    private static byte[] Render(string markdown, Func<byte[]> render)
    {
        try
        {
            return render();
        }
        catch (DocumentConversionException ex) when (
            ex.InnerException is not null
            && MarkdownFailureDiagnosis.Describe(ex.InnerException, markdown) is not null)
        {
            throw new DocumentConversionException(
                MarkdownFailureDiagnosis.Describe(ex.InnerException!, markdown)!, ex.InnerException!);
        }
    }

    /// <summary>
    /// Converts <paramref name="markdown"/> and writes the PDF to <paramref name="destination"/>.
    ///
    /// <paramref name="destination"/> is <b>written</b> and is neither disposed, closed nor sought,
    /// so it may be forward-only — an HTTP response body, for instance.
    /// </summary>
    /// <remarks>
    /// The PDF reaches <paramref name="destination"/> <b>as it is produced</b> rather than being
    /// assembled in full first, which is the one place this class is more than two calls in a row:
    /// it hands the destination to the renderer the way
    /// <see cref="HtmlToPdfConverter"/> does. Buffering instead would be simpler and would give up
    /// exactly the property the <c>Stream</c> overloads exist for.
    /// </remarks>
    /// <param name="markdown">The Markdown to convert.</param>
    /// <param name="destination">The stream the PDF is written to.</param>
    /// <param name="ct">Cancels the conversion and the write.</param>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="destination"/> is not writable.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The Markdown could not be converted or written.</exception>
    public static async Task ConvertAsync(
        string markdown, Stream destination, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(markdown);
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        // THROUGH Convert - this converter's own byte[] path, not a re-composition of the stages.
        //
        // This used to hand the package to OfficeIMO's writer, which skipped ListMarkerSubstitution
        // and the negative-indent clamp. Same divergence measured across the DOCX and HTML paths on
        // 2026-08-20; see HtmlToPdfConverter.EmitAsync for the numbers.
        //
        // AND CALLING Convert(markdown) IS THE POINT, rather than composing
        // MarkdownToDocxConverter + DocxToPdfConverter here. The first attempt did compose them and
        // was still wrong: it picked up the PDF repairs and missed Render(), the wrapper that
        // re-describes an ordered list starting below 1. `0. item` got "Failed to render DOCX to
        // PDF. See the inner exception" while the byte[] path named the construct. The parity test
        // caught it. Re-composing a pipeline is how the two paths drift; calling the sibling is how
        // they cannot.
        //
        // It also retires the expandable-vs-read-only question this comment used to argue about:
        // Convert owns the copy OfficeIMO's read/write open needs, in the one place that decides it.
        var pdf = Convert(markdown);

        ct.ThrowIfCancellationRequested();

        using var scratch = new MemoryStream(pdf, writable: false);
        await StreamPipeline.EmitAsync(scratch, destination, FailureMessage, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Converts <paramref name="markdown"/> to a PDF, and reports what the conversion could not
    /// carry across.
    /// </summary>
    /// <remarks>
    /// <b>The warnings come from the Markdown → DOCX half only.</b> The DOCX → PDF half renders
    /// rather than converts and produces no report, so a caveat that applies to it — a chart or a
    /// shape effect the renderer cannot draw — will not appear here. Saying so is the point: the
    /// README's known-limitations table already records that the PDF renderers have no warning
    /// channel, and this method must not be read as having closed that gap.
    /// </remarks>
    /// <param name="markdown">The Markdown to convert.</param>
    /// <exception cref="ArgumentNullException"><paramref name="markdown"/> is null.</exception>
    /// <exception cref="DocumentConversionException">The Markdown could not be converted.</exception>
    public static ConversionResult<byte[]> ConvertWithReport(string markdown)
    {
        ArgumentNullException.ThrowIfNull(markdown);

        var docx = MarkdownToDocxConverter.ConvertWithReport(markdown);
        return new ConversionResult<byte[]>(
            Render(markdown, () => DocxToPdfConverter.Convert(docx.Value)), docx.Warnings);
    }
}
