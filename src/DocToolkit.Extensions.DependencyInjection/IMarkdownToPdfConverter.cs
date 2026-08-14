namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Converts Markdown to PDF by way of DOCX. Registered by
/// <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.
/// </summary>
/// <remarks>
/// A composition of <see cref="IMarkdownToDocxConverter"/> and <see cref="IDocxToPdfConverter"/>,
/// exactly as <see cref="IHtmlToPdfConverter"/> pivots through DOCX. Everything the Markdown
/// importer guarantees carries over unchanged, because this performs no conversion of its own:
/// <b>nothing here reaches the network or the disk.</b>
///
/// The fidelity caveats of DOCX → PDF apply.
/// </remarks>
public interface IMarkdownToPdfConverter
{
    /// <summary>Converts <paramref name="markdown"/> to a PDF.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="markdown"/> is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be converted.</exception>
    byte[] Convert(string markdown);

    /// <summary>
    /// Converts <paramref name="markdown"/> and writes the PDF to <paramref name="destination"/>.
    /// The PDF reaches the destination as it is produced rather than being assembled in full
    /// first; the stream is written but neither disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="destination"/> is not writable.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be converted or written.</exception>
    Task ConvertAsync(string markdown, Stream destination, CancellationToken ct = default);

    /// <summary>
    /// Converts <paramref name="markdown"/> and reports what the conversion could not carry
    /// across.
    /// </summary>
    /// <remarks>
    /// <b>The warnings come from the Markdown → DOCX half only.</b> The DOCX → PDF half renders
    /// rather than converts and produces no report, so a caveat that applies to it will not appear
    /// here.
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="markdown"/> is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be converted.</exception>
    DocToolkit.ConversionResult<byte[]> ConvertWithReport(string markdown);
}
