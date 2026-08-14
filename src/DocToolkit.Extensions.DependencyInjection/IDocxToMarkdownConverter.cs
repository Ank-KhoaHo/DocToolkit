namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Converts a Word (.docx) package to Markdown, keeping the structure
/// <see cref="IDocxEditor.ExtractText(byte[])"/> throws away. Registered by
/// <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.
///
/// A heading becomes <c>#</c> and a table becomes a pipe table; images come back as data: URIs.
/// </summary>
public interface IDocxToMarkdownConverter
{
    /// <summary>Converts the .docx in <paramref name="docx"/>.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be converted.</exception>
    string Convert(byte[] docx);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and returns the converted text.
    /// <paramref name="source"/> is read to its end and is not disposed, closed or sought, so it
    /// may be forward-only.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be converted.</exception>
    Task<string> ConvertAsync(Stream source, CancellationToken ct = default);

    /// <summary>
    /// Converts the .docx in <paramref name="docx"/> and reports what the conversion could not
    /// carry across.
    /// </summary>
    /// <remarks><see cref="Convert(byte[])"/> returns the same Markdown.</remarks>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be converted.</exception>
    DocToolkit.ConversionResult<string> ConvertWithReport(byte[] docx);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and converts it, reporting what the conversion
    /// could not carry across. <paramref name="source"/> is read to its end and is not disposed,
    /// closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be converted.</exception>
    Task<DocToolkit.ConversionResult<string>> ConvertWithReportAsync(
        Stream source, CancellationToken ct = default);
}
