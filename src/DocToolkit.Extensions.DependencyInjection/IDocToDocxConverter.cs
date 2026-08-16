namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Reads a Word 97-2003 binary document (.doc) and converts it to a .docx package, or reads its
/// text directly. Registered by <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.
/// </summary>
/// <remarks>
/// <b>Import only</b> — there is no .doc writing, because the underlying library reports native
/// .doc saving as unsupported.
///
/// <b><see cref="Convert(byte[])"/> refuses by default</b> when the source holds content a .docx
/// cannot carry, which in practice is any .doc containing a table. See
/// <see cref="DocToolkit.LegacyDocOptions"/>. <see cref="ExtractText(byte[])"/> takes no options and
/// never refuses.
/// </remarks>
public interface IDocToDocxConverter
{
    /// <summary>Converts the legacy .doc in <paramref name="doc"/> to a .docx package.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="doc"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="doc"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// It could not be converted, or it holds content a .docx cannot carry and
    /// <see cref="DocToolkit.LegacyDocOptions.AllowContentLoss"/> was not set.
    /// </exception>
    byte[] Convert(byte[] doc);

    /// <inheritdoc cref="Convert(byte[])"/>
    /// <param name="doc">The Word 97-2003 binary document to convert.</param>
    /// <param name="options">How to treat content the .docx cannot carry. Null means refuse.</param>
    byte[] Convert(byte[] doc, DocToolkit.LegacyDocOptions? options);

    /// <summary>
    /// Converts the legacy .doc in <paramref name="doc"/> and reports what the import could not
    /// carry across.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="doc"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="doc"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be converted.</exception>
    DocToolkit.ConversionResult<byte[]> ConvertWithReport(
        byte[] doc, DocToolkit.LegacyDocOptions? options = null);

    /// <summary>Reads the text of the legacy .doc in <paramref name="doc"/>, table cells included.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="doc"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="doc"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be read.</exception>
    string ExtractText(byte[] doc);

    /// <summary>
    /// Reads a .doc from <paramref name="source"/> and writes the converted .docx to
    /// <paramref name="destination"/>. Neither stream is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either stream is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or
    /// <paramref name="destination"/> is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be converted.</exception>
    Task ConvertAsync(Stream source, Stream destination, CancellationToken ct = default);

    /// <inheritdoc cref="ConvertAsync(Stream, Stream, CancellationToken)"/>
    /// <param name="source">The stream the .doc is read from.</param>
    /// <param name="destination">The stream the .docx package is written to.</param>
    /// <param name="options">How to treat content the .docx cannot carry. Null means refuse.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    Task ConvertAsync(
        Stream source, Stream destination, DocToolkit.LegacyDocOptions? options,
        CancellationToken ct = default);

    /// <summary>
    /// Reads a .doc from <paramref name="source"/> and returns its text. <paramref name="source"/>
    /// is read to its end and is not disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be read.</exception>
    Task<string> ExtractTextAsync(Stream source, CancellationToken ct = default);
}
