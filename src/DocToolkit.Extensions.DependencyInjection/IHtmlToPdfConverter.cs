namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Converts HTML straight to PDF by pivoting through DOCX. Registered by
/// <see cref="ServiceCollectionExtensions.AddDocToolkit"/>; remote image download is controlled
/// once, at registration, via <see cref="DocToolkitOptions.AllowRemoteImageDownload"/>, and
/// bounded by <see cref="DocToolkitOptions.RemoteImage"/>.
/// </summary>
public interface IHtmlToPdfConverter
{
    /// <summary>Converts <paramref name="html"/> straight to PDF bytes.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The HTML could not be converted.</exception>
    Task<byte[]> ConvertAsync(string html, CancellationToken ct = default);

    /// <summary>
    /// Converts <paramref name="html"/> and writes the PDF to <paramref name="destination"/>.
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought.
    ///
    /// The PDF is written straight through as the renderer produces it, so a failure part-way
    /// leaves whatever had already been produced on <paramref name="destination"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> or <paramref name="destination"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="destination"/> is not writable.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The HTML could not be converted or written.</exception>
    Task ConvertAsync(string html, Stream destination, CancellationToken ct = default);

    /// <summary>
    /// As above, laid out on <paramref name="page"/> rather than the A4 default.
    /// </summary>
    /// <param name="html">The markup to convert.</param>
    /// <param name="page">The page size, orientation and margins.</param>
    /// <param name="ct">Cancels the conversion.</param>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> or <paramref name="page"/> is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The HTML could not be converted.</exception>
    Task<byte[]> ConvertAsync(string html, DocToolkit.PageSetup page, CancellationToken ct = default);

    /// <summary>
    /// As above, laid out on <paramref name="page"/> rather than the A4 default.
    /// </summary>
    /// <param name="html">The markup to convert.</param>
    /// <param name="page">The page size, orientation and margins.</param>
    /// <param name="destination">The stream the result is written to.</param>
    /// <param name="ct">Cancels the conversion and the write.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="destination"/> is not writable.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The HTML could not be converted.</exception>
    Task ConvertAsync(string html, DocToolkit.PageSetup page, Stream destination, CancellationToken ct = default);
}
