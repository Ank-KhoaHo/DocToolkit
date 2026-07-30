namespace DocToolkit;

/// <summary>
/// Converts HTML to PDF by pivoting through DOCX.
///
/// There is no permissive, NuGet-only, Linux-safe library that renders HTML to PDF directly:
/// the only free renderers are browsers, and a browser is a native binary. Pivoting through
/// DOCX keeps the whole chain pure managed. See learning-docs/dotnet-doc-libs/report.html.
/// </summary>
public static class HtmlToPdfConverter
{
    /// <summary>
    /// Converts <paramref name="html"/> straight to PDF bytes.
    ///
    /// <b>No network access, and safe in an air-gapped environment.</b> Nothing the markup
    /// references is fetched, and no font is resolved over the network - see
    /// <see cref="ConvertAsync(string, bool, CancellationToken)"/> to opt in.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted.</exception>
    public static Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)
        => ConvertAsync(html, allowRemoteImageDownload: false, ct);

    /// <summary>
    /// Converts <paramref name="html"/> straight to PDF bytes, optionally downloading and
    /// embedding images referenced by absolute <c>http</c>/<c>https</c> URLs.
    ///
    /// <b>Passing <c>true</c> for <paramref name="allowRemoteImageDownload"/> will fail in an
    /// air-gapped or otherwise offline environment</b>, because the HTML stage then issues
    /// outbound HTTP requests to whatever hosts the markup names and a host that does not answer
    /// fails the whole conversion. See
    /// <see cref="HtmlToDocxConverter.ConvertAsync(string, bool, CancellationToken)"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted.</exception>
    public static async Task<byte[]> ConvertAsync(
        string html, bool allowRemoteImageDownload, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(html);
        var docx = await HtmlToDocxConverter.ConvertAsync(html, allowRemoteImageDownload, ct);
        ct.ThrowIfCancellationRequested();
        return DocxToPdfConverter.Convert(docx);
    }

    /// <summary>
    /// Converts <paramref name="html"/> and writes the PDF to <paramref name="destination"/>.
    ///
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought - it belongs to the caller, and may be write-only and
    /// forward-only, such as an HTTP response body.
    ///
    /// <b>No network access, and safe in an air-gapped environment.</b> Nothing the markup
    /// references is fetched, and no font is resolved over the network - see
    /// <see cref="ConvertAsync(string, bool, Stream, CancellationToken)"/> to opt in.
    /// </summary>
    /// <param name="html">The markup to convert.</param>
    /// <param name="destination">The stream the PDF is written to.</param>
    /// <param name="ct">Cancels the conversion and the write to <paramref name="destination"/>.</param>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="destination"/> is not writable.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted or written.</exception>
    public static Task ConvertAsync(string html, Stream destination, CancellationToken ct = default)
        => ConvertAsync(html, allowRemoteImageDownload: false, destination, ct);

    /// <summary>
    /// Converts <paramref name="html"/> and writes the PDF to <paramref name="destination"/>,
    /// optionally downloading and embedding images referenced by absolute <c>http</c>/<c>https</c>
    /// URLs.
    ///
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought - it belongs to the caller, and may be write-only and
    /// forward-only, such as an HTTP response body. The PDF is written straight through as the
    /// renderer produces it, so a failure part-way leaves whatever had already been produced on
    /// <paramref name="destination"/>.
    ///
    /// <b>Passing <c>true</c> for <paramref name="allowRemoteImageDownload"/> will fail in an
    /// air-gapped or otherwise offline environment</b>; see
    /// <see cref="HtmlToDocxConverter.ConvertAsync(string, bool, CancellationToken)"/>.
    /// </summary>
    /// <param name="html">The markup to convert.</param>
    /// <param name="allowRemoteImageDownload">Whether to fetch images named by absolute URLs.</param>
    /// <param name="destination">The stream the PDF is written to.</param>
    /// <param name="ct">Cancels the conversion and the write to <paramref name="destination"/>.</param>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> or <paramref name="destination"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="destination"/> is not writable.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted or written.</exception>
    public static async Task ConvertAsync(
        string html, bool allowRemoteImageDownload, Stream destination, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(html);
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        // Still a composition of the other two converters, not a third conversion: the HTML stage
        // builds the package and the DOCX stage renders it. The difference from the byte[] path is
        // that the package is handed over as the buffer it already is, so the intermediate .docx is
        // never serialised to an array and read back, and the PDF is never buffered at all.
        using var docx = await HtmlToDocxConverter
            .BuildPackageAsync(html, allowRemoteImageDownload, ct)
            .ConfigureAwait(false);

        ct.ThrowIfCancellationRequested();
        await DocxToPdfConverter.RenderAsync(docx, destination, ct).ConfigureAwait(false);
    }

    /// <summary>Converts <paramref name="html"/> and writes the PDF to <paramref name="outputPath"/>.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="outputPath"/> is blank.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted.</exception>
    public static async Task ConvertToFileAsync(string html, string outputPath, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        var pdf = await ConvertAsync(html, ct);
        await File.WriteAllBytesAsync(outputPath, pdf, ct);
    }
}
