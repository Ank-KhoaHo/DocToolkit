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
    /// <remarks>
    /// The document is laid out on <see cref="PageSetup.A4"/>. Use
    /// <see cref="ConvertAsync(string, PageSetup, CancellationToken)"/> for anything else.
    /// </remarks>
    public static Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)
        => ConvertAsync(html, allowRemoteImageDownload: false, ct);

    /// <summary>
    /// Converts <paramref name="html"/> to PDF bytes, laid out on <paramref name="page"/>.
    /// Remote images are not downloaded.
    ///
    /// Page setup reaches the PDF because this pivots through DOCX and OfficeIMO honours the
    /// document's <c>w:sectPr</c> - measured, and pinned by <c>PageSetupOutputTests</c> rather
    /// than assumed, since an OfficeIMO upgrade could revert it with every DOCX test still green.
    /// </summary>
    /// <param name="html">The markup to convert.</param>
    /// <param name="page">The page size, orientation and margins.</param>
    /// <param name="ct">Cancels the conversion.</param>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> or <paramref name="page"/> is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted.</exception>
    public static async Task<byte[]> ConvertAsync(
        string html, PageSetup page, CancellationToken ct = default)
    {
        byte[] docx = await HtmlToDocxConverter.ConvertAsync(html, page, ct).ConfigureAwait(false);
        return DocxToPdfConverter.Convert(docx);
    }

    /// <summary>
    /// Converts <paramref name="html"/> straight to PDF bytes, optionally downloading and
    /// embedding images referenced by absolute <c>http</c>/<c>https</c> URLs.
    ///
    /// Passing <c>true</c> for <paramref name="allowRemoteImageDownload"/> still succeeds in an
    /// air-gapped or otherwise offline environment: the HTML stage refuses loopback, private and
    /// link-local hosts and caps every fetch at 10 seconds and 5 MB, and a host that cannot be
    /// reached simply leaves that image out of the result rather than failing the conversion. See
    /// <see cref="HtmlToDocxConverter.ConvertAsync(string, bool, CancellationToken)"/> for what
    /// this does and does not reach, including why it can never reach a private or internal host.
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
    /// Converts <paramref name="html"/> straight to PDF bytes, downloading and embedding images
    /// referenced by absolute <c>http</c>/<c>https</c> URLs, bounded by <paramref name="options"/>.
    ///
    /// <b>This still succeeds in an air-gapped or otherwise offline environment</b>: an
    /// unreachable host is skipped, not fatal; see
    /// <see cref="HtmlToDocxConverter.ConvertAsync(string, RemoteImageOptions, CancellationToken)"/>
    /// for what <paramref name="options"/> does and does not bound. This method only composes that
    /// HTML stage with the DOCX-to-PDF render stage - the fetch itself happens there.
    /// </summary>
    /// <param name="html">The markup to convert.</param>
    /// <param name="options">Bounds on the remote-image fetches this conversion is allowed to make.</param>
    /// <param name="ct">Cancels the conversion, including any in-flight image fetch.</param>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> or <paramref name="options"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="options"/> has a <see cref="RemoteImageOptions.Timeout"/> or
    /// <see cref="RemoteImageOptions.MaxBytesPerImage"/> that is not greater than zero.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="options"/>' <see cref="RemoteImageOptions.AllowedHosts"/> contains a blank entry.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted.</exception>
    public static async Task<byte[]> ConvertAsync(
        string html, RemoteImageOptions options, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(html);
        ArgumentNullException.ThrowIfNull(options);
        options.Validate();
        var docx = await HtmlToDocxConverter.ConvertAsync(html, options, ct);
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
    /// <remarks>
    /// The document is laid out on <see cref="PageSetup.A4"/>. Use
    /// <see cref="ConvertAsync(string, PageSetup, Stream, CancellationToken)"/> for anything else.
    /// </remarks>
    public static Task ConvertAsync(string html, Stream destination, CancellationToken ct = default)
        => ConvertAsync(html, allowRemoteImageDownload: false, destination, ct);

    /// <summary>
    /// Converts <paramref name="html"/> and writes the PDF, laid out on <paramref name="page"/>,
    /// to <paramref name="destination"/>. Remote images are not downloaded.
    ///
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought. As with the other PDF paths it is handed to
    /// OfficeIMO's own writer rather than buffered whole.
    /// </summary>
    /// <param name="html">The markup to convert.</param>
    /// <param name="page">The page size, orientation and margins.</param>
    /// <param name="destination">The stream the PDF is written to.</param>
    /// <param name="ct">Cancels the conversion and the write to <paramref name="destination"/>.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="destination"/> is not writable.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted or written.</exception>
    public static async Task ConvertAsync(
        string html, PageSetup page, Stream destination, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(html);
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var docx = await HtmlToDocxConverter
            .BuildPackageAsync(html, null, page, ct)
            .ConfigureAwait(false);

        ct.ThrowIfCancellationRequested();
        await DocxToPdfConverter.RenderAsync(docx, destination, ct).ConfigureAwait(false);
    }

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
    /// Passing <c>true</c> for <paramref name="allowRemoteImageDownload"/> still succeeds in an
    /// air-gapped or otherwise offline environment; see
    /// <see cref="HtmlToDocxConverter.ConvertAsync(string, bool, CancellationToken)"/> for what
    /// this does and does not reach, including why it can never reach a private or internal host.
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
            .BuildPackageAsync(html, allowRemoteImageDownload ? new RemoteImageOptions() : null, PageSetup.A4, ct)
            .ConfigureAwait(false);

        ct.ThrowIfCancellationRequested();
        await DocxToPdfConverter.RenderAsync(docx, destination, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Converts <paramref name="html"/> and writes the PDF to <paramref name="destination"/>,
    /// downloading and embedding images referenced by absolute <c>http</c>/<c>https</c> URLs,
    /// bounded by <paramref name="options"/>.
    ///
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought - it belongs to the caller, and may be write-only and
    /// forward-only, such as an HTTP response body. The PDF is written straight through as the
    /// renderer produces it, so a failure part-way leaves whatever had already been produced on
    /// <paramref name="destination"/>.
    ///
    /// <b>This still succeeds in an air-gapped or otherwise offline environment</b>: an
    /// unreachable host is skipped, not fatal; see
    /// <see cref="HtmlToDocxConverter.ConvertAsync(string, RemoteImageOptions, CancellationToken)"/>
    /// for what <paramref name="options"/> does and does not bound. This is still a composition of
    /// the other two converters, not a third conversion: the HTML stage builds the package, bounded
    /// by <paramref name="options"/>, and the DOCX stage renders it.
    /// </summary>
    /// <param name="html">The markup to convert.</param>
    /// <param name="options">Bounds on the remote-image fetches this conversion is allowed to make.</param>
    /// <param name="destination">The stream the PDF is written to.</param>
    /// <param name="ct">Cancels the conversion and the write to <paramref name="destination"/>.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="html"/>, <paramref name="options"/> or <paramref name="destination"/> is null.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="options"/> has a <see cref="RemoteImageOptions.Timeout"/> or
    /// <see cref="RemoteImageOptions.MaxBytesPerImage"/> that is not greater than zero.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="destination"/> is not writable, or <paramref name="options"/>'
    /// <see cref="RemoteImageOptions.AllowedHosts"/> contains a blank entry.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted or written.</exception>
    public static async Task ConvertAsync(
        string html, RemoteImageOptions options, Stream destination, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(html);
        ArgumentNullException.ThrowIfNull(options);
        options.Validate();
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var docx = await HtmlToDocxConverter
            .BuildPackageAsync(html, options, PageSetup.A4, ct)
            .ConfigureAwait(false);

        ct.ThrowIfCancellationRequested();
        await DocxToPdfConverter.RenderAsync(docx, destination, ct).ConfigureAwait(false);
    }

    /// <summary>Converts <paramref name="html"/> and writes the PDF to <paramref name="outputPath"/>.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="outputPath"/> is blank.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted.</exception>
    /// <remarks>
    /// The document is laid out on <see cref="PageSetup.A4"/>. Use
    /// <see cref="ConvertToFileAsync(string, PageSetup, string, CancellationToken)"/> for anything else.
    /// </remarks>
    public static async Task ConvertToFileAsync(string html, string outputPath, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        var pdf = await ConvertAsync(html, ct);
        await File.WriteAllBytesAsync(outputPath, pdf, ct);
    }

    /// <summary>
    /// Converts <paramref name="html"/> and writes the PDF, laid out on <paramref name="page"/>,
    /// to <paramref name="outputPath"/>.
    /// </summary>
    /// <param name="html">The markup to convert.</param>
    /// <param name="page">The page size, orientation and margins.</param>
    /// <param name="outputPath">Where to write the PDF. Overwritten if it exists.</param>
    /// <param name="ct">Cancels the conversion and the write.</param>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> or <paramref name="page"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="outputPath"/> is blank.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted.</exception>
    public static async Task ConvertToFileAsync(
        string html, PageSetup page, string outputPath, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var pdf = await ConvertAsync(html, page, ct).ConfigureAwait(false);
        await File.WriteAllBytesAsync(outputPath, pdf, ct).ConfigureAwait(false);
    }
}
