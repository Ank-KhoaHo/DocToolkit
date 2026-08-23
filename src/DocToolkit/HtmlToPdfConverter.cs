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
    /// Converts <paramref name="html"/> straight to PDF bytes, using <paramref name="fonts"/> for
    /// characters the renderer cannot otherwise encode.
    /// </summary>
    /// <remarks>
    /// <b>Whether a page containing non-Latin text renders otherwise depends on the machine.</b> The
    /// renderer falls back to whatever fonts the host happens to have, so the same page converts on
    /// one and is refused on another - measured, a Windows box offers fallbacks that do not cover
    /// Cyrillic. Supplying the font takes the machine out of the answer.
    ///
    /// Nothing is fetched, exactly as with <see cref="ConvertAsync(string, CancellationToken)"/>:
    /// the font comes from the caller as bytes. See <see cref="PdfFontOptions"/> for the one side
    /// effect worth knowing about.
    ///
    /// The document is laid out on <see cref="PageSetup.A4"/>. Use
    /// <see cref="ConvertAsync(string, PageSetup, CancellationToken)"/> for anything else.
    /// </remarks>
    /// <param name="html">The markup to convert.</param>
    /// <param name="fonts">Fonts to fall back to.</param>
    /// <param name="ct">Cancels the conversion.</param>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> or <paramref name="fonts"/> is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted.</exception>
    public static async Task<byte[]> ConvertAsync(
        string html, PdfFontOptions fonts, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(html);
        ArgumentNullException.ThrowIfNull(fonts);

        return await HtmlForPdf.RenderAsync(
            html, h => HtmlToDocxConverter.ConvertAsync(h, PageSetup.A4, ct), fonts, ct)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Converts <paramref name="html"/> to PDF bytes, applying page setup, remote-image policy and
    /// fonts <b>together</b>.
    /// </summary>
    /// <remarks>
    /// <b>This is the only overload that can express all three at once</b>, which is why it exists.
    /// The others each fix two of the axes at their defaults, so a caller needing fonts on a
    /// non-A4 page - or fonts alongside remote images - had no signature to call and
    /// <c>DocToolkitOptions.Fonts</c> could not reach the HTML path at all.
    ///
    /// <para><b>Remote fetching is opt-in and stays that way.</b>
    /// <see cref="HtmlToPdfOptions.RemoteImage"/> being <see langword="null"/> - its default - opens
    /// no socket, so this overload is safe in an air-gapped environment exactly like the
    /// others.</para>
    /// </remarks>
    /// <param name="html">The markup to convert.</param>
    /// <param name="options">Page setup, remote-image policy and fonts. See <see cref="HtmlToPdfOptions"/>.</param>
    /// <param name="ct">Cancels the conversion, including the PDF render.</param>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="options"/>' <see cref="HtmlToPdfOptions.RemoteImage"/> has an
    /// <see cref="RemoteImageOptions.AllowedHosts"/> entry that is blank.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted.</exception>
    public static async Task<byte[]> ConvertAsync(
        string html, HtmlToPdfOptions options, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(html);
        ArgumentNullException.ThrowIfNull(options);
        options.RemoteImage?.Validate();

        // Which delegate, not which converter: HTML to PDF stays a COMPOSITION of the other two
        // converters. The branch is on whether remote images were opted into, because the core
        // expresses "offline" and "bounded fetch" as two different HtmlToDocxConverter overloads
        // rather than as a nullable argument.
        Func<string, Task<byte[]>> toDocx = options.RemoteImage is null
            ? h => HtmlToDocxConverter.ConvertAsync(h, options.Page, ct)
            : h => HtmlToDocxConverter.ConvertAsync(h, options.Page, options.RemoteImage, ct);

        return await HtmlForPdf.RenderAsync(html, toDocx, options.Fonts, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Converts <paramref name="html"/> straight to PDF bytes, laid out on
    /// <see cref="PageSetup.A4"/>.
    ///
    /// <b>No network access, and safe in an air-gapped environment</b> - nothing the markup
    /// references is fetched. See <see cref="ConvertAsync(string, RemoteImageOptions, CancellationToken)"/>
    /// to opt in.
    /// </summary>
    /// <param name="html">The markup to convert.</param>
    /// <param name="ct">Cancels the conversion.</param>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted.</exception>
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
    /// <example>
    /// <code source="../../tests/DocToolkit.Tests/DocumentationExamples.cs" region="HtmlToPdf"/>
    /// </example>
    public static async Task<byte[]> ConvertAsync(
        string html, PageSetup page, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(html);
        return await HtmlForPdf.RenderAsync(html, h => HtmlToDocxConverter.ConvertAsync(h, page, ct), ct: ct).ConfigureAwait(false);
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
        return await HtmlForPdf.RenderAsync(html, h => HtmlToDocxConverter.ConvertAsync(h, allowRemoteImageDownload, ct), ct: ct).ConfigureAwait(false);
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
        return await HtmlForPdf.RenderAsync(html, h => HtmlToDocxConverter.ConvertAsync(h, options, ct), ct: ct).ConfigureAwait(false);
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
    /// The PDF is rendered whole and then written, so a failure leaves
    /// <paramref name="destination"/> UNTOUCHED rather than carrying a truncated document.
    /// Until 2026-08-20 it was written straight through as the renderer produced it; that
    /// prevented the repair-and-retry the array overloads apply, and measurably diverged from
    /// them - see HtmlToPdfConverter's private EmitAsync for the numbers.
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

        await EmitAsync(
            await HtmlForPdf.RenderAsync(html, h => HtmlToDocxConverter.ConvertAsync(h, page, ct), ct: ct)
                .ConfigureAwait(false),
            destination, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Writes a finished PDF onto a caller's stream.
    /// </summary>
    /// <remarks>
    /// <b>The reason every <c>Stream</c> overload here buffers rather than writing straight
    /// through.</b> Until 2026-08-20 they handed the package to OfficeIMO's writer directly, which
    /// meant they could not go through <see cref="HtmlForPdf"/> at all: its repairs retry a failed
    /// render, and a retry cannot un-write bytes already on somebody's HTTP response body. So the
    /// <c>byte[]</c> overloads got the repairs and these did not.
    ///
    /// <b>That was a real divergence, not a theoretical one.</b> Measured over real files: a page
    /// whose internal links use <c>&lt;a name&gt;</c> - 27 of 181 real .gov pages - converted
    /// through <c>ConvertAsync(html)</c> and was refused through <c>ConvertAsync(html,
    /// destination)</c>, and 4 of 99 real Word documents did the same on the DOCX path.
    ///
    /// <b>Buffering costs one PDF of memory and buys two things.</b> Every overload now answers
    /// identically, and a failure leaves <c>destination</c> UNTOUCHED rather than carrying a
    /// truncated PDF - which is better than what the old doc comments promised. The
    /// <c>Stream</c> overloads were never a memory optimisation anyway: <c>DrainAsync</c> buffers
    /// the source whatever happens, measured at 238 MB against 233 MB for the array path.
    /// </remarks>
    private static async Task EmitAsync(byte[] pdf, Stream destination, CancellationToken ct)
    {
        using var scratch = new MemoryStream(pdf, writable: false);
        await StreamPipeline
            .EmitAsync(scratch, destination, "Failed to write the PDF. See the inner exception for details.", ct)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Converts <paramref name="html"/> and writes the PDF to <paramref name="destination"/>,
    /// optionally downloading and embedding images referenced by absolute <c>http</c>/<c>https</c>
    /// URLs.
    ///
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought - it belongs to the caller, and may be write-only and
    /// The PDF is rendered whole and then written, so a failure leaves
    /// <paramref name="destination"/> UNTOUCHED rather than carrying a truncated document.
    /// Until 2026-08-20 it was written straight through as the renderer produced it; that
    /// prevented the repair-and-retry the array overloads apply, and measurably diverged from
    /// them - see HtmlToPdfConverter's private EmitAsync for the numbers.
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

        // Still a composition of the other two converters, not a third conversion, and it now goes
        // through the SAME funnel the byte[] overloads use - see EmitAsync for why that means
        // buffering, and what it was costing not to.
        await EmitAsync(
            await HtmlForPdf.RenderAsync(
                html, h => HtmlToDocxConverter.ConvertAsync(h, allowRemoteImageDownload, ct), ct: ct)
                .ConfigureAwait(false),
            destination, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Converts <paramref name="html"/> and writes the PDF to <paramref name="destination"/>,
    /// downloading and embedding images referenced by absolute <c>http</c>/<c>https</c> URLs,
    /// bounded by <paramref name="options"/>.
    ///
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought - it belongs to the caller, and may be write-only and
    /// The PDF is rendered whole and then written, so a failure leaves
    /// <paramref name="destination"/> UNTOUCHED rather than carrying a truncated document.
    /// Until 2026-08-20 it was written straight through as the renderer produced it; that
    /// prevented the repair-and-retry the array overloads apply, and measurably diverged from
    /// them - see HtmlToPdfConverter's private EmitAsync for the numbers.
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

        await EmitAsync(
            await HtmlForPdf.RenderAsync(html, h => HtmlToDocxConverter.ConvertAsync(h, options, ct), ct: ct)
                .ConfigureAwait(false),
            destination, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Converts <paramref name="html"/> to a PDF laid out on <paramref name="page"/>, fetching
    /// remote images under <paramref name="options"/>.
    /// </summary>
    /// <remarks>
    /// The combination the other overloads cannot express: <c>(html, page)</c> always converts
    /// offline, and <c>(html, options)</c> always lays out on A4. Both silently discarded half of
    /// what a caller wanting Letter <i>and</i> an allow-list had asked for.
    /// </remarks>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
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
        string html, PageSetup page, RemoteImageOptions options, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(html);
        ArgumentNullException.ThrowIfNull(page);
        ArgumentNullException.ThrowIfNull(options);
        options.Validate();
        return await HtmlForPdf.RenderAsync(html, h => HtmlToDocxConverter.ConvertAsync(h, page, options, ct), ct: ct).ConfigureAwait(false);
    }

    /// <inheritdoc cref="ConvertAsync(string, PageSetup, RemoteImageOptions, CancellationToken)"/>
    /// <exception cref="ArgumentException">
    /// <paramref name="destination"/> is not writable, or <paramref name="options"/>'
    /// <see cref="RemoteImageOptions.AllowedHosts"/> contains a blank entry.
    /// </exception>
    public static async Task ConvertAsync(
        string html, PageSetup page, RemoteImageOptions options, Stream destination,
        CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(html);
        ArgumentNullException.ThrowIfNull(page);
        ArgumentNullException.ThrowIfNull(options);
        options.Validate();
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        await EmitAsync(
            await HtmlForPdf.RenderAsync(html, h => HtmlToDocxConverter.ConvertAsync(h, page, options, ct), ct: ct)
                .ConfigureAwait(false),
            destination, ct).ConfigureAwait(false);
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
        var pdf = await ConvertAsync(html, ct).ConfigureAwait(false);
        await File.WriteAllBytesAsync(outputPath, pdf, ct).ConfigureAwait(false);
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
        ArgumentNullException.ThrowIfNull(html);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var pdf = await ConvertAsync(html, page, ct).ConfigureAwait(false);
        await File.WriteAllBytesAsync(outputPath, pdf, ct).ConfigureAwait(false);
    }

    /// <inheritdoc cref="ConvertAsync(string, HtmlToPdfOptions, CancellationToken)" path="/summary|/remarks"/>
    /// <param name="html">The markup to convert.</param>
    /// <param name="options">Page setup, remote-image policy and fonts.</param>
    /// <param name="destination">
    /// The stream the PDF is written to, from its current position. <b>Not</b> disposed, closed or
    /// sought - it belongs to the caller and may be write-only and forward-only.
    /// </param>
    /// <param name="ct">Cancels the conversion and the write to <paramref name="destination"/>.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="destination"/> is not writable.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted or written.</exception>
    public static async Task ConvertAsync(
        string html, HtmlToPdfOptions options, Stream destination, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(html);
        ArgumentNullException.ThrowIfNull(options);
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        await EmitAsync(await ConvertAsync(html, options, ct).ConfigureAwait(false), destination, ct)
            .ConfigureAwait(false);
    }

    /// <inheritdoc cref="ConvertAsync(string, HtmlToPdfOptions, CancellationToken)" path="/summary|/remarks"/>
    /// <param name="html">The markup to convert.</param>
    /// <param name="options">Page setup, remote-image policy and fonts.</param>
    /// <param name="outputPath">The file the PDF is written to. Overwritten if it exists.</param>
    /// <param name="ct">Cancels the conversion and the write.</param>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> or <paramref name="options"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="outputPath"/> is null, empty or whitespace.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted.</exception>
    public static async Task ConvertToFileAsync(
        string html, HtmlToPdfOptions options, string outputPath, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(html);
        ArgumentNullException.ThrowIfNull(options);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var pdf = await ConvertAsync(html, options, ct).ConfigureAwait(false);
        await File.WriteAllBytesAsync(outputPath, pdf, ct).ConfigureAwait(false);
    }
}
