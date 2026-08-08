using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;
using HtmlToOpenXml.IO;

namespace DocToolkit;

/// <summary>Converts an HTML fragment into a Word (.docx) package.</summary>
public static class HtmlToDocxConverter
{
    /// <summary>
    /// Converts <paramref name="html"/> to the bytes of a .docx package.
    ///
    /// <b>No network access, and safe in an air-gapped environment.</b> Nothing the markup
    /// references is fetched - not images, not stylesheets, not scripts, whether named by
    /// <c>http</c>, <c>https</c> or <c>file</c>. Only <c>data:</c> URI images are embedded. A
    /// byte[]-in, byte[]-out conversion that quietly fetched whatever URL happened to be in the
    /// markup would hand every caller an SSRF reach and an unbounded hang, so remote fetching is
    /// opt-in via <see cref="ConvertAsync(string, bool, CancellationToken)"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted.</exception>
    /// <remarks>
    /// The document is laid out on <see cref="PageSetup.A4"/>. Use
    /// <see cref="ConvertAsync(string, PageSetup, CancellationToken)"/> for anything else.
    /// </remarks>
    /// <example>
    /// <code source="../../tests/DocToolkit.Tests/DocumentationExamples.cs" region="HtmlToDocx"/>
    /// </example>
    public static Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)
        => ConvertAsync(html, allowRemoteImageDownload: false, ct);

    /// <summary>
    /// Converts <paramref name="html"/> to the bytes of a .docx package laid out on
    /// <paramref name="page"/>. Remote images are not downloaded; see
    /// <see cref="ConvertAsync(string, RemoteImageOptions, CancellationToken)"/> to opt in.
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
        using var package = await BuildPackageAsync(html, null, page, ct).ConfigureAwait(false);
        return package.ToArray();
    }

    /// <summary>
    /// Converts <paramref name="html"/> to the bytes of a .docx package, optionally downloading
    /// and embedding images referenced by absolute <c>http</c>/<c>https</c> URLs.
    ///
    /// <b>Passing <c>true</c> for <paramref name="allowRemoteImageDownload"/> routes fetches
    /// through a <see cref="RemoteImageOptions"/> with every default left in place.</b> Loopback,
    /// private and link-local hosts - including <c>169.254.169.254</c>, the cloud metadata
    /// endpoint - are refused, and every fetch is capped at 10 seconds and 5 MB. <b>A host that
    /// cannot be reached, refuses the connection, or does not serve the image is skipped: that
    /// image is left out of the result, and the conversion still succeeds</b>, at a cost of up to
    /// 10 seconds for each image it cannot reach. That includes an air-gapped or otherwise offline
    /// environment - the conversion completes, just with every remote image silently absent, one
    /// 10-second wait at a time. This is the only API on DocToolkit that opens a network
    /// connection; everything else, including <paramref name="allowRemoteImageDownload"/> left
    /// <c>false</c>, is offline.
    ///
    /// This overload can never reach a private or internal host - an intranet image server, for
    /// example - because a <c>bool</c> has no way to carry
    /// <see cref="RemoteImageOptions.AllowPrivateAddresses"/>. A caller that needs one must use
    /// <see cref="ConvertAsync(string, RemoteImageOptions, CancellationToken)"/> with
    /// <see cref="RemoteImageOptions.AllowPrivateAddresses"/> set <c>true</c>; otherwise a consumer
    /// converting intranet-hosted markup with <paramref name="allowRemoteImageDownload"/>
    /// <c>true</c> gets a document with that image quietly missing, not an exception explaining why.
    ///
    /// Only pass <c>true</c> for markup you trust, and prefer to bound it with
    /// <paramref name="ct"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted.</exception>
    public static async Task<byte[]> ConvertAsync(
        string html, bool allowRemoteImageDownload, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(html);
        ct.ThrowIfCancellationRequested();

        // ToArray() is valid after the WordprocessingDocument has been disposed - the
        // MemoryStream keeps its buffer. It does, however, allocate a second full copy of the
        // package, which is what ConvertAsync(html, destination, ct) exists to avoid.
        using var package = await BuildPackageAsync(
            html, allowRemoteImageDownload ? new RemoteImageOptions() : null, PageSetup.A4, ct).ConfigureAwait(false);
        return package.ToArray();
    }

    /// <summary>
    /// Converts <paramref name="html"/> to the bytes of a .docx package, downloading and embedding
    /// images referenced by absolute <c>http</c>/<c>https</c> URLs, bounded by
    /// <paramref name="options"/>.
    ///
    /// <b>This still succeeds in an air-gapped or otherwise offline environment.</b> A fetch that
    /// cannot leave the machine is caught the same way as a host that refuses to serve the image:
    /// that image is skipped, never the whole conversion. What <paramref name="options"/> adds
    /// over <see cref="ConvertAsync(string, bool, CancellationToken)"/> is a per-fetch timeout, a
    /// byte cap, an optional host allow-list and a block on loopback/private/link-local addresses
    /// (so a hostile document cannot use this opt-in to reach <c>169.254.169.254</c> or an internal
    /// service, unless <paramref name="options"/> sets
    /// <see cref="RemoteImageOptions.AllowPrivateAddresses"/>) - all active by default, so
    /// <c>new RemoteImageOptions()</c> already narrows the unbounded form considerably. Offline,
    /// that means every remote image is silently missing from the result, at a cost of up to
    /// <see cref="RemoteImageOptions.Timeout"/> per image - not a failed conversion. This is still
    /// not a complete SSRF defence; see <see cref="RemoteImageOptions"/>.
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
        ct.ThrowIfCancellationRequested();

        using var package = await BuildPackageAsync(html, options, PageSetup.A4, ct).ConfigureAwait(false);
        return package.ToArray();
    }

    /// <summary>
    /// Converts <paramref name="html"/> and writes the .docx to <paramref name="destination"/>.
    ///
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought - it belongs to the caller, and may be write-only and
    /// forward-only, such as an HTTP response body. Remote images are not downloaded; see
    /// <see cref="ConvertAsync(string, bool, Stream, CancellationToken)"/> to opt in.
    ///
    /// <b>No network access, and safe in an air-gapped environment</b>, exactly as for
    /// <see cref="ConvertAsync(string, CancellationToken)"/>.
    /// </summary>
    /// <param name="html">The markup to convert.</param>
    /// <param name="destination">The stream the .docx package is written to.</param>
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
    /// Converts <paramref name="html"/> and writes the .docx, laid out on <paramref name="page"/>,
    /// to <paramref name="destination"/>. Remote images are not downloaded.
    ///
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought — it belongs to the caller, and may be write-only and
    /// forward-only, such as an HTTP response body.
    /// </summary>
    /// <param name="html">The markup to convert.</param>
    /// <param name="page">The page size, orientation and margins.</param>
    /// <param name="destination">The stream the .docx package is written to.</param>
    /// <param name="ct">Cancels the conversion and the write to <paramref name="destination"/>.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="destination"/> is not writable.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted or written.</exception>
    public static async Task ConvertAsync(
        string html, PageSetup page, Stream destination, CancellationToken ct = default)
    {
        StreamPipeline.RequireWritable(destination, nameof(destination));

        using var package = await BuildPackageAsync(html, null, page, ct).ConfigureAwait(false);
        await StreamPipeline
            .EmitAsync(package, destination, "Failed to convert HTML to DOCX.", ct)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Converts <paramref name="html"/> and writes the .docx to <paramref name="destination"/>,
    /// optionally downloading and embedding images referenced by absolute <c>http</c>/<c>https</c>
    /// URLs.
    ///
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought - it belongs to the caller, and may be write-only and
    /// forward-only, such as an HTTP response body.
    ///
    /// Passing <c>true</c> for <paramref name="allowRemoteImageDownload"/> still succeeds in an
    /// air-gapped or otherwise offline environment; see
    /// <see cref="ConvertAsync(string, bool, CancellationToken)"/> for what it does and does not
    /// reach, including why it can never reach a private or internal host.
    /// </summary>
    /// <param name="html">The markup to convert.</param>
    /// <param name="allowRemoteImageDownload">Whether to fetch images named by absolute URLs.</param>
    /// <param name="destination">The stream the .docx package is written to.</param>
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

        using var package = await BuildPackageAsync(
            html, allowRemoteImageDownload ? new RemoteImageOptions() : null, PageSetup.A4, ct).ConfigureAwait(false);
        await StreamPipeline
            .EmitAsync(package, destination, "Failed to convert HTML to DOCX.", ct)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Converts <paramref name="html"/> and writes the .docx to <paramref name="destination"/>,
    /// downloading and embedding images referenced by absolute <c>http</c>/<c>https</c> URLs,
    /// bounded by <paramref name="options"/>.
    ///
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought - it belongs to the caller, and may be write-only and
    /// forward-only, such as an HTTP response body.
    ///
    /// <b>This still succeeds in an air-gapped or otherwise offline environment</b>: an
    /// unreachable host is skipped, not fatal; see
    /// <see cref="ConvertAsync(string, RemoteImageOptions, CancellationToken)"/> for what
    /// <paramref name="options"/> does and does not bound.
    /// </summary>
    /// <param name="html">The markup to convert.</param>
    /// <param name="options">Bounds on the remote-image fetches this conversion is allowed to make.</param>
    /// <param name="destination">The stream the .docx package is written to.</param>
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

        using var package = await BuildPackageAsync(html, options, PageSetup.A4, ct).ConfigureAwait(false);
        await StreamPipeline
            .EmitAsync(package, destination, "Failed to convert HTML to DOCX.", ct)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Builds the .docx package into a scratch buffer, positioned at 0.
    ///
    /// The one place the package is actually produced, so the <c>byte[]</c> overload and the
    /// <see cref="Stream"/> overload cannot drift apart, and so
    /// <see cref="HtmlToPdfConverter"/> can hand the buffer straight to the renderer instead of
    /// round-tripping it through an array on the way.
    ///
    /// <c>WordprocessingDocument.Create</c> needs a readable, writable, seekable stream - a ZIP's
    /// central directory is written at the end and the writer seeks back over its own output - so
    /// the package cannot be built directly onto a caller's forward-only destination.
    /// </summary>
    internal static async Task<MemoryStream> BuildPackageAsync(
        string html, RemoteImageOptions? options, PageSetup page, CancellationToken ct)
    {
        // Outside the try below on purpose: its catch-all wraps everything in
        // DocumentConversionException, and an argument fault must surface unwrapped.
        ArgumentNullException.ThrowIfNull(page);

        var ms = new MemoryStream();
        try
        {
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document(new Body());

                // Two independent locks on the default path, because DocToolkit's users have no
                // internet access at all and a single setting is a policy, not a guarantee:
                //
                //  1. ImageProcessing. HtmlToOpenXml defaults to ImageProcessingMode.Embed, which
                //     downloads remote images. EmbedDataUriOnly skips anything that is not a
                //     data: URI - including file:// paths, which would otherwise read the caller's
                //     disk.
                //  2. The resource loader itself. Both branches below hand HtmlConverter an
                //     explicit IWebRequest, so HtmlToOpenXml's own DefaultWebRequest - which speaks
                //     http, https and file, and downloads through a process-wide static HttpClient
                //     whose headers it mutates per request (not thread-safe in 3.5.0) - is never
                //     constructed on either path, so "no network" survives a future change of heart
                //     about what EmbedDataUriOnly means. options is null means offline:
                //     OfflineResourceLoader fetches nothing at all. options non-null means the
                //     caller opted in: a GuardedResourceLoader bounds what it fetches to options's
                //     timeout, byte cap, host allow-list and private-address block.
                var converter = new HtmlConverter(
                    mainPart,
                    options is null
                        ? OfflineResourceLoader.Instance
                        : new GuardedResourceLoader(options))
                {
                    ImageProcessing = options is null
                        ? ImageProcessingMode.EmbedDataUriOnly
                        : ImageProcessingMode.Embed,
                };

                await converter.ParseBody(html, ct);

                // HtmlToOpenXml emits no w:sectPr of its own - measured, and the reason this
                // exists: a document that states no page setup renders on whatever paper the
                // reader's Word template happens to name, so the same HTML lands on Letter in
                // the US and A4 elsewhere. Appended after ParseBody so it is the body's last
                // child, which is the only position Word accepts.
                mainPart.Document.Body!.AppendChild(SectionPropertiesFactory.Build(page));

                mainPart.Document.Save();
            }
        }
        catch (OperationCanceledException)
        {
            ms.Dispose();
            throw;
        }
        catch (Exception ex)
        {
            ms.Dispose();
            throw new DocumentConversionException("Failed to convert HTML to DOCX.", ex);
        }

        ms.Position = 0;
        return ms;
    }

    /// <summary>
    /// Converts <paramref name="html"/> and writes the .docx to <paramref name="outputPath"/>.
    /// Remote images are not downloaded; see <see cref="ConvertAsync(string, CancellationToken)"/>.
    /// </summary>
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
        var bytes = await ConvertAsync(html, ct);
        await File.WriteAllBytesAsync(outputPath, bytes, ct);
    }

    /// <summary>
    /// Converts <paramref name="html"/> and writes the .docx, laid out on <paramref name="page"/>,
    /// to <paramref name="outputPath"/>. Remote images are not downloaded.
    /// </summary>
    /// <param name="html">The markup to convert.</param>
    /// <param name="page">The page size, orientation and margins.</param>
    /// <param name="outputPath">Where to write the document. Overwritten if it exists.</param>
    /// <param name="ct">Cancels the conversion and the write.</param>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> or <paramref name="page"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="outputPath"/> is blank.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The HTML could not be converted.</exception>
    public static async Task ConvertToFileAsync(
        string html, PageSetup page, string outputPath, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var bytes = await ConvertAsync(html, page, ct).ConfigureAwait(false);
        await File.WriteAllBytesAsync(outputPath, bytes, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// The resource loader used on the no-network path: it supports no protocol and fetches
    /// nothing, ever.
    ///
    /// <see cref="IWebRequest"/> is HtmlToOpenXml's own hook for resolving a document's
    /// subresources. Supplying a refusing implementation is what turns "we selected an image mode
    /// that happens not to download" into "the component that could download was never built".
    /// Data-URI images are decoded by the parser itself and never routed through here, so
    /// self-contained documents still convert in full.
    /// </summary>
    private sealed class OfflineResourceLoader : IWebRequest
    {
        public static readonly OfflineResourceLoader Instance = new();

        /// <summary>No protocol is supported - not http, not https, not file.</summary>
        public bool SupportsProtocol(string protocol) => false;

        /// <summary>Never called, given <see cref="SupportsProtocol"/>; returns nothing if it is.</summary>
        public Task<Resource?> FetchAsync(Uri requestUri, CancellationToken cancellationToken = default)
            => Task.FromResult<Resource?>(null);
    }
}
