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
    public static Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)
        => ConvertAsync(html, allowRemoteImageDownload: false, ct);

    /// <summary>
    /// Converts <paramref name="html"/> to the bytes of a .docx package, optionally downloading
    /// and embedding images referenced by absolute <c>http</c>/<c>https</c> URLs.
    ///
    /// <b>Passing <c>true</c> for <paramref name="allowRemoteImageDownload"/> will fail in an
    /// air-gapped or otherwise offline environment.</b> It makes this method issue outbound HTTP
    /// requests to whatever hosts the markup names, on the calling thread's time budget, and a
    /// host that does not serve the image fails the whole conversion - on a machine with no route
    /// to the internet that failure arrives only after the OS connect timeout, once per image.
    /// This is the only API on DocToolkit that opens a network connection; everything else,
    /// including <paramref name="allowRemoteImageDownload"/> left <c>false</c>, is offline.
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
        using var package = await BuildPackageAsync(html, allowRemoteImageDownload, ct).ConfigureAwait(false);
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
    public static Task ConvertAsync(string html, Stream destination, CancellationToken ct = default)
        => ConvertAsync(html, allowRemoteImageDownload: false, destination, ct);

    /// <summary>
    /// Converts <paramref name="html"/> and writes the .docx to <paramref name="destination"/>,
    /// optionally downloading and embedding images referenced by absolute <c>http</c>/<c>https</c>
    /// URLs.
    ///
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought - it belongs to the caller, and may be write-only and
    /// forward-only, such as an HTTP response body.
    ///
    /// <b>Passing <c>true</c> for <paramref name="allowRemoteImageDownload"/> will fail in an
    /// air-gapped or otherwise offline environment</b>; see
    /// <see cref="ConvertAsync(string, bool, CancellationToken)"/> for what that costs.
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

        using var package = await BuildPackageAsync(html, allowRemoteImageDownload, ct).ConfigureAwait(false);
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
        string html, bool allowRemoteImageDownload, CancellationToken ct)
    {
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
                //  2. The resource loader itself. HtmlToOpenXml's DefaultWebRequest speaks http,
                //     https and file, and downloads through a process-wide static HttpClient whose
                //     headers it mutates per request (not thread-safe in 3.5.0). Handing it
                //     OfflineResourceLoader instead means the object capable of fetching is never
                //     constructed, so "no network" survives a future change of heart about what
                //     EmbedDataUriOnly means.
                var converter = new HtmlConverter(
                    mainPart,
                    allowRemoteImageDownload ? null : OfflineResourceLoader.Instance)
                {
                    ImageProcessing = allowRemoteImageDownload
                        ? ImageProcessingMode.Embed
                        : ImageProcessingMode.EmbedDataUriOnly,
                };

                await converter.ParseBody(html, ct);
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
    public static async Task ConvertToFileAsync(string html, string outputPath, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        var bytes = await ConvertAsync(html, ct);
        await File.WriteAllBytesAsync(outputPath, bytes, ct);
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
