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

        using var ms = new MemoryStream();
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
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            throw new DocumentConversionException("Failed to convert HTML to DOCX.", ex);
        }

        // ToArray() is valid after the package is disposed - MemoryStream keeps its buffer.
        return ms.ToArray();
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
