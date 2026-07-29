using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;

namespace DocToolkit;

/// <summary>Converts an HTML fragment into a Word (.docx) package.</summary>
public static class HtmlToDocxConverter
{
    /// <summary>
    /// Converts <paramref name="html"/> to the bytes of a .docx package.
    ///
    /// <b>No network access.</b> Images referenced by an absolute <c>http</c>/<c>https</c> URL are
    /// skipped rather than downloaded; only <c>data:</c> URI images are embedded. A byte[]-in,
    /// byte[]-out conversion that quietly fetched whatever URL happened to be in the markup would
    /// hand every caller an SSRF reach and an unbounded hang, so remote fetching is opt-in via
    /// <see cref="ConvertAsync(string, bool, CancellationToken)"/>.
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
    /// Passing <c>true</c> for <paramref name="allowRemoteImageDownload"/> makes this method issue
    /// outbound HTTP requests to whatever hosts the markup names, on the calling thread's time
    /// budget. Only do that for markup you trust, and prefer to bound it with
    /// <paramref name="ct"/>. A host that fails to serve the image fails the whole conversion.
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

                var converter = new HtmlConverter(mainPart)
                {
                    // HtmlToOpenXml defaults to ImageProcessingMode.Embed, which downloads remote
                    // images. A redistributable package must not make network calls a caller did
                    // not ask for.
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
}
