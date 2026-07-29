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
