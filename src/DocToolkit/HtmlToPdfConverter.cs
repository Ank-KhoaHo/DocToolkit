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
    /// <summary>Converts <paramref name="html"/> straight to PDF bytes.</summary>
    public static async Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(html);
        var docx = await HtmlToDocxConverter.ConvertAsync(html, ct);
        ct.ThrowIfCancellationRequested();
        return DocxToPdfConverter.Convert(docx);
    }

    /// <summary>Converts <paramref name="html"/> and writes the PDF to <paramref name="outputPath"/>.</summary>
    public static async Task ConvertToFileAsync(string html, string outputPath, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        var pdf = await ConvertAsync(html, ct);
        await File.WriteAllBytesAsync(outputPath, pdf, ct);
    }
}
