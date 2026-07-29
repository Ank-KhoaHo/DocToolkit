using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlToOpenXml;

namespace DocToolkit;

/// <summary>Converts an HTML fragment into a Word (.docx) package.</summary>
public static class HtmlToDocxConverter
{
    /// <summary>Converts <paramref name="html"/> to the bytes of a .docx package.</summary>
    public static async Task<byte[]> ConvertAsync(string html, CancellationToken ct = default)
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
                var converter = new HtmlConverter(mainPart);
                await converter.ParseBody(html);
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

    /// <summary>Converts <paramref name="html"/> and writes the .docx to <paramref name="outputPath"/>.</summary>
    public static async Task ConvertToFileAsync(string html, string outputPath, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        var bytes = await ConvertAsync(html, ct);
        await File.WriteAllBytesAsync(outputPath, bytes, ct);
    }
}
