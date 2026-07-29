using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

namespace DocToolkit;

/// <summary>Renders a Word (.docx) package to PDF. Pure managed - no browser, no LibreOffice.</summary>
public static class DocxToPdfConverter
{
    /// <summary>Renders the .docx in <paramref name="docx"/> and returns PDF bytes.</summary>
    public static byte[] Convert(byte[] docx)
    {
        ArgumentNullException.ThrowIfNull(docx);
        if (docx.Length == 0)
            throw new ArgumentException("DOCX content was empty.", nameof(docx));

        try
        {
            // Copy into an expandable stream: OfficeIMO opens the package read/write.
            using var input = new MemoryStream();
            input.Write(docx, 0, docx.Length);
            input.Position = 0;

            using var word = WordDocument.Load(input);
            return word.ToPdf();
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException("Failed to render DOCX to PDF.", ex);
        }
    }

    /// <summary>Renders <paramref name="inputPath"/> to a PDF at <paramref name="outputPath"/>.</summary>
    public static void ConvertFile(string inputPath, string outputPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        File.WriteAllBytes(outputPath, Convert(File.ReadAllBytes(inputPath)));
    }
}
