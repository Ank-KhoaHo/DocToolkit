using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

namespace DocToolkit.Tests;

/// <summary>
/// The renderer driven with the caller's fonts and NO shaping - the output each converter produced
/// before A122, built here so a test can compare the converter against it.
/// </summary>
/// <remarks>
/// <b>Bytes, not extracted text, are the comparison, and that is measured rather than chosen.</b>
/// Shaped and unshaped Hebrew and Arabic PDFs extract to the identical string, so a text assertion
/// cannot tell whether shaping ran. Each method repeats exactly what its converter does to the input
/// first, so the only difference left is the shaping itself.
/// </remarks>
internal static class UnshapedRender
{
    private static PdfOptions FontsOnly(PdfFontOptions fonts) =>
        new() { EmbeddedFontFallbacks = fonts.ToFallbackSet() };

    private static PdfResourcePolicy Policy() =>
        new() { AllowRemoteResourceResolution = false, AllowLocalFileAccess = false };

    private static MemoryStream Expandable(byte[] bytes)
    {
        var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        return stream;
    }

    public static byte[] Docx(byte[] docx, PdfFontOptions fonts)
    {
        using var input = Expandable(ListMarkerSubstitution.Apply(docx));
        using var word = WordDocument.Load(input);
        return word.ToPdfBytes(new WordToPdfOptions { PdfOptions = FontsOnly(fonts) });
    }

    public static byte[] Xlsx(byte[] xlsx, PdfFontOptions fonts)
    {
        using var input = Expandable(xlsx);
        using var workbook = ExcelDocument.Load(input);
        workbook.Calculate();
        return workbook.ToPdfBytes(new ExcelToPdfOptions { ResourcePolicy = Policy(), PdfOptions = FontsOnly(fonts) });
    }

    public static byte[] Pptx(byte[] pptx, PdfFontOptions fonts)
    {
        using var input = Expandable(pptx);
        using var presentation = PowerPointPresentation.Load(input);
        return presentation.ToPdfBytes(new PowerPointToPdfOptions { ResourcePolicy = Policy(), PdfOptions = FontsOnly(fonts) });
    }
}
