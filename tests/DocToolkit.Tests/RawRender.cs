using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

namespace DocToolkit.Tests;

/// <summary>
/// Renders a DOCX with OfficeIMO exactly as <see cref="DocxToPdfConverter"/> does and none of the
/// package's repairs, to tell what the renderer itself says about a document (A123).
/// </summary>
/// <remarks>
/// <b>The call must be the converter's own.</b> Its <c>Render</c> calls <c>ToPdfBytes()</c> with no
/// options when there are no fonts, and <c>ToPdfBytes(options)</c> with only the fonts otherwise, and
/// its remarks say an empty <c>WordToPdfOptions</c> is not assumed equivalent to none. A fixture
/// "refused before" under a different call would prove nothing about the path users take.
/// </remarks>
internal static class RawRender
{
    /// <summary>The PDF the renderer produces for <paramref name="docx"/>; throws what it refuses with.</summary>
    internal static byte[] Pdf(byte[] docx, PdfFontOptions? fonts = null)
    {
        using var input = new MemoryStream();
        input.Write(docx, 0, docx.Length);
        input.Position = 0;
        using var word = WordDocument.Load(input);

        // Mirrors DocxToPdfConverter.Render and SaveOptions.
        WordToPdfOptions? options = fonts is null ? null : new WordToPdfOptions { PdfOptions = fonts.ToPdfOptions() };
        return options is null ? word.ToPdfBytes() : word.ToPdfBytes(options);
    }

    /// <summary>What the renderer says about these bytes as they are, or null if it renders them.</summary>
    internal static Exception? Refusal(byte[] docx, PdfFontOptions? fonts = null)
    {
        try
        {
            Pdf(docx, fonts);
            return null;
        }
        catch (Exception ex)
        {
            return ex;
        }
    }
}
