using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace DocToolkit;

/// <summary>
/// The only file in this library that references PdfPig.
///
/// Kept separate for a concrete reason, not tidiness: <c>PdfDocument</c> exists in BOTH
/// <c>PdfSharp.Pdf</c> and <c>UglyToad.PdfPig</c>, and <see cref="PdfEditor"/> already uses
/// PDFsharp's. Mixing them in one file would force an alias onto every existing line.
///
/// It also makes the ownership boundary physical: PDFsharp owns the document's structure - pages,
/// merge, rotate, metadata - and PdfPig owns its content. Neither does the other's job.
/// </summary>
internal static class PdfTextExtractor
{
    /// <summary>
    /// Each page's text, in document order.
    /// </summary>
    /// <remarks>
    /// <b>Uses <c>ContentOrderTextExtractor</c>, never <c>page.Text</c>.</b> Measured 2026-08-12:
    /// <c>page.Text</c> concatenates with no separator, returning
    /// <c>"Acme CorporationInvoice 42"</c> where the correct result is
    /// <c>"Acme Corporation\nInvoice 42"</c>. That is exactly the defect fixed in
    /// <see cref="DocxEditor.ExtractText(byte[])"/> two days before this was written, and a
    /// substring assertion would not notice it.
    /// </remarks>
    public static IReadOnlyList<string> Pages(byte[] pdf)
    {
        try
        {
            using var document = PdfDocument.Open(pdf);

            var pages = new List<string>(document.NumberOfPages);
            foreach (var page in document.GetPages())
            {
                pages.Add(ContentOrderTextExtractor.GetText(page) ?? string.Empty);
            }

            return pages;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read the PDF.", ex);
        }
    }
}
