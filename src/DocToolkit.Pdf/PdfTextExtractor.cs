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
    /// <c>DocxEditor.ExtractText</c> two days before this was written, and a
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
            throw new DocumentConversionException(UnreadableMessage, ex);
        }
    }

    /// <summary>
    /// Each page's words with their positions, in document order.
    /// </summary>
    /// <remarks>
    /// <b>Uses <c>page.GetWords()</c>, not the text extractor above.</b> The two answer different
    /// questions and neither substitutes for the other: <c>ContentOrderTextExtractor</c> returns a
    /// reading-ordered string and discards geometry, while <c>GetWords</c> keeps each word's
    /// rectangle. Deriving one from the other would mean re-segmenting a string that no longer
    /// knows where anything was.
    ///
    /// PdfPig's rectangle is translated into <see cref="PdfBounds"/> here rather than surfaced, so
    /// this stays the only file referencing PdfPig.
    /// </remarks>
    public static IReadOnlyList<IReadOnlyList<PdfWord>> Words(byte[] pdf)
    {
        try
        {
            using var document = PdfDocument.Open(pdf);

            var pages = new List<IReadOnlyList<PdfWord>>(document.NumberOfPages);
            foreach (var page in document.GetPages())
            {
                var words = new List<PdfWord>();
                foreach (var word in page.GetWords())
                {
                    var b = word.BoundingBox;
                    words.Add(new PdfWord(word.Text, new PdfBounds(b.Left, b.Bottom, b.Width, b.Height)));
                }

                pages.Add(words);
            }

            return pages;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException(UnreadableMessage, ex);
        }
    }

    /// <summary>
    /// Each page's images, in document order.
    /// </summary>
    /// <remarks>
    /// <b>A failure to decode ONE image is not a failure of the call.</b> <c>TryGetPng</c> returns
    /// false for colour spaces and filters it cannot re-encode, and that image arrives with a null
    /// <see cref="PdfImage.Png"/> while every other image on the page still arrives. Throwing
    /// instead would let one exotic image cost the caller the whole document, which is the
    /// behaviour this deliberately avoids.
    ///
    /// <c>IPdfImage.Bounds</c> is <c>[Obsolete]</c> in PdfPig 0.1.16; <c>BoundingBox</c> is the
    /// replacement and is what this reads.
    /// </remarks>
    public static IReadOnlyList<IReadOnlyList<PdfImage>> Images(byte[] pdf)
    {
        try
        {
            using var document = PdfDocument.Open(pdf);

            var pages = new List<IReadOnlyList<PdfImage>>(document.NumberOfPages);
            foreach (var page in document.GetPages())
            {
                var images = new List<PdfImage>();
                foreach (var image in page.GetImages())
                {
                    var b = image.BoundingBox;
                    images.Add(new PdfImage(
                        image.TryGetPng(out var png) ? png : null,
                        image.WidthInSamples,
                        image.HeightInSamples,
                        new PdfBounds(b.Left, b.Bottom, b.Width, b.Height)));
                }

                pages.Add(images);
            }

            return pages;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException(UnreadableMessage, ex);
        }
    }

    /// <summary>
    /// The one wording for "this is not a readable PDF", shared by every reader in this file.
    /// </summary>
    /// <remarks>
    /// Extracted when <see cref="Words"/> and <see cref="Images"/> joined <see cref="Pages"/>:
    /// three copies of a user-facing sentence is three chances for them to drift, and a caller
    /// matching on the message would then see a different one depending on which member failed.
    /// </remarks>
    private const string UnreadableMessage =
        "Failed to read the PDF. This usually means the PDF is password-protected, "
        + "truncated, or not actually a PDF — check the source bytes.";

}
