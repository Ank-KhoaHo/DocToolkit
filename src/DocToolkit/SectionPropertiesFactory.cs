using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit;

/// <summary>
/// The one place a <see cref="PageSetup"/> becomes a <c>w:sectPr</c>.
///
/// Single site on purpose, the same reasoning as <c>WorkbookEditor.SetCellValue</c>: both DOCX
/// producers — <see cref="DocxDocumentWriter"/> and <see cref="HtmlToDocxConverter"/> — go through
/// here, so they cannot disagree about what a page setup means. A second conversion site is how
/// that guarantee gets lost.
///
/// <see cref="PageSetup"/> validates; this converts. Keeping the two apart is what lets that type's
/// tests run without OpenXml.
/// </summary>
internal static class SectionPropertiesFactory
{
    /// <summary>
    /// Word's own default distance from the page edge to the header and footer. Written explicitly
    /// rather than left absent: an omitted <c>w:header</c> is not "Word's default", it is 0, which
    /// puts a header hard against the paper edge.
    /// </summary>
    private const uint HeaderFooterTwentieths = 720;

    /// <summary>
    /// Builds the <c>w:sectPr</c> for <paramref name="page"/>. The caller appends it as the
    /// <b>last</b> child of <c>w:body</c> — anywhere else and Word declares the file corrupt.
    /// </summary>
    public static SectionProperties Build(PageSetup page)
    {
        ArgumentNullException.ThrowIfNull(page);

        var size = new PageSize
        {
            Width = (UInt32Value)(uint)ToTwentieths(page.WidthPoints),
            Height = (UInt32Value)(uint)ToTwentieths(page.HeightPoints),

            // Word reads the dimensions, but its page-setup UI and several renderers read this. A
            // landscape page still claiming portrait is a document that disagrees with itself.
            Orient = page.WidthPoints > page.HeightPoints
                ? PageOrientationValues.Landscape
                : PageOrientationValues.Portrait,
        };

        // Top and Bottom are Int32Value while Right, Left, Header, Footer and Gutter are
        // UInt32Value. That asymmetry is the ECMA schema's, not a mistake to tidy: a top margin may
        // be negative (content above the header), a left margin may not.
        var margin = new PageMargin
        {
            Top = ToTwentieths(page.TopPoints),
            Right = (UInt32Value)(uint)ToTwentieths(page.RightPoints),
            Bottom = ToTwentieths(page.BottomPoints),
            Left = (UInt32Value)(uint)ToTwentieths(page.LeftPoints),
            Header = HeaderFooterTwentieths,
            Footer = HeaderFooterTwentieths,
            Gutter = 0U,
        };

        return new SectionProperties(size, margin);
    }

    /// <summary>
    /// Points to the twentieths of a point OOXML stores. Rounded rather than truncated: the input
    /// is already an approximation of a physical measurement, and truncating loses a whole
    /// twentieth on values like A4's 595.2756 pt.
    ///
    /// <see cref="PageSetup"/> has already refused anything that would overflow.
    /// </summary>
    public static int ToTwentieths(double points) =>
        (int)Math.Round(points * 20, MidpointRounding.AwayFromZero);
}
