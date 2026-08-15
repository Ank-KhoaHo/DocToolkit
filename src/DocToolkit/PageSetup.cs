using System.Globalization;

namespace DocToolkit;

/// <summary>
/// The paper a generated document is laid out on: size, orientation and margins, all in points.
///
/// Immutable and built by factories — <see cref="A4"/>, <see cref="Letter"/> or
/// <see cref="Custom"/> — then copied by <see cref="Landscape"/> and <c>WithMargins</c>. The same
/// shape as <see cref="DocxBlock"/>, <see cref="PptxSlide"/>, <see cref="XlsxSheet"/> and
/// <see cref="XlsxFormula"/>, so the five read as a set.
///
/// <b>Points throughout</b>, because <see cref="DocxBlock.Image"/> already takes
/// <c>widthPoints</c>/<c>heightPoints</c>, and a library with two length units is a library that
/// will eventually mix them up. OOXML stores these as twentieths of a point; that conversion lives
/// in <c>SectionPropertiesFactory</c> and never reaches a caller.
///
/// <b>Only two presets.</b> A3, Legal and the rest are one line each and can be added when somebody
/// asks; <see cref="Custom"/> covers them meanwhile.
///
/// <b>There is no <c>Portrait()</c>.</b> It has no defensible answer for a square custom page —
/// "the longer side vertical" is undefined when there is no longer side. Both presets are already
/// portrait and this type is immutable, so start from the preset again.
/// </summary>
public sealed class PageSetup
{
    /// <summary>One inch. The default on every margin, and the value both presets carry.</summary>
    private const double DefaultMarginPoints = 72;

    /// <summary>
    /// The largest dimension whose twentieths-of-a-point form still fits the integer the OOXML
    /// attributes hold. <c>w:pgMar/@top</c> and <c>@bottom</c> are signed, so the signed limit is
    /// the binding one and is applied to every value here — a page this size is 745 miles across,
    /// so nothing real is being refused.
    /// </summary>
    private const double MaxPoints = int.MaxValue / 20d;

    private PageSetup(
        double widthPoints, double heightPoints,
        double topPoints, double rightPoints, double bottomPoints, double leftPoints,
        DocxHeader? header = null, DocxHeader? footer = null,
        DocxHeader? firstPageHeader = null, DocxHeader? firstPageFooter = null,
        bool hasDistinctFirstPage = false)
    {
        WidthPoints = widthPoints;
        HeightPoints = heightPoints;
        TopPoints = topPoints;
        RightPoints = rightPoints;
        BottomPoints = bottomPoints;
        LeftPoints = leftPoints;
        Header = header;
        Footer = footer;
        FirstPageHeader = firstPageHeader;
        FirstPageFooter = firstPageFooter;
        HasDistinctFirstPage = hasDistinctFirstPage;
    }

    /// <summary>
    /// ISO A4 portrait — 210 × 297 mm — with one-inch margins. The default for every producer.
    ///
    /// The one-decimal values are not sloppiness: 210 mm is 595.2756 pt, and 595.3 × 20 is the
    /// 11906 twentieths Word itself writes for A4. Rounding to 595 would write 11900 and render a
    /// page 0.3 pt narrower than every other tool's A4.
    /// </summary>
    public static PageSetup A4 { get; } = new(
        595.3, 841.9,
        DefaultMarginPoints, DefaultMarginPoints, DefaultMarginPoints, DefaultMarginPoints);

    /// <summary>US Letter portrait — 8.5 × 11 in — with one-inch margins.</summary>
    public static PageSetup Letter { get; } = new(
        612, 792,
        DefaultMarginPoints, DefaultMarginPoints, DefaultMarginPoints, DefaultMarginPoints);

    /// <summary>The page width, in points.</summary>
    public double WidthPoints { get; }

    /// <summary>The page height, in points.</summary>
    public double HeightPoints { get; }

    /// <summary>The top margin, in points.</summary>
    public double TopPoints { get; }

    /// <summary>The right margin, in points.</summary>
    public double RightPoints { get; }

    /// <summary>The bottom margin, in points.</summary>
    public double BottomPoints { get; }

    /// <summary>The left margin, in points.</summary>
    public double LeftPoints { get; }

    /// <summary>The header on every page, or null for none.</summary>
    public DocxHeader? Header { get; }

    /// <summary>The footer on every page, or null for none.</summary>
    public DocxHeader? Footer { get; }

    /// <summary>The header on page one when <see cref="HasDistinctFirstPage"/>, or null for none.</summary>
    public DocxHeader? FirstPageHeader { get; }

    /// <summary>The footer on page one when <see cref="HasDistinctFirstPage"/>, or null for none.</summary>
    public DocxHeader? FirstPageFooter { get; }

    /// <summary>
    /// Whether page one is treated differently — set by calling <see cref="WithFirstPage"/>, and
    /// emitted as <c>w:titlePg</c>.
    /// </summary>
    public bool HasDistinctFirstPage { get; }

    /// <summary>
    /// A page of the given size, in points, with one-inch margins.
    /// </summary>
    /// <param name="widthPoints">The page width. Must be positive.</param>
    /// <param name="heightPoints">The page height. Must be positive.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Either dimension is not a positive, finite number no larger than <c>int.MaxValue / 20</c>.
    /// </exception>
    /// <remarks>
    /// <b>The content area is deliberately NOT validated here</b>, unlike
    /// <see cref="WithMargins(double, double, double, double)"/> and <see cref="Landscape"/>. A page smaller than the one-inch
    /// defaults — <c>Custom(100, 400)</c> — is a legitimate starting point for
    /// <c>Custom(100, 400).WithMargins(0, 60, 0, 40)</c>, and guarding construction would make a
    /// small page impossible to build at all rather than merely awkward. Tried 2026-08-15 and
    /// reverted: it broke three tests that construct exactly that shape.
    ///
    /// The cost is real and is the caller's to avoid: a <c>Custom</c> page too small for the
    /// default margins, used without <see cref="WithMargins(double, double, double, double)"/>, produces a document Word renders
    /// blank. Nothing refuses it, because the only place that could is the conversion boundary,
    /// and moving the check there would change when a shipped API throws.
    /// </remarks>
    public static PageSetup Custom(double widthPoints, double heightPoints)
    {
        RequireDimension(widthPoints, nameof(widthPoints));
        RequireDimension(heightPoints, nameof(heightPoints));

        return new PageSetup(
            widthPoints, heightPoints,
            DefaultMarginPoints, DefaultMarginPoints, DefaultMarginPoints, DefaultMarginPoints);
    }

    /// <summary>
    /// A copy with the width and height swapped, keeping the margins as they are.
    ///
    /// This <b>swaps</b> rather than normalises, so calling it twice returns to where you started.
    /// </summary>
    /// <exception cref="ArgumentException">
    /// The swap leaves no content area — the margins were valid against the old dimensions and are
    /// not against the new ones.
    /// </exception>
    public PageSetup Landscape()
    {
        RequireContentArea(
            HeightPoints, WidthPoints, TopPoints, RightPoints, BottomPoints, LeftPoints);

        return new PageSetup(
            HeightPoints, WidthPoints,
            TopPoints, RightPoints, BottomPoints, LeftPoints,
            Header, Footer, FirstPageHeader, FirstPageFooter, HasDistinctFirstPage);
    }

    /// <summary>
    /// A copy with all four margins set to <paramref name="points"/>.
    /// </summary>
    /// <param name="points">The margin to apply on all four sides.</param>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="points"/> is negative, NaN or too large.</exception>
    /// <exception cref="ArgumentException">The margins leave no content area.</exception>
    public PageSetup WithMargins(double points) => WithMargins(points, points, points, points);

    /// <summary>
    /// A copy with the four margins set individually, clockwise from the top — the order CSS uses.
    /// </summary>
    /// <param name="topPoints">The top margin.</param>
    /// <param name="rightPoints">The right margin.</param>
    /// <param name="bottomPoints">The bottom margin.</param>
    /// <param name="leftPoints">The left margin.</param>
    /// <exception cref="ArgumentOutOfRangeException">A margin is negative, NaN or too large.</exception>
    /// <exception cref="ArgumentException">
    /// The horizontal or vertical margins together are at least the page's width or height. Such a
    /// document opens and renders as a blank page rather than failing, so it is refused here.
    /// </exception>
    public PageSetup WithMargins(
        double topPoints, double rightPoints, double bottomPoints, double leftPoints)
    {
        RequireMargin(topPoints, nameof(topPoints));
        RequireMargin(rightPoints, nameof(rightPoints));
        RequireMargin(bottomPoints, nameof(bottomPoints));
        RequireMargin(leftPoints, nameof(leftPoints));
        RequireContentArea(
            WidthPoints, HeightPoints, topPoints, rightPoints, bottomPoints, leftPoints);

        return new PageSetup(
            WidthPoints, HeightPoints,
            topPoints, rightPoints, bottomPoints, leftPoints,
            Header, Footer, FirstPageHeader, FirstPageFooter, HasDistinctFirstPage);
    }

    /// <summary>A copy carrying <paramref name="header"/> on every page.</summary>
    /// <param name="header">The header content.</param>
    /// <exception cref="ArgumentNullException"><paramref name="header"/> is null.</exception>
    public PageSetup WithHeader(DocxHeader header)
    {
        ArgumentNullException.ThrowIfNull(header);
        return With(header: header);
    }

    /// <summary>A copy carrying <paramref name="footer"/> on every page.</summary>
    /// <param name="footer">The footer content.</param>
    /// <exception cref="ArgumentNullException"><paramref name="footer"/> is null.</exception>
    public PageSetup WithFooter(DocxHeader footer)
    {
        ArgumentNullException.ThrowIfNull(footer);
        return With(footer: footer);
    }

    /// <summary>
    /// A copy whose first page is treated separately.
    /// </summary>
    /// <remarks>
    /// Calling this is the switch: it emits <c>w:titlePg</c>, and then <b>null means blank on page
    /// one</b> rather than "use the ordinary one". That mirrors the format — an absent first-page
    /// reference produces no header, and there is no inheritance to fall back on — and it makes the
    /// common case sayable: a title page with nothing running across it.
    ///
    /// Not calling this at all leaves page one looking like every other page.
    /// </remarks>
    /// <param name="header">Page one's header, or null for none.</param>
    /// <param name="footer">Page one's footer, or null for none.</param>
    public PageSetup WithFirstPage(DocxHeader? header, DocxHeader? footer)
        => new(
            WidthPoints, HeightPoints, TopPoints, RightPoints, BottomPoints, LeftPoints,
            Header, Footer, header, footer, hasDistinctFirstPage: true);

    /// <summary>
    /// A copy with the paper unchanged and one header slot replaced. Every other <c>With</c> method
    /// must route its result through the constructor's header parameters too — a derived PageSetup
    /// that silently dropped the header would be the exact failure this library refuses.
    /// </summary>
    private PageSetup With(DocxHeader? header = null, DocxHeader? footer = null)
        => new(
            WidthPoints, HeightPoints, TopPoints, RightPoints, BottomPoints, LeftPoints,
            header ?? Header, footer ?? Footer,
            FirstPageHeader, FirstPageFooter, HasDistinctFirstPage);

    /// <inheritdoc/>
    public override string ToString() => string.Format(
        CultureInfo.InvariantCulture,
        "{0} x {1} pt, margins {2}/{3}/{4}/{5} pt",
        WidthPoints, HeightPoints, TopPoints, RightPoints, BottomPoints, LeftPoints);

    // The negated forms below are deliberate: every comparison against NaN is false, so `!(v > 0)`
    // rejects NaN while `v <= 0` silently accepts it.

    private static void RequireDimension(double value, string paramName)
    {
        if (!(value > 0) || value > MaxPoints)
        {
            throw new ArgumentOutOfRangeException(
                paramName, value,
                FormattableString.Invariant(
                    $"Page dimensions must be greater than 0 and no more than {MaxPoints} points."));
        }
    }

    private static void RequireMargin(double value, string paramName)
    {
        if (!(value >= 0) || value > MaxPoints)
        {
            throw new ArgumentOutOfRangeException(
                paramName, value,
                FormattableString.Invariant(
                    $"Margins must be 0 or greater and no more than {MaxPoints} points."));
        }
    }

    private static void RequireContentArea(
        double width, double height,
        double top, double right, double bottom, double left)
    {
        if (left + right >= width)
        {
            throw new ArgumentException(
                FormattableString.Invariant(
                    $"The left and right margins ({left} + {right} pt) leave no content area on a {width} pt wide page."),
                nameof(left));
        }

        if (top + bottom >= height)
        {
            throw new ArgumentException(
                FormattableString.Invariant(
                    $"The top and bottom margins ({top} + {bottom} pt) leave no content area on a {height} pt tall page."),
                nameof(top));
        }
    }
}
