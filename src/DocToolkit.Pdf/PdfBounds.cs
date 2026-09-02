namespace DocToolkit;

/// <summary>
/// Where something sits on a PDF page, in PDF user-space points.
/// </summary>
/// <remarks>
/// <b>The origin is the BOTTOM-left corner of the page, and <see cref="Bottom"/> grows upwards.</b>
/// That is the PDF coordinate system rather than a screen one, and it is the commonest surprise
/// when these values are first plotted — a word near the top of an A4 page has a
/// <see cref="Bottom"/> near 842, not near 0. It is stated here rather than left to be discovered.
///
/// One point is 1/72 inch, so A4 is 595 × 842 and US Letter is 612 × 792.
///
/// <para>
/// This is a value type deliberately. A page of prose produces one of these per word, and a
/// reference type would allocate twice per word for no gain — the whole thing is four doubles.
/// </para>
///
/// <para>
/// It is also this library's own type rather than PdfPig's <c>PdfRectangle</c>, which keeps PdfPig
/// out of the public API. <c>PdfTextExtractor</c> is the only file here that references PdfPig, and
/// returning one of its types from <see cref="PdfEditor"/> would end that boundary — a consumer
/// would then need a PdfPig reference to name the result of a call.
/// </para>
/// </remarks>
public readonly record struct PdfBounds
{
    /// <summary>Creates a rectangle from its lower-left corner and its size.</summary>
    /// <param name="left">Distance from the page's left edge, in points.</param>
    /// <param name="bottom">Distance from the page's BOTTOM edge, in points.</param>
    /// <param name="width">Width in points.</param>
    /// <param name="height">Height in points.</param>
    public PdfBounds(double left, double bottom, double width, double height)
    {
        Left = left;
        Bottom = bottom;
        Width = width;
        Height = height;
    }

    /// <summary>Distance from the page's left edge, in points.</summary>
    public double Left { get; }

    /// <summary>
    /// Distance from the page's <b>bottom</b> edge, in points. Larger means higher up the page.
    /// </summary>
    public double Bottom { get; }

    /// <summary>Width in points.</summary>
    public double Width { get; }

    /// <summary>Height in points.</summary>
    public double Height { get; }

    /// <summary>Distance from the page's left edge to the right-hand side, in points.</summary>
    public double Right => Left + Width;

    /// <summary>Distance from the page's bottom edge to the top side, in points.</summary>
    public double Top => Bottom + Height;
}
