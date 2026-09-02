namespace DocToolkit;

/// <summary>
/// One image drawn on a PDF page: its pixels, its size in pixels, and where it was placed.
/// </summary>
/// <remarks>
/// <b><see cref="Png"/> can be null, and that is not an error.</b> A PDF may store an image in a
/// colour space or filter combination that cannot be re-encoded as PNG without interpreting it.
/// Returning null for that one image, rather than throwing, means a single exotic image does not
/// cost the caller every other image on the page. Check it before use.
///
/// <para>
/// <b><see cref="PixelWidth"/> is not <see cref="PdfBounds.Width"/>.</b> The first is how many
/// pixels the stored image has; the second is how large the page draws it, in points. A 8×8 image
/// placed at 48×48 points is normal — that is scaling, not corruption — so effective DPI is
/// <see cref="PixelWidth"/> ÷ (<see cref="PdfBounds.Width"/> ÷ 72).
/// </para>
///
/// <para>
/// <b>These bytes are a re-encoding, not the file that was embedded.</b> They are the image's
/// pixels rendered to PNG, so they will not be byte-identical to an original PNG or JPEG — the
/// byte count usually differs. Anything needing the original stored stream needs a different
/// operation than this one.
/// </para>
/// </remarks>
public sealed class PdfImage
{
    internal PdfImage(byte[]? png, int pixelWidth, int pixelHeight, PdfBounds bounds)
    {
        Png = png;
        PixelWidth = pixelWidth;
        PixelHeight = pixelHeight;
        Bounds = bounds;
    }

    /// <summary>
    /// The image's pixels encoded as PNG, or <see langword="null"/> when they could not be
    /// re-encoded. See this type's remarks — null is a normal outcome, not a failure.
    /// </summary>
    public byte[]? Png { get; }

    /// <summary>How many pixels wide the stored image is.</summary>
    public int PixelWidth { get; }

    /// <summary>How many pixels tall the stored image is.</summary>
    public int PixelHeight { get; }

    /// <summary>
    /// Where the image was drawn on its page, in PDF user-space points. Unrelated to
    /// <see cref="PixelWidth"/> — see this type's remarks.
    /// </summary>
    public PdfBounds Bounds { get; }

    /// <summary>Pixel size, placed size and whether the pixels decoded, for test failure messages.</summary>
    public override string ToString() =>
        $"{PixelWidth}x{PixelHeight}px at ({Bounds.Left:F1},{Bounds.Bottom:F1}) "
        + $"{Bounds.Width:F1}x{Bounds.Height:F1}pt, {(Png is null ? "no PNG" : Png.Length + " PNG bytes")}";
}
