using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocToolkit;

/// <summary>
/// Builds the picture shape that replaces a placeholder box, and works out where it goes.
/// </summary>
internal static class PptxPictureFactory
{
    /// <summary>
    /// Scales an image to fit entirely inside a box, preserving its aspect ratio, and centres it.
    ///
    /// Scaling applies in BOTH directions — an image smaller than its box is enlarged. That is
    /// deliberate: a rule that sometimes fills the box and sometimes does not is surprising, and a
    /// caller supplying a tiny logo for a large box has a source problem that a silently
    /// half-filled box would hide.
    ///
    /// All values are EMU. Integer division on the centring is intentional; a rounding error of
    /// half an EMU is 1/914400 of an inch.
    /// </summary>
    public static (long X, long Y, long Cx, long Cy) Fit(
        long boxX, long boxY, long boxCx, long boxCy, long imageCx, long imageCy)
    {
        ArgumentOutOfRangeException.ThrowIfLessThanOrEqual(boxCx, 0);
        ArgumentOutOfRangeException.ThrowIfLessThanOrEqual(boxCy, 0);
        ArgumentOutOfRangeException.ThrowIfLessThanOrEqual(imageCx, 0);
        ArgumentOutOfRangeException.ThrowIfLessThanOrEqual(imageCy, 0);

        var scale = Math.Min((double)boxCx / imageCx, (double)boxCy / imageCy);

        var cx = (long)Math.Round(imageCx * scale);
        var cy = (long)Math.Round(imageCy * scale);

        return (boxX + (boxCx - cx) / 2, boxY + (boxCy - cy) / 2, cx, cy);
    }
}
