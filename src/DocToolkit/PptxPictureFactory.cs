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

    /// <summary>
    /// A <c>p:pic</c> shape referencing an image already added to the owning slide part.
    ///
    /// Element order is fixed by the schema — non-visual properties, then the blip fill, then the
    /// shape properties — and PowerPoint reports a repairable file rather than a schema error if it
    /// is wrong, so do not reorder these.
    /// </summary>
    /// <param name="id">Reuse the REPLACED shape's id. A 1:1 swap needs no new id.</param>
    /// <param name="name">Labels the shape in PowerPoint's selection pane.</param>
    /// <param name="relationshipId">From <c>slidePart.GetIdOfPart(imagePart)</c>.</param>
    /// <param name="x">Left offset of the picture, in EMU.</param>
    /// <param name="y">Top offset of the picture, in EMU.</param>
    /// <param name="cx">Width of the picture, in EMU.</param>
    /// <param name="cy">Height of the picture, in EMU.</param>
    public static P.Picture Picture(
        uint id, string name, string relationshipId, long x, long y, long cx, long cy) =>
        new(
            new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualPictureDrawingProperties(
                    new A.PictureLocks { NoChangeAspect = true }),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.BlipFill(
                new A.Blip { Embed = relationshipId },
                new A.Stretch(new A.FillRectangle())),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = x, Y = y },
                    new A.Extents { Cx = cx, Cy = cy }),
                new A.PresetGeometry(new A.AdjustValueList())
                {
                    Preset = A.ShapeTypeValues.Rectangle,
                }));
}
