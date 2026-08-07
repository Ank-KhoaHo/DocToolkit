using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocToolkit;

/// <summary>
/// Builds the DrawingML for an inline image.
///
/// Verbose, and deliberately so: this is the markup Word itself emits. The legacy VML alternative
/// (<c>w:pict</c>) is a fraction of the size, and this repo even has VML fixtures already for the
/// text-box tests — but VML is deprecated, and an image written that way looks subtly unlike one a
/// human inserted.
/// </summary>
internal static class DrawingFactory
{
    /// <param name="relationshipId">
    /// The image part's relationship id, resolved in the part that OWNS the paragraph. A header's
    /// image referenced by a main-document relationship id resolves in the wrong scope: Word opens
    /// the file and shows nothing.
    /// </param>
    /// <param name="name">Labels the object in Word's selection pane. Not the alt text.</param>
    /// <param name="description">
    /// The alt text — <c>wp:docPr/@descr</c>, what a screen reader announces. Null OMITS the
    /// attribute, which is deliberate: this used to be set to <paramref name="name"/>, so every
    /// image shipped alt text reading "Image 1". That is worse than none, because it is announced
    /// as though it described the picture rather than signalling that nothing describes it.
    /// </param>
    /// <param name="id">
    /// Must be unique across the whole document. A duplicate makes Word declare the file corrupt
    /// and offer to repair it.
    /// </param>
    /// <param name="widthEmu">Rendered width in EMUs. 1 point = 12,700; 1 pixel at 96 DPI = 9,525.</param>
    /// <param name="heightEmu">Rendered height in EMUs, in the same units as <paramref name="widthEmu"/>.</param>
    public static Drawing InlineImage(
        string relationshipId, string name, uint id, long widthEmu, long heightEmu,
        string? description = null) =>
        new(new DW.Inline(
            new DW.Extent { Cx = widthEmu, Cy = heightEmu },
            new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
            // Description is assigned only when there is one. Assigning null does NOT omit the
            // attribute - StringValue's implicit conversion yields an instance wrapping null, which
            // serialises as descr="", and an empty alt text is a different statement from an absent
            // one. Measured, after a test asserting absence found descr="" instead.
            description is null
                ? new DW.DocProperties { Id = id, Name = name }
                : new DW.DocProperties { Id = id, Name = name, Description = description },
            new DW.NonVisualGraphicFrameDrawingProperties(
                new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(
                new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = 0U, Name = name },
                            new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(
                            new A.Blip { Embed = relationshipId },
                            new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0L, Y = 0L },
                                new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                            new A.PresetGeometry(new A.AdjustValueList())
                            {
                                Preset = A.ShapeTypeValues.Rectangle,
                            })))
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture",
                }))
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U,
        });
}
