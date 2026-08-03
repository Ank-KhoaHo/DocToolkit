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
    /// <param name="name">Shown in Word's selection and accessibility panes.</param>
    /// <param name="id">
    /// Must be unique across the whole document. A duplicate makes Word declare the file corrupt
    /// and offer to repair it.
    /// </param>
    /// <param name="widthEmu">Rendered width in EMUs. 1 point = 12,700; 1 pixel at 96 DPI = 9,525.</param>
    /// <param name="heightEmu">Rendered height in EMUs, in the same units as <paramref name="widthEmu"/>.</param>
    public static Drawing InlineImage(
        string relationshipId, string name, uint id, long widthEmu, long heightEmu) =>
        new(new DW.Inline(
            new DW.Extent { Cx = widthEmu, Cy = heightEmu },
            new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
            new DW.DocProperties { Id = id, Name = name, Description = name },
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
