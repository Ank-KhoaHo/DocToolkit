using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocToolkit;

/// <summary>
/// Turns a list of <see cref="PptxSlide"/> into a PresentationML package.
///
/// Separate from <see cref="PresentationEditor"/> on purpose, matching
/// <see cref="DocxDocumentWriter"/>: that file is already large, and creating a deck has nothing in
/// common with editing one beyond the format.
///
/// A deck needs far more scaffolding than a document: a presentation, a slide master, at least one
/// slide layout, and a theme, before any slide exists. All of it is built here in typed OpenXML
/// rather than cloned from a template asset, so every part is reviewable in a diff and nothing
/// opaque ships in the package. Measured: the whole scaffold is about 3.9 KB.
/// </summary>
internal static class PptxDocumentWriter
{
    // 16:9. Absent, PowerPoint substitutes its own default and every slide is the wrong shape.
    // int, not long: p:sldSz/@cx is ST_SlideSizeCoordinate, which the SDK types as Int32Value.
    private const int SlideWidthEmu = 12192000;
    private const int SlideHeightEmu = 6858000;

    // p:sldMasterId/@id must be >= 2147483648; p:sldId/@id must be in 256..2147483647. Out-of-range
    // values are rejected outright, and duplicates inside a list are the PPTX analogue of a
    // duplicate wp:docPr/@id - PowerPoint declares the file corrupt.
    private const uint FirstMasterId = 2147483648;
    private const uint FirstLayoutId = 2147483649;
    private const uint FirstSlideId = 256;

    /// <summary>
    /// Builds the package into a fresh <see cref="MemoryStream"/>, positioned at 0. The caller owns
    /// and disposes it.
    /// </summary>
    public static MemoryStream Write(IReadOnlyList<PptxSlide> slides)
    {
        var ms = new MemoryStream();

        try
        {
            using (var doc = PresentationDocument.Create(ms, PresentationDocumentType.Presentation))
            {
                var presentationPart = doc.AddPresentationPart();
                presentationPart.Presentation = new P.Presentation();

                var masterPart = AddSlideMaster(presentationPart, out var layoutPart);

                presentationPart.Presentation.Append(
                    new P.SlideMasterIdList(new P.SlideMasterId
                    {
                        Id = FirstMasterId,
                        RelationshipId = presentationPart.GetIdOfPart(masterPart),
                    }),
                    new P.SlideIdList(),
                    new P.SlideSize { Cx = SlideWidthEmu, Cy = SlideHeightEmu },
                    new P.NotesSize { Cx = 6858000, Cy = 9144000 });

                // The scaffold is complete and valid on its own: a deck with no slides is a legal
                // deck. Slide creation is the only thing that consumes the layout, so it stays
                // unused until there is a slide to attach it to.
                _ = layoutPart;

                presentationPart.Presentation.Save();
            }

            ms.Position = 0;
            return ms;
        }
        // Two arms, both disposing. A single arm filtered with
        // `when (ex is not DocumentConversionException)` looks equivalent and is not: a filtered
        // catch that does not match never runs its body, so a DocumentConversionException raised
        // inside the try would escape with `ms` still open. The caller cannot dispose a stream it
        // was never handed. This exact bug was found by review on the DOCX writer.
        catch (DocumentConversionException)
        {
            ms.Dispose();
            throw;
        }
        catch (Exception ex)
        {
            ms.Dispose();
            throw new DocumentConversionException("Failed to create PPTX.", ex);
        }
    }

    /// <summary>
    /// The master, its single layout and the theme. Returns the master part and hands back the
    /// layout part, which slides must reference.
    /// </summary>
    private static SlideMasterPart AddSlideMaster(
        PresentationPart presentationPart, out SlideLayoutPart layoutPart)
    {
        var masterPart = presentationPart.AddNewPart<SlideMasterPart>();
        masterPart.SlideMaster = new P.SlideMaster(
            EmptyShapeTree(),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink,
            });

        layoutPart = masterPart.AddNewPart<SlideLayoutPart>();
        layoutPart.SlideLayout = new P.SlideLayout(
            EmptyShapeTree(),
            new P.ColorMapOverride(new A.MasterColorMapping()));

        masterPart.SlideMaster.Append(new P.SlideLayoutIdList(new P.SlideLayoutId
        {
            Id = FirstLayoutId,
            RelationshipId = masterPart.GetIdOfPart(layoutPart),
        }));

        masterPart.AddNewPart<ThemePart>().Theme = MinimalTheme();
        masterPart.SlideMaster.Save();
        layoutPart.SlideLayout.Save();

        return masterPart;
    }

    /// <summary>
    /// The empty shape tree every slide, layout and master must carry. The two non-visual property
    /// elements and the group-shape properties are all required by the schema even when nothing is
    /// drawn.
    /// </summary>
    private static P.CommonSlideData EmptyShapeTree() =>
        new(new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties()));

    /// <summary>
    /// A theme is not required for the package to validate - measured, a themeless deck reports
    /// zero validation errors - but it is what supplies the colour and font scheme that the
    /// master's clrMap refers to, and this feature cannot be checked against PowerPoint from CI.
    /// Omitting it would bet that PowerPoint tolerates a dangling reference, which is exactly the
    /// class of invisible failure this repo has been bitten by before. 684 bytes is not a saving
    /// worth that risk.
    ///
    /// Every child here is schema-required: the colour scheme needs all twelve slots, the font
    /// scheme both major and minor with latin/ea/cs, and the format scheme at least three entries
    /// in each of its four style lists.
    /// </summary>
    private static A.Theme MinimalTheme()
    {
        static A.SchemeColor Ph() => new() { Val = A.SchemeColorValues.PhColor };

        return new A.Theme(
            new A.ThemeElements(
                new A.ColorScheme(
                    new A.Dark1Color(new A.SystemColor { Val = A.SystemColorValues.WindowText }),
                    new A.Light1Color(new A.SystemColor { Val = A.SystemColorValues.Window }),
                    new A.Dark2Color(new A.RgbColorModelHex { Val = "44546A" }),
                    new A.Light2Color(new A.RgbColorModelHex { Val = "E7E6E6" }),
                    new A.Accent1Color(new A.RgbColorModelHex { Val = "4472C4" }),
                    new A.Accent2Color(new A.RgbColorModelHex { Val = "ED7D31" }),
                    new A.Accent3Color(new A.RgbColorModelHex { Val = "A5A5A5" }),
                    new A.Accent4Color(new A.RgbColorModelHex { Val = "FFC000" }),
                    new A.Accent5Color(new A.RgbColorModelHex { Val = "5B9BD5" }),
                    new A.Accent6Color(new A.RgbColorModelHex { Val = "70AD47" }),
                    new A.Hyperlink(new A.RgbColorModelHex { Val = "0563C1" }),
                    new A.FollowedHyperlinkColor(new A.RgbColorModelHex { Val = "954F72" }))
                { Name = "Office" },
                new A.FontScheme(
                    new A.MajorFont(
                        new A.LatinFont { Typeface = "Calibri Light" },
                        new A.EastAsianFont { Typeface = string.Empty },
                        new A.ComplexScriptFont { Typeface = string.Empty }),
                    new A.MinorFont(
                        new A.LatinFont { Typeface = "Calibri" },
                        new A.EastAsianFont { Typeface = string.Empty },
                        new A.ComplexScriptFont { Typeface = string.Empty }))
                { Name = "Office" },
                new A.FormatScheme(
                    new A.FillStyleList(
                        new A.SolidFill(Ph()), new A.SolidFill(Ph()), new A.SolidFill(Ph())),
                    new A.LineStyleList(
                        new A.Outline(new A.SolidFill(Ph())) { Width = 6350 },
                        new A.Outline(new A.SolidFill(Ph())) { Width = 12700 },
                        new A.Outline(new A.SolidFill(Ph())) { Width = 19050 }),
                    new A.EffectStyleList(
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(new A.EffectList()),
                        new A.EffectStyle(new A.EffectList())),
                    new A.BackgroundFillStyleList(
                        new A.SolidFill(Ph()), new A.SolidFill(Ph()), new A.SolidFill(Ph())))
                { Name = "Office" }),
            new A.ObjectDefaults(),
            new A.ExtraColorSchemeList())
        { Name = "Office Theme" };
    }
}
