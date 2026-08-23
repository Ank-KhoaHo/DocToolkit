using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocToolkit;

/// <summary>
/// Turns a list of <see cref="PptxSlide"/> into a PresentationML package.
///
/// Separate from <see cref="PresentationEditor"/> on purpose, matching
/// <c>DocxDocumentWriter</c>: that file is already large, and creating a deck has nothing in
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

    // p:notesSz is required by the schema. US Letter portrait, which is what PowerPoint itself
    // writes for a 16:9 deck - the notes page is a printed page and does not track the slide's
    // aspect ratio, so this deliberately is not SlideWidthEmu/SlideHeightEmu.
    private const int NotesWidthEmu = 6858000;
    private const int NotesHeightEmu = 9144000;

    // p:sldMasterId/@id must be >= 2147483648; p:sldId/@id must be in 256..2147483647. Out-of-range
    // values are rejected outright, and duplicates inside a list are the PPTX analogue of a
    // duplicate wp:docPr/@id - PowerPoint declares the file corrupt.
    private const uint FirstMasterId = 2147483648;
    private const uint FirstLayoutId = 2147483649;
    private const uint FirstSlideId = 256;

    // The title and body boxes, shared by the layout's placeholders and every slide's. They must
    // agree: a slide placeholder inherits geometry from the layout placeholder with the same
    // type/idx, so differing boxes mean "Reset Slide" visibly moves the shape.
    private const long TitleXEmu = 838200;
    private const long TitleYEmu = 365125;
    private const long TitleWidthEmu = 10515600;
    private const long TitleHeightEmu = 1325563;
    private const long BodyXEmu = 838200;
    private const long BodyYEmu = 1825625;
    private const long BodyWidthEmu = 10515600;
    private const long BodyHeightEmu = 4351338;

    // p:ph/@idx on the body placeholder. Type alone cannot distinguish two body placeholders on one
    // layout; idx is what binds a slide's shape to the layout's. The title placeholder carries no
    // idx - a layout has at most one, and ECMA-376 treats an absent idx as 0.
    private const uint BodyPlaceholderIndex = 1;

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
                    new P.NotesSize { Cx = NotesWidthEmu, Cy = NotesHeightEmu });

                var slideIdList = presentationPart.Presentation.GetFirstChild<P.SlideIdList>()!;
                var nextSlideId = FirstSlideId;

                foreach (var slide in slides)
                {
                    var slidePart = presentationPart.AddNewPart<SlidePart>();
                    slidePart.Slide = BuildSlide(slide);
                    slidePart.AddPart(layoutPart);
                    slidePart.Slide.Save();

                    slideIdList.Append(new P.SlideId
                    {
                        Id = nextSlideId,
                        RelationshipId = presentationPart.GetIdOfPart(slidePart),
                    });

                    nextSlideId++;
                }

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
            throw new DocumentConversionException("Failed to create PPTX. See the inner exception for details.", ex);
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
            LayoutShapeTree(),
            new P.ColorMapOverride(new A.MasterColorMapping()))
        {
            // Without these PowerPoint's layout gallery shows the part's filename - "SlideLayout2".
            // "obj" is the type PowerPoint uses for its own built-in Title and Content, which is
            // what this writer builds.
            Type = P.SlideLayoutValues.Object,
        };

        // Master->layout and layout->master are TWO separate relationships, not one expressed from
        // both ends. AddNewPart<SlideLayoutPart>() above created only the first; p:sldLayoutIdLst
        // below records that same direction again, as data rather than as a package relationship.
        // Without this line the package has no ppt/slideLayouts/_rels/ directory at all. That is
        // schema-valid - OpenXmlValidator checks part XML against the schema, not the relationship
        // graph - so every test was green while PowerPoint's Presentations.Open failed with E_FAIL.
        // Verified by controlled experiment against PowerPoint 16.0.
        layoutPart.AddPart(masterPart);

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
    /// The layout's shape tree: the title and body placeholders every slide inherits from. They
    /// hold no text — a layout placeholder defines the box and the role, and PowerPoint renders the
    /// slide's own copy. Their geometry must match the slide's, because "Reset Slide" restores the
    /// slide shape to the layout's box.
    ///
    /// The <c>name</c> on <see cref="P.CommonSlideData"/> is what PowerPoint's layout gallery
    /// labels this layout; without it the label is the part filename, "SlideLayout2".
    ///
    /// Measured against PowerPoint 16.0, the same deck built without any of this: every slide
    /// reported <c>Shapes.HasTitle = False</c> and <c>CustomLayout.Name = "SlideLayout2"</c>, so
    /// Outline View listed the slides with no titles at all. With it: <c>HasTitle = True</c>,
    /// <c>Shapes.Title.TextFrame.TextRange.Text</c> returns the title, layout name
    /// "Title and Content". Every schema validator passed both ways — a p:ph is a role, and roles
    /// are not something a schema can check.
    /// </summary>
    private static P.CommonSlideData LayoutShapeTree()
    {
        var tree = new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties());

        tree.Append(PlaceholderShape(
            id: 2U, name: "Title Placeholder 1",
            xEmu: TitleXEmu, yEmu: TitleYEmu, widthEmu: TitleWidthEmu, heightEmu: TitleHeightEmu,
            placeholder: new P.PlaceholderShape { Type = P.PlaceholderValues.Title }));

        tree.Append(PlaceholderShape(
            id: 3U, name: "Content Placeholder 2",
            xEmu: BodyXEmu, yEmu: BodyYEmu, widthEmu: BodyWidthEmu, heightEmu: BodyHeightEmu,
            placeholder: new P.PlaceholderShape
            {
                Type = P.PlaceholderValues.Body,
                Index = BodyPlaceholderIndex,
            }));

        return new P.CommonSlideData(tree) { Name = "Title and Content" };
    }

    /// <summary>
    /// An empty placeholder shape for the layout. A text body is schema-required on
    /// <see cref="P.Shape"/> even when the shape carries no text.
    /// </summary>
    private static P.Shape PlaceholderShape(
        uint id, string name, long xEmu, long yEmu, long widthEmu, long heightEmu,
        P.PlaceholderShape placeholder) =>
        new(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties(placeholder)),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = xEmu, Y = yEmu },
                    new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            new P.TextBody(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph()));

    /// <summary>
    /// A theme is not required for the package to validate - measured, a themeless deck reports
    /// zero validation errors - but it is what supplies the colour and font scheme that the
    /// master's clrMap refers to, and this feature cannot be checked against PowerPoint from CI.
    /// Omitting it would bet that PowerPoint tolerates a dangling reference, which is exactly the
    /// class of invisible failure this repo has been bitten by before. 684 bytes is not a saving
    /// worth that risk.
    ///
    /// Every child here is schema-required: the colour scheme needs all twelve slots, the font
    /// scheme both major and minor with latin/ea/cs, and the format scheme exactly three entries
    /// in each of its four style lists — ECMA-376 fixes the count, it is not merely a minimum.
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

    /// <summary>
    /// One slide: a title shape and, when there are bullets, a body shape beneath it.
    ///
    /// Both are placed with explicit EMU boxes because a slide has no flow layout - every shape
    /// carries its own position and size, unlike a paragraph in a document.
    /// </summary>
    private static P.Slide BuildSlide(PptxSlide slide)
    {
        var tree = new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties());

        tree.Append(TextShape(
            id: 2U, name: "Title",
            xEmu: TitleXEmu, yEmu: TitleYEmu, widthEmu: TitleWidthEmu, heightEmu: TitleHeightEmu,
            fontSizeHundredths: 4400,
            lines: new[] { slide.Title },
            bulleted: false,
            placeholder: new P.PlaceholderShape { Type = P.PlaceholderValues.Title }));

        if (slide.Bullets.Count > 0)
        {
            tree.Append(TextShape(
                id: 3U, name: "Content",
                xEmu: BodyXEmu, yEmu: BodyYEmu, widthEmu: BodyWidthEmu, heightEmu: BodyHeightEmu,
                fontSizeHundredths: 2000,
                lines: slide.Bullets,
                bulleted: true,
                placeholder: new P.PlaceholderShape
                {
                    Type = P.PlaceholderValues.Body,
                    Index = BodyPlaceholderIndex,
                }));
        }

        return new P.Slide(new P.CommonSlideData(tree), new P.ColorMapOverride(new A.MasterColorMapping()));
    }

    /// <summary>
    /// A text box holding one paragraph per line. <paramref name="id"/> must be unique within the
    /// slide's shape tree; 1 is taken by the group shape itself.
    ///
    /// <paramref name="bulleted"/> controls whether each paragraph gets an explicit bullet
    /// character. This is not cosmetic, and it stays explicit even though the shape is now a
    /// placeholder: bullet formatting would be inherited from the master's <c>p:txStyles</c>, which
    /// this minimal master does not define, so there is still nothing in the chain to fall back on.
    /// Without an explicit <see cref="A.CharacterBullet"/>, PowerPoint reports Bullet.Visible = 0
    /// and renders a plain line - schema-valid, and silently not what "bullet" promised.
    ///
    /// <paramref name="placeholder"/> is what makes the shape a title or body rather than a loose
    /// text box. Every visual property here is still set explicitly and overrides what the
    /// placeholder would inherit, so adding it changes roles and not appearance.
    /// </summary>
    private static P.Shape TextShape(
        uint id, string name, long xEmu, long yEmu, long widthEmu, long heightEmu,
        int fontSizeHundredths, IReadOnlyList<string> lines, bool bulleted,
        P.PlaceholderShape placeholder)
    {
        var body = new P.TextBody(new A.BodyProperties(), new A.ListStyle());

        foreach (var line in lines)
        {
            var paragraph = new A.Paragraph();

            if (bulleted)
            {
                paragraph.Append(new A.ParagraphProperties(new A.CharacterBullet { Char = "•" })
                {
                    LeftMargin = 285750,
                    Indent = -285750,
                });
            }

            paragraph.Append(new A.Run(
                new A.RunProperties { Language = "en-US", FontSize = fontSizeHundredths },
                new A.Text(line)));

            body.Append(paragraph);
        }

        return new P.Shape(
            new P.NonVisualShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = name },
                new P.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new P.ApplicationNonVisualDrawingProperties(placeholder)),
            new P.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = xEmu, Y = yEmu },
                    new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
            body);
    }
}
