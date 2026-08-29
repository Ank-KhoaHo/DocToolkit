using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocToolkit.Tests;

/// <summary>
/// Which half of an <c>a:xfrm</c> a fixture should omit — see
/// <see cref="PptxFixtures.SampleAttachedToLayoutWithAnIncompleteTitleBox"/>. Top-level rather
/// than nested in <see cref="PptxFixtures"/>: that class is <c>internal</c>, and a public
/// <c>[Theory]</c> method's parameter type must be at least as accessible as the method itself,
/// which a type nested in an internal class cannot be however it is itself modified.
/// <c>CT_Transform2D</c> declares both halves optional, so a real layout can be missing either
/// one independently.
/// </summary>
public enum XfrmPart { Offset, Extents }

/// <summary>
/// Fixtures derived from the committed one-slide sample deck.
///
/// Building a valid .pptx from nothing needs presentation, master, layout and theme parts, so
/// these clone the real fixture and mutate it instead: extra slides come from cloning the sample
/// slide part, and the deck order is then set independently of the part order so the two can be
/// told apart.
///
/// Several fixtures here (<see cref="DeckWithPlaceholderBox"/>,
/// <see cref="DeckWithUnpositionedPlaceholder"/>) locate the sample slide's text with
/// <c>.Single()</c> over its shapes and <c>.First()</c> over its text runs. That only works
/// because <c>sample.pptx</c> holds exactly one shape and one text run — an invariant nothing
/// enforces except this comment, so a change to that asset must preserve it.
/// </summary>
internal static class PptxFixtures
{
    private static readonly string SampleAssetPath =
        Path.Join(AppContext.BaseDirectory, "assets", "sample.pptx");

    public static byte[] Sample() => File.ReadAllBytes(SampleAssetPath);

    /// <summary>
    /// A deck whose slide *parts* are created in the order of <paramref name="partOrderTexts"/>.
    /// When <paramref name="reverseDeckOrder"/> is set the p:sldIdLst is then reversed, so the
    /// order PowerPoint shows is the exact opposite of the part-relationship order — which is what
    /// separates "walk SlideParts" from "walk SlideIdList".
    /// </summary>
    public static byte[] MultiSlideDeck(IReadOnlyList<string> partOrderTexts, bool reverseDeckOrder)
    {
        using var ms = Load(Sample());

        using (var doc = PresentationDocument.Open(ms, true))
        {
            var presentationPart = doc.PresentationPart!;
            var slideIdList = presentationPart.Presentation!.SlideIdList!;
            var template = presentationPart.SlideParts.Single();

            SetSoleText(template, partOrderTexts[0]);

            var nextId = slideIdList.Elements<P.SlideId>().Max(s => s.Id!.Value) + 1;
            foreach (var text in partOrderTexts.Skip(1))
            {
                var clone = presentationPart.AddNewPart<SlidePart>();
                clone.Slide = (P.Slide)template.Slide!.CloneNode(true);
                clone.AddPart(template.SlideLayoutPart!);
                SetSoleText(clone, text);

                slideIdList.Append(new P.SlideId
                {
                    Id = nextId++,
                    RelationshipId = presentationPart.GetIdOfPart(clone),
                });
            }

            if (reverseDeckOrder)
            {
                var ids = slideIdList.Elements<P.SlideId>().ToList();
                foreach (var id in ids) id.Remove();
                foreach (var id in Enumerable.Reverse(ids)) slideIdList.Append(id);
            }

            presentationPart.Presentation.Save();
        }

        return ms.ToArray();
    }

    /// <summary>
    /// A two-slide deck where each slide has its OWN, DISTINCT layout — unlike every other
    /// fixture in this file, which clones <c>sample.pptx</c>'s single layout onto every slide.
    /// Needed to prove <c>InsertSlides</c> picks the layout of the slide ADJACENT to the
    /// insertion point rather than always the deck's first layout: a deck built with only one
    /// layout could never tell the two apart.
    ///
    /// The second layout is a clone of the first (so it is schema-shaped correctly) added to the
    /// same slide master, named <c>"Second Layout"</c> so a test can identify it by name.
    /// </summary>
    public static byte[] MultiLayoutDeck(string firstSlideText, string secondSlideText)
    {
        using var ms = Load(Sample());

        using (var doc = PresentationDocument.Open(ms, true))
        {
            var presentationPart = doc.PresentationPart!;
            var slideIdList = presentationPart.Presentation!.SlideIdList!;
            var firstSlidePart = presentationPart.SlideParts.Single();
            var masterPart = firstSlidePart.SlideLayoutPart!.SlideMasterPart!;

            SetSoleText(firstSlidePart, firstSlideText);

            var secondLayoutPart = masterPart.AddNewPart<SlideLayoutPart>();
            secondLayoutPart.SlideLayout =
                (P.SlideLayout)firstSlidePart.SlideLayoutPart!.SlideLayout!.CloneNode(true);
            secondLayoutPart.SlideLayout.CommonSlideData!.Name = "Second Layout";
            secondLayoutPart.AddPart(masterPart);
            secondLayoutPart.SlideLayout.Save();

            var layoutIdList = masterPart.SlideMaster!.SlideLayoutIdList!;
            var nextLayoutId = layoutIdList.Elements<P.SlideLayoutId>().Max(l => l.Id!.Value) + 1;
            layoutIdList.Append(new P.SlideLayoutId
            {
                Id = nextLayoutId,
                RelationshipId = masterPart.GetIdOfPart(secondLayoutPart),
            });
            masterPart.SlideMaster.Save();

            var secondSlidePart = presentationPart.AddNewPart<SlidePart>();
            secondSlidePart.Slide = (P.Slide)firstSlidePart.Slide!.CloneNode(true);
            secondSlidePart.AddPart(secondLayoutPart);
            SetSoleText(secondSlidePart, secondSlideText);

            var nextSlideId = slideIdList.Elements<P.SlideId>().Max(s => s.Id!.Value) + 1;
            slideIdList.Append(new P.SlideId
            {
                Id = nextSlideId,
                RelationshipId = presentationPart.GetIdOfPart(secondSlidePart),
            });

            presentationPart.Presentation.Save();
        }

        return ms.ToArray();
    }

    /// <summary>
    /// A single-slide deck built via <see cref="PresentationEditor.Create"/> — so its layout's
    /// placeholders start out with EXACTLY the types <c>PptxDocumentWriter.BuildSlide</c> writes
    /// (<c>title</c>, <c>body</c> idx <c>1</c>) — with those SAME placeholders' geometry then
    /// relocated far from <c>PptxDocumentWriter</c>'s own constants, types left untouched. This is
    /// the one fixture that can prove geometry inheritance actually tracks the layout: a fixture
    /// whose layout geometry happens to already match <c>PptxDocumentWriter</c>'s constants could
    /// never tell "inherited" apart from "coincidentally identical explicit box".
    /// </summary>
    public static byte[] DeckWithRelocatedLayoutPlaceholders()
    {
        var deck = DocToolkit.PresentationEditor.Create(new[] { DocToolkit.PptxSlide.Titled("First", "One") });

        using var ms = Load(deck);
        using (var doc = PresentationDocument.Open(ms, true))
        {
            var layoutPart = doc.PresentationPart!.SlideParts.Single().SlideLayoutPart!;
            foreach (var shape in layoutPart.SlideLayout!.Descendants<P.Shape>())
            {
                var ph = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                    ?.GetFirstChild<P.PlaceholderShape>();
                var xfrm = shape.ShapeProperties?.Transform2D;
                if (xfrm?.Offset is null || xfrm.Extents is null) continue;

                if (ph?.Type?.Value == P.PlaceholderValues.Title)
                {
                    xfrm.Offset.X = 500000; xfrm.Offset.Y = 5000000; // near the bottom
                    xfrm.Extents.Cx = 8000000; xfrm.Extents.Cy = 1000000;
                }
                else if (ph?.Type?.Value == P.PlaceholderValues.Body && (ph?.Index?.Value ?? 0) == 1)
                {
                    xfrm.Offset.X = 500000; xfrm.Offset.Y = 500000; // near the top
                    xfrm.Extents.Cx = 8000000; xfrm.Extents.Cy = 2000000;
                }
            }
            layoutPart.SlideLayout.Save();
        }
        return ms.ToArray();
    }

    /// <summary>
    /// <c>sample.pptx</c> with its one slide re-pointed at another of the ELEVEN real,
    /// PowerPoint-authored layouts its package already ships — <c>slideLayout1.xml</c> ("Title
    /// Slide") is the only one wired to a slide as committed, so the other ten are unreachable
    /// without this.
    ///
    /// <b>Why this fixture has to exist at all.</b> Every other layout fixture here is built by
    /// <see cref="DocToolkit.PresentationEditor.Create"/>, whose layout always carries explicit
    /// placeholder geometry because the writer puts it there. Real PowerPoint layouts routinely do
    /// not: they declare a placeholder's ROLE and inherit its position from the slide master.
    /// Measured across all eleven here, four do exactly that — "Title and Content" (the default for
    /// a new body slide), "Two Content", "Title Only" and "Title and Vertical Text". A fixture that
    /// can only produce positioned layouts cannot see the case those four represent, which is the
    /// "what you measure is the fixture" trap <c>CLAUDE.md</c> records for <c>DocxForm</c>.
    ///
    /// <b>The mechanism, measured rather than assumed.</b> A <see cref="SlidePart"/> allows only one
    /// <see cref="SlideLayoutPart"/> relationship, so the old one is deleted before the new one is
    /// added. <c>DeletePart</c> drops the RELATIONSHIP, not the part: the slide master still
    /// references every layout, so the package still holds all eleven afterwards (verified — and the
    /// result validates clean under <c>OpenXmlValidator</c>, both before and after an insertion).
    ///
    /// Layouts are found by their <c>p:cSld/@name</c> rather than by part URI, because the name is
    /// what makes a test read as the case it is testing. An unknown name throws rather than
    /// silently selecting a different layout.
    /// </summary>
    public static byte[] SampleAttachedToLayout(string layoutName)
    {
        using var ms = Load(Sample());

        using (var doc = PresentationDocument.Open(ms, true))
        {
            var presentationPart = doc.PresentationPart!;
            var slidePart = presentationPart.SlideParts.Single();

            var layouts = presentationPart.SlideMasterParts.SelectMany(m => m.SlideLayoutParts).ToList();
            var target = layouts.FirstOrDefault(
                l => l.SlideLayout?.CommonSlideData?.Name?.Value == layoutName)
                ?? throw new InvalidOperationException(
                    $"sample.pptx ships no layout named '{layoutName}'. Available: " +
                    string.Join(", ", layouts.Select(
                        l => l.SlideLayout?.CommonSlideData?.Name?.Value ?? "<unnamed>")));

            slidePart.DeletePart(slidePart.SlideLayoutPart!);
            slidePart.AddPart(target);

            slidePart.Slide!.Save();
            presentationPart.Presentation!.Save();
        }

        return ms.ToArray();
    }

    /// <summary>
    /// <see cref="SampleAttachedToLayout"/> for <paramref name="layoutName"/>, with that layout's
    /// title placeholder's own <c>a:xfrm</c> then stripped down to JUST one half —
    /// <paramref name="missingPart"/> selects which — reproducing an <see cref="A.Transform2D"/>
    /// that is present but not a usable box. <c>CT_Transform2D</c> declares both <c>a:off</c> and
    /// <c>a:ext</c> as optional, so either shape is schema-valid on a real PowerPoint-authored
    /// layout, not a fixture-only shape: a layout author can genuinely produce either one.
    ///
    /// <paramref name="missingPart"/> exists so both halves of
    /// <c>LayoutHasMatchingPositionedPlaceholder</c>'s completeness check
    /// (<c>layoutXfrm?.Offset is null || layoutXfrm.Extents is null</c>) are exercised by a real
    /// test. Before this parameter, only the <c>Extents</c>-missing case had one — nothing in the
    /// suite would have failed if the <c>Offset is null</c> half of that condition were deleted.
    /// </summary>
    public static byte[] SampleAttachedToLayoutWithAnIncompleteTitleBox(
        string layoutName, XfrmPart missingPart = XfrmPart.Extents)
    {
        using var ms = Load(SampleAttachedToLayout(layoutName));

        using (var doc = PresentationDocument.Open(ms, true))
        {
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            var layoutPart = slidePart.SlideLayoutPart!;

            var titleShape = layoutPart.SlideLayout!.Descendants<P.Shape>().First(shape =>
                shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                    ?.GetFirstChild<P.PlaceholderShape>()?.Type?.Value == P.PlaceholderValues.Title);

            var xfrm = titleShape.ShapeProperties!.Transform2D!;
            if (missingPart == XfrmPart.Extents) xfrm.Extents!.Remove();
            else xfrm.Offset!.Remove();

            layoutPart.SlideLayout.Save();
        }

        return ms.ToArray();
    }

    /// <summary>
    /// <see cref="SampleAttachedToLayout"/> for <paramref name="layoutName"/>, with that layout's
    /// title placeholder moved INSIDE a newly appended <c>p:grpSp</c> — matching role (type
    /// <c>title</c>) and still carrying its own complete <c>a:xfrm</c> (both <c>a:off</c> and
    /// <c>a:ext</c>, unchanged), but no longer one of the layout's TOP-LEVEL shapes.
    ///
    /// Reproduces the case <c>LayoutHasMatchingPositionedPlaceholder</c>'s top-level-only walk
    /// exists to refuse: <c>Descendants&lt;P.Shape&gt;()</c> also matches a shape nested inside a
    /// group, but this repo's render pipeline (<c>PptxToPdfConverter</c>/OfficeIMO) resolves a
    /// slide's inherited geometry from a layout's TOP-LEVEL shape tree only. A grouped placeholder
    /// is therefore a role-and-geometry match this library must NOT inherit from — schema-valid
    /// both before and after the move (verified via <c>OpenXmlValidator</c>), the same way
    /// <see cref="SampleWithPlaceholderInGroup"/> reproduces the analogous case for
    /// <c>ReplaceImage</c>. The body placeholder is left untouched (still top-level, still a
    /// complete box), so a test using this fixture can tell "this placeholder correctly does not
    /// match" apart from "nothing on this layout matches anything".
    /// </summary>
    public static byte[] SampleAttachedToLayoutWithTitleInGroup(string layoutName)
    {
        using var ms = Load(SampleAttachedToLayout(layoutName));

        using (var doc = PresentationDocument.Open(ms, true))
        {
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            var layoutPart = slidePart.SlideLayoutPart!;
            var shapeTree = layoutPart.SlideLayout!.CommonSlideData!.ShapeTree!;

            var titleShape = shapeTree.Elements<P.Shape>().First(shape =>
                shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
                    ?.GetFirstChild<P.PlaceholderShape>()?.Type?.Value == P.PlaceholderValues.Title);

            var titleXfrm = titleShape.ShapeProperties!.Transform2D!;
            var x = titleXfrm.Offset!.X!.Value;
            var y = titleXfrm.Offset.Y!.Value;
            var cx = titleXfrm.Extents!.Cx!.Value;
            var cy = titleXfrm.Extents.Cy!.Value;

            var nextId = shapeTree.Descendants<P.NonVisualDrawingProperties>()
                .Select(p => p.Id?.Value ?? 0U).DefaultIfEmpty(0U).Max() + 1;

            titleShape.Remove(); // detach -- still carries its own complete a:xfrm, type and index

            shapeTree.AppendChild(new P.GroupShape(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties { Id = nextId, Name = "Grouped Title" },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(
                    new A.TransformGroup(
                        new A.Offset { X = x, Y = y },
                        new A.Extents { Cx = cx, Cy = cy },
                        new A.ChildOffset { X = x, Y = y },
                        new A.ChildExtents { Cx = cx, Cy = cy })),
                titleShape));

            layoutPart.SlideLayout.Save();
        }

        return ms.ToArray();
    }

    /// <summary>
    /// <c>sample.pptx</c>'s slide re-pointed at a newly created <see cref="SlideLayoutPart"/> whose
    /// root <see cref="P.SlideLayout"/> element is deliberately never assigned — the case
    /// <c>LayoutHasMatchingPositionedPlaceholder</c>'s null-<c>SlideLayout</c> guard exists to
    /// survive. Before this branch, <c>InsertSlides</c> never read a layout's XML content at all,
    /// so a <see cref="SlideLayoutPart"/> with no root element was not a failure mode it could
    /// reach; reading the layout's placeholders to decide on geometry inheritance made it reachable.
    ///
    /// Registered in the master's <c>SlideLayoutIdList</c> the same way <see cref="MultiLayoutDeck"/>
    /// registers its second layout, so the deck stays shaped like one a real package would produce
    /// rather than merely one <c>ResolveLayoutForInsertion</c> happens to tolerate.
    /// </summary>
    public static byte[] SampleAttachedToAnUnassignedLayout()
    {
        using var ms = Load(Sample());

        using (var doc = PresentationDocument.Open(ms, true))
        {
            var presentationPart = doc.PresentationPart!;
            var slidePart = presentationPart.SlideParts.Single();
            var masterPart = slidePart.SlideLayoutPart!.SlideMasterPart!;

            var emptyLayoutPart = masterPart.AddNewPart<SlideLayoutPart>();
            emptyLayoutPart.AddPart(masterPart);
            // SlideLayout is deliberately never assigned: the part exists with no root element,
            // so its SlideLayout getter returns null rather than throwing.

            var layoutIdList = masterPart.SlideMaster!.SlideLayoutIdList!;
            var nextLayoutId = layoutIdList.Elements<P.SlideLayoutId>().Max(l => l.Id!.Value) + 1;
            layoutIdList.Append(new P.SlideLayoutId
            {
                Id = nextLayoutId,
                RelationshipId = masterPart.GetIdOfPart(emptyLayoutPart),
            });
            masterPart.SlideMaster.Save();

            slidePart.DeletePart(slidePart.SlideLayoutPart!);
            slidePart.AddPart(emptyLayoutPart);

            slidePart.Slide!.Save();
            presentationPart.Presentation!.Save();
        }

        return ms.ToArray();
    }

    /// <summary>Rewrites the sample deck's single text-box paragraph as the given runs.</summary>
    public static byte[] SampleWithRuns(params (string Text, bool Bold)[] runs) => Mutate(slide =>
    {
        var paragraph = slide.Descendants<A.Paragraph>().First(p => p.Descendants<A.Text>().Any());
        paragraph.RemoveAllChildren();

        foreach (var (text, bold) in runs)
        {
            var properties = new A.RunProperties { Language = "en-US" };
            if (bold) properties.Bold = true;
            paragraph.AppendChild(new A.Run(properties, new A.Text(text)));
        }
    });

    /// <summary>
    /// The sample deck plus a real one-cell table in a p:graphicFrame. Table text lives in
    /// a:tc/a:txBody, not in a p:sp, so it is invisible to any walk that only looks at shapes.
    /// </summary>
    public static byte[] SampleWithTableCell(string cellText) => Mutate(slide =>
        slide.CommonSlideData!.ShapeTree!.AppendChild(new P.GraphicFrame(
            new P.NonVisualGraphicFrameProperties(
                new P.NonVisualDrawingProperties { Id = 99U, Name = "Table 1" },
                new P.NonVisualGraphicFrameDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.Transform(
                new A.Offset { X = 0L, Y = 0L },
                new A.Extents { Cx = 3000000L, Cy = 500000L }),
            new A.Graphic(
                new A.GraphicData(
                    new A.Table(
                        new A.TableProperties(),
                        new A.TableGrid(new A.GridColumn { Width = 3000000L }),
                        new A.TableRow(
                            new A.TableCell(
                                new A.TextBody(
                                    new A.BodyProperties(),
                                    new A.ListStyle(),
                                    new A.Paragraph(new A.Run(
                                        new A.RunProperties { Language = "en-US" },
                                        new A.Text(cellText)))),
                                new A.TableCellProperties()))
                        { Height = 500000L }))
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" }))));

    /// <summary>
    /// The sample deck plus a real p:grpSp holding one shape whose text is nothing but
    /// <paramref name="placeholder"/>. ReplaceImage walks only a slide's direct p:sp children
    /// (see the comment in <c>PresentationEditor.ReplaceImageCore</c>), so this is what exercises
    /// "the placeholder genuinely exists on the slide, but only inside a group" rather than "does
    /// not exist at all" — the two must produce different refusal messages.
    /// </summary>
    public static byte[] SampleWithPlaceholderInGroup(string placeholder) => Mutate(slide =>
        slide.CommonSlideData!.ShapeTree!.AppendChild(new P.GroupShape(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = 98U, Name = "Group 1" },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(
                new A.TransformGroup(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 1000000L, Cy = 1000000L },
                    new A.ChildOffset { X = 0L, Y = 0L },
                    new A.ChildExtents { Cx = 1000000L, Cy = 1000000L })),
            new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties { Id = 97U, Name = "Grouped Shape" },
                    new P.NonVisualShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset { X = 0L, Y = 0L },
                        new A.Extents { Cx = 500000L, Cy = 500000L }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
                new P.TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(
                        new A.RunProperties { Language = "en-US" },
                        new A.Text(placeholder))))))));

    /// <summary>Formatting of every run on the deck's first slide, in document order.</summary>
    public static (string Text, bool Bold)[] RunsOfFirstSlide(byte[] pptx)
    {
        using var ms = Load(pptx);
        using var doc = PresentationDocument.Open(ms, false);
        var slide = doc.PresentationPart!.SlideParts.First().Slide!;
        return slide.Descendants<A.Run>()
                    .Select(r => (r.Text?.Text ?? string.Empty, r.RunProperties?.Bold?.Value == true))
                    .ToArray();
    }

    /// <summary>
    /// A one-slide deck holding a single shape at the exact position given, whose text is
    /// nothing but <paramref name="placeholder"/>.
    ///
    /// Deliberately a DRAWN box — an explicit <c>a:xfrm</c> — because that is what a designer
    /// produces in PowerPoint, and what ReplaceImage must require: a shape that inherits its
    /// position from a layout carries no <c>a:xfrm</c> of its own, and that is exactly the case
    /// the feature has to reject. Built the same way every fixture in this file is — by cloning
    /// the sample deck rather than assembling presentation/master/layout/theme parts from
    /// nothing (see the class doc comment) — and the sample's sole shape already carries its
    /// own <c>a:xfrm</c>, so this only has to overwrite its offset/extents and its text rather
    /// than manufacture a new shape.
    /// </summary>
    public static byte[] DeckWithPlaceholderBox(
        string placeholder, long x = 1000000, long y = 2000000, long cx = 4000000, long cy = 3000000) =>
        Mutate(slide =>
        {
            var shape = slide.CommonSlideData!.ShapeTree!.Elements<P.Shape>().Single();
            var xfrm = shape.ShapeProperties!.Transform2D!;
            xfrm.Offset = new A.Offset { X = x, Y = y };
            xfrm.Extents = new A.Extents { Cx = cx, Cy = cy };

            slide.Descendants<A.Text>().First().Text = placeholder;
        });

    /// <summary>
    /// A one-slide deck holding a single shape whose text is nothing but
    /// <paramref name="placeholder"/>, but with no <c>a:xfrm</c> of its own — the case
    /// <see cref="DeckWithPlaceholderBox"/>'s doc comment says ReplaceImage must reject: a shape
    /// that inherits its position from a layout rather than one a designer drew. Removing the
    /// sample shape's own <c>a:xfrm</c> leaves the deck schema-valid, since <c>p:spPr</c> with no
    /// <c>a:xfrm</c> child is valid — the shape then simply inherits whatever position its layout
    /// gives it, which is exactly the "nowhere to put the image" case.
    /// </summary>
    public static byte[] DeckWithUnpositionedPlaceholder(string placeholder) => Mutate(slide =>
    {
        var shape = slide.CommonSlideData!.ShapeTree!.Elements<P.Shape>().Single();
        shape.ShapeProperties!.Transform2D!.Remove();

        slide.Descendants<A.Text>().First().Text = placeholder;
    });

    /// <summary>Schema-validation errors for the whole package (empty means valid).</summary>
    public static IReadOnlyList<ValidationErrorInfo> Validate(byte[] pptx)
    {
        using var ms = Load(pptx);
        using var doc = PresentationDocument.Open(ms, false);
        return new OpenXmlValidator().Validate(doc).ToList();
    }

    private static byte[] Mutate(Action<P.Slide> mutate)
    {
        using var ms = Load(Sample());
        using (var doc = PresentationDocument.Open(ms, true))
        {
            var slidePart = doc.PresentationPart!.SlideParts.Single();
            mutate(slidePart.Slide!);
            slidePart.Slide!.Save();
        }

        return ms.ToArray();
    }

    private static void SetSoleText(SlidePart slidePart, string text)
    {
        slidePart.Slide!.Descendants<A.Text>().First().Text = text;
        slidePart.Slide.Save();
    }

    private static MemoryStream Load(byte[] bytes)
    {
        var ms = new MemoryStream();
        ms.Write(bytes, 0, bytes.Length);
        ms.Position = 0;
        return ms;
    }
}
