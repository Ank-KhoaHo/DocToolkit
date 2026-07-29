using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocToolkit.Tests;

/// <summary>
/// Fixtures derived from the committed one-slide sample deck.
///
/// Building a valid .pptx from nothing needs presentation, master, layout and theme parts, so
/// these clone the real fixture and mutate it instead: extra slides come from cloning the sample
/// slide part, and the deck order is then set independently of the part order so the two can be
/// told apart.
/// </summary>
internal static class PptxFixtures
{
    private static readonly string SampleAssetPath =
        Path.Combine(AppContext.BaseDirectory, "assets", "sample.pptx");

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
            var slideIdList = presentationPart.Presentation.SlideIdList!;
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
