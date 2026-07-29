using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocToolkit;

/// <summary>Opens and edits PowerPoint (.pptx) presentations.</summary>
public static class PresentationEditor
{
    /// <summary>Number of slides in the deck, as counted from the deck's slide list.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static int SlideCount(byte[] pptx)
    {
        using var ms = OpenForWrite(pptx);

        try
        {
            using var doc = OpenDocument(ms, false);
            return SlidesInDeckOrder(PresentationPartOf(doc)).Count();
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read PPTX.", ex);
        }
    }

    /// <summary>
    /// All text found on every slide, one entry per text-bearing body, in the order the deck is
    /// presented rather than the order the slide parts happen to be related.
    ///
    /// A "text-bearing body" is any element holding &lt;a:p&gt; paragraphs: an ordinary shape's
    /// &lt;p:txBody&gt;, a shape nested in a group, and a table cell's &lt;a:txBody&gt; alike.
    /// Paragraphs within one body are joined with newlines. This is deliberately the same walk
    /// <see cref="ReplaceText"/> performs, so anything this reports is something that can be
    /// replaced and vice versa. Speaker notes and slide masters/layouts are not included.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static IReadOnlyList<string> ExtractText(byte[] pptx)
    {
        using var ms = OpenForWrite(pptx);

        try
        {
            using var doc = OpenDocument(ms, false);

            var results = new List<string>();
            foreach (var slidePart in SlidesInDeckOrder(PresentationPartOf(doc)))
            {
                var slide = slidePart.Slide;
                if (slide is null) continue;

                // PowerPoint stores shape text as a:t runs under a:p paragraphs - Wordprocessing's
                // w:t is the DOCX equivalent, not what PPTX uses. Grouping by the paragraph's
                // parent yields one entry per text body while preserving document order.
                foreach (var body in slide.Descendants<A.Paragraph>().GroupBy(p => p.Parent))
                {
                    var bodyText = string.Join(
                        "\n",
                        body.Select(p => string.Concat(p.Descendants<A.Text>().Select(t => t.Text))));

                    if (bodyText.Length > 0) results.Add(bodyText);
                }
            }

            return results;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read PPTX.", ex);
        }
    }

    /// <summary>
    /// Replaces every key with its value across all slide text, returning updated bytes.
    ///
    /// PowerPoint routinely splits a single visible word across several &lt;a:t&gt; runs
    /// (spell-check state, formatting changes), so a naive per-run replace misses any placeholder
    /// that straddles a run boundary. Substitution therefore happens against the concatenated text
    /// of each paragraph, but the result is spliced back into only the runs the match actually
    /// overlaps: runs outside a match keep their text and their formatting untouched. When a
    /// placeholder does straddle runs, the replacement value is written into the run holding its
    /// first character and so inherits that run's formatting.
    ///
    /// Keys are matched in a single left-to-right pass and the longest key wins at any given
    /// offset, so a substituted value is never rescanned for further placeholders. Slides are
    /// visited in deck order; speaker notes and slide masters/layouts are not touched.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or edited.</exception>
    public static byte[] ReplaceText(byte[] pptx, IReadOnlyDictionary<string, string> replacements)
    {
        ArgumentNullException.ThrowIfNull(replacements);

        using var ms = OpenForWrite(pptx);
        try
        {
            using (var doc = OpenDocument(ms, true))
            {
                foreach (var slidePart in SlidesInDeckOrder(PresentationPartOf(doc)))
                {
                    var slide = slidePart.Slide;
                    if (slide is null) continue;

                    foreach (var paragraph in slide.Descendants<A.Paragraph>())
                    {
                        var texts = paragraph.Descendants<A.Text>().ToList();
                        if (texts.Count == 0) continue;

                        RunTextSplicer.Apply(
                            texts, static t => t.Text, static (t, v) => t.Text = v, replacements);
                    }

                    slide.Save();
                }
            }
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to edit PPTX.", ex);
        }

        return ms.ToArray();
    }

    /// <summary>
    /// Slide parts in the order the deck presents them.
    ///
    /// <c>PresentationPart.SlideParts</c> is part-relationship order, which has nothing to do with
    /// slide order - reorder a deck in PowerPoint and the two diverge completely. The authoritative
    /// order is <c>p:sldIdLst</c>, so resolve each <c>p:sldId</c> through its relationship id.
    /// </summary>
    private static IEnumerable<SlidePart> SlidesInDeckOrder(PresentationPart presentationPart)
    {
        var slideIdList = presentationPart.Presentation?.SlideIdList;
        if (slideIdList is null)
        {
            // No slide list at all: nothing better to go on than the relationship order.
            foreach (var slidePart in presentationPart.SlideParts) yield return slidePart;
            yield break;
        }

        foreach (var slideId in slideIdList.Elements<P.SlideId>())
        {
            var relationshipId = slideId.RelationshipId?.Value;
            if (string.IsNullOrEmpty(relationshipId)) continue;

            OpenXmlPart part;
            try { part = presentationPart.GetPartById(relationshipId); }
            catch (ArgumentOutOfRangeException) { continue; }

            if (part is SlidePart slidePart) yield return slidePart;
        }
    }

    private static PresentationPart PresentationPartOf(PresentationDocument doc)
        => doc.PresentationPart
           ?? throw new DocumentConversionException("Presentation has no presentation part.");

    private static PresentationDocument OpenDocument(MemoryStream ms, bool isEditable)
    {
        try
        {
            return PresentationDocument.Open(ms, isEditable);
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException("Failed to open PPTX.", ex);
        }
    }

    private static MemoryStream OpenForWrite(byte[] pptx)
    {
        ArgumentNullException.ThrowIfNull(pptx);
        if (pptx.Length == 0)
            throw new ArgumentException("Presentation content was empty.", nameof(pptx));

        var ms = new MemoryStream();
        ms.Write(pptx, 0, pptx.Length);
        ms.Position = 0;
        return ms;
    }
}
