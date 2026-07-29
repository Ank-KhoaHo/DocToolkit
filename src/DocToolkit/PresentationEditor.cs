using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace DocToolkit;

/// <summary>Opens and edits PowerPoint (.pptx) presentations.</summary>
public static class PresentationEditor
{
    /// <summary>Number of slides in the deck.</summary>
    public static int SlideCount(byte[] pptx)
    {
        using var ms = OpenForWrite(pptx);
        using var doc = OpenDocument(ms, false);

        var presentationPart = doc.PresentationPart
            ?? throw new DocumentConversionException("Presentation has no presentation part.");
        return presentationPart.SlideParts.Count();
    }

    /// <summary>All text found on every slide, one entry per text-bearing shape.</summary>
    public static IReadOnlyList<string> ExtractText(byte[] pptx)
    {
        using var ms = OpenForWrite(pptx);
        using var doc = OpenDocument(ms, false);

        var presentationPart = doc.PresentationPart
            ?? throw new DocumentConversionException("Presentation has no presentation part.");

        var results = new List<string>();
        foreach (var slidePart in presentationPart.SlideParts)
        {
            var slide = slidePart.Slide;
            if (slide is null) continue;

            // PowerPoint stores shape text as a:t runs under a:p paragraphs within a shape's
            // p:txBody — Wordprocessing.Text (w:t) is the DOCX equivalent, not what PPTX uses.
            foreach (var shape in slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>())
            {
                var txBody = shape.TextBody;
                if (txBody is null) continue;

                var shapeText = string.Join(
                    "\n",
                    txBody.Elements<A.Paragraph>()
                        .Select(p => string.Concat(p.Descendants<A.Text>().Select(t => t.Text))));

                if (shapeText.Length > 0) results.Add(shapeText);
            }
        }

        return results;
    }

    /// <summary>
    /// Replaces every key with its value across all slide text, returning updated bytes.
    ///
    /// PowerPoint often splits a single visible word across several &lt;a:t&gt; runs (spell-check
    /// state, formatting changes), so a naive per-run replace misses placeholders that straddle
    /// runs. This merges the runs of each paragraph before substituting, then writes the result
    /// back to the first run and blanks the others — mirrors the approach in DocxEditor.
    /// </summary>
    public static byte[] ReplaceText(byte[] pptx, IReadOnlyDictionary<string, string> replacements)
    {
        ArgumentNullException.ThrowIfNull(replacements);

        using var ms = OpenForWrite(pptx);
        try
        {
            using (var doc = OpenDocument(ms, true))
            {
                var presentationPart = doc.PresentationPart
                    ?? throw new DocumentConversionException("Presentation has no presentation part.");

                foreach (var slidePart in presentationPart.SlideParts)
                {
                    var slide = slidePart.Slide;
                    if (slide is null) continue;

                    foreach (var paragraph in slide.Descendants<A.Paragraph>())
                    {
                        var texts = paragraph.Descendants<A.Text>().ToList();
                        if (texts.Count == 0) continue;

                        var merged = string.Concat(texts.Select(t => t.Text));
                        var updated = merged;
                        foreach (var (key, value) in replacements)
                            updated = updated.Replace(key, value ?? string.Empty);

                        if (updated == merged) continue;

                        // Put all text on the first run and blank the rest, preserving formatting.
                        texts[0].Text = updated;
                        for (var i = 1; i < texts.Count; i++) texts[i].Text = string.Empty;
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
