using ShapeCrawler;

namespace DocToolkit;

/// <summary>Opens and edits PowerPoint (.pptx) presentations.</summary>
public static class PresentationEditor
{
    /// <summary>Number of slides in the deck.</summary>
    public static int SlideCount(byte[] pptx)
    {
        using var pres = Open(pptx);
        return pres.Slides.Count;
    }

    /// <summary>All text found on every slide, one entry per text-bearing shape.</summary>
    public static IReadOnlyList<string> ExtractText(byte[] pptx)
    {
        using var pres = Open(pptx);
        var results = new List<string>();

        // slide.GetTexts() returns IList<ITextBox> in the installed version (0.79.4), not
        // IList<string> as the brief's API note stated — project each box's .Text instead.
        for (var n = 1; n <= pres.Slides.Count; n++)
            results.AddRange(pres.Slide(n).GetTexts().Select(tb => tb.Text));

        return results;
    }

    /// <summary>Replaces every key with its value in all text boxes, returning updated bytes.</summary>
    public static byte[] ReplaceText(byte[] pptx, IReadOnlyDictionary<string, string> replacements)
    {
        ArgumentNullException.ThrowIfNull(replacements);

        using var pres = Open(pptx);
        for (var n = 1; n <= pres.Slides.Count; n++)
        {
            foreach (var shape in pres.Slide(n).Shapes)
            {
                var textBox = shape.TextBox;
                if (textBox is null) continue;

                var original = textBox.Text;
                var updated = original;
                foreach (var (key, value) in replacements)
                    updated = updated.Replace(key, value ?? string.Empty);

                if (updated != original) textBox.SetText(updated);
            }
        }

        using var ms = new MemoryStream();
        pres.Save(ms);
        return ms.ToArray();
    }

    private static Presentation Open(byte[] pptx)
    {
        ArgumentNullException.ThrowIfNull(pptx);
        if (pptx.Length == 0)
            throw new ArgumentException("Presentation content was empty.", nameof(pptx));

        try
        {
            var ms = new MemoryStream();
            ms.Write(pptx, 0, pptx.Length);
            ms.Position = 0;
            return new Presentation(ms);
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException("Failed to open PPTX.", ex);
        }
    }
}
