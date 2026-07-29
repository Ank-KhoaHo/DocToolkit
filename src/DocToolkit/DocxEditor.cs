using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit;

/// <summary>Opens and edits an existing .docx package.</summary>
public static class DocxEditor
{
    /// <summary>
    /// Replaces every key with its value across the document body, its headers and footers, and
    /// its footnotes and endnotes.
    ///
    /// Word routinely splits a single visible word across several &lt;w:t&gt; runs (spell-check
    /// state, formatting changes, a language switch), so a naive per-run replace misses any
    /// placeholder that straddles a run boundary. Substitution therefore happens against the
    /// concatenated text of each paragraph, but the result is spliced back into only the runs the
    /// match actually overlaps: runs outside a match — including the runs inside a
    /// &lt;w:hyperlink&gt; — keep their text and their formatting untouched. When a placeholder
    /// does straddle runs, the replacement value is written into the run holding its first
    /// character and so inherits that run's formatting.
    ///
    /// Text boxes (&lt;w:txbxContent&gt;) nest whole paragraphs inside a run of the enclosing
    /// paragraph. They are treated as the separate paragraphs they are, so a placeholder inside a
    /// text box is replaced and a text box without one is left alone.
    ///
    /// Keys are matched in a single left-to-right pass and the longest key wins at any given
    /// offset, so a substituted value is never rescanned for further placeholders.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or edited.</exception>
    public static byte[] ReplaceText(byte[] docx, IReadOnlyDictionary<string, string> replacements)
    {
        ArgumentNullException.ThrowIfNull(docx);
        ArgumentNullException.ThrowIfNull(replacements);
        if (docx.Length == 0)
            throw new ArgumentException("DOCX content was empty.", nameof(docx));

        using var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        ms.Position = 0;

        try
        {
            using (var doc = WordprocessingDocument.Open(ms, true))
            {
                var main = doc.MainDocumentPart
                           ?? throw new DocumentConversionException("Document has no main part.");
                var body = main.Document?.Body
                           ?? throw new DocumentConversionException("Document has no body.");

                ReplaceIn(body, replacements);
                main.Document!.Save();

                // A placeholder in a header or footer used to come back unreplaced with no error
                // at all, which is a silent wrong answer for the "fill a template" use case.
                foreach (var part in main.HeaderParts)
                {
                    if (part.Header is null) continue;
                    ReplaceIn(part.Header, replacements);
                    part.Header.Save();
                }

                foreach (var part in main.FooterParts)
                {
                    if (part.Footer is null) continue;
                    ReplaceIn(part.Footer, replacements);
                    part.Footer.Save();
                }

                if (main.FootnotesPart?.Footnotes is { } footnotes)
                {
                    ReplaceIn(footnotes, replacements);
                    footnotes.Save();
                }

                if (main.EndnotesPart?.Endnotes is { } endnotes)
                {
                    ReplaceIn(endnotes, replacements);
                    endnotes.Save();
                }
            }
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to edit DOCX.", ex);
        }

        return ms.ToArray();
    }

    /// <summary>
    /// Returns the plain text of the document body. Headers, footers, footnotes and endnotes are
    /// <b>not</b> included — call <see cref="ExtractText(byte[], bool)"/> for those.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static string ExtractText(byte[] docx) => ExtractText(docx, includeHeadersAndFooters: false);

    /// <summary>
    /// Returns the plain text of the document. When <paramref name="includeHeadersAndFooters"/> is
    /// true the body text is followed by each header part and then each footer part, separated by
    /// newlines; footnotes and endnotes are never included.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static string ExtractText(byte[] docx, bool includeHeadersAndFooters)
    {
        ArgumentNullException.ThrowIfNull(docx);
        if (docx.Length == 0)
            throw new ArgumentException("DOCX content was empty.", nameof(docx));

        try
        {
            using var ms = new MemoryStream(docx, writable: false);
            using var doc = WordprocessingDocument.Open(ms, false);

            var main = doc.MainDocumentPart;
            var bodyText = main?.Document?.Body?.InnerText ?? string.Empty;
            if (!includeHeadersAndFooters || main is null) return bodyText;

            var sb = new StringBuilder(bodyText);
            foreach (var text in main.HeaderParts.Select(p => p.Header?.InnerText)
                                     .Concat(main.FooterParts.Select(p => p.Footer?.InnerText)))
            {
                if (string.IsNullOrEmpty(text)) continue;
                if (sb.Length > 0) sb.Append('\n');
                sb.Append(text);
            }

            return sb.ToString();
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read DOCX.", ex);
        }
    }

    private static void ReplaceIn(OpenXmlElement root, IReadOnlyDictionary<string, string> replacements)
    {
        foreach (var paragraph in root.Descendants<Paragraph>())
            ReplaceInParagraph(paragraph, replacements);
    }

    private static void ReplaceInParagraph(Paragraph paragraph, IReadOnlyDictionary<string, string> replacements)
    {
        // Only the text this paragraph owns directly. A text box nests entire w:p elements inside
        // a run of this paragraph, and Descendants<Text>() walks straight into them; folding that
        // text into this paragraph's merged string relocated the text box's content on every
        // replacement. Those nested paragraphs are visited on their own by the caller's walk.
        var texts = paragraph.Descendants<Text>()
                             .Where(t => t.Ancestors<Paragraph>().FirstOrDefault() == paragraph)
                             .ToList();
        if (texts.Count == 0) return;

        RunTextSplicer.Apply(texts, static t => t.Text, WriteText, replacements);
    }

    private static void WriteText(Text node, string value)
    {
        node.Text = value;

        // Leading or trailing whitespace is dropped by consumers unless the run opts out of
        // whitespace collapsing.
        if (value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1])))
            node.Space = SpaceProcessingModeValues.Preserve;
    }
}
