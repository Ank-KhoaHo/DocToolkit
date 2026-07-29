using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit;

/// <summary>Opens and edits an existing .docx package.</summary>
public static class DocxEditor
{
    /// <summary>
    /// Replaces every key with its value across the document body.
    ///
    /// Word often splits a single visible word across several &lt;w:t&gt; runs (spell-check state,
    /// formatting changes), so a naive per-run replace misses placeholders. This merges the runs
    /// of each paragraph before substituting.
    /// </summary>
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
                var body = doc.MainDocumentPart?.Document?.Body
                           ?? throw new DocumentConversionException("Document has no body.");

                foreach (var paragraph in body.Descendants<Paragraph>())
                {
                    var texts = paragraph.Descendants<Text>().ToList();
                    if (texts.Count == 0) continue;

                    var merged = string.Concat(texts.Select(t => t.Text));
                    var updated = merged;
                    foreach (var (key, value) in replacements)
                        updated = updated.Replace(key, value ?? string.Empty);

                    if (updated == merged) continue;

                    // Put all text on the first run and blank the rest, preserving its formatting.
                    texts[0].Text = updated;
                    texts[0].Space = SpaceProcessingModeValues.Preserve;
                    for (var i = 1; i < texts.Count; i++) texts[i].Text = string.Empty;
                }

                doc.MainDocumentPart!.Document.Save();
            }
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to edit DOCX.", ex);
        }

        return ms.ToArray();
    }

    /// <summary>Returns the plain text of the document body.</summary>
    public static string ExtractText(byte[] docx)
    {
        ArgumentNullException.ThrowIfNull(docx);
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart?.Document?.Body?.InnerText ?? string.Empty;
    }
}
