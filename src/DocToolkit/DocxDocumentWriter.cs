using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit;

/// <summary>
/// Turns a list of <see cref="DocxBlock"/> into a WordprocessingML package.
///
/// Separate from <see cref="DocxEditor"/> on purpose: that file is already the largest in the repo,
/// and creating a document has nothing in common with editing one beyond the format.
/// </summary>
internal static class DocxDocumentWriter
{
    /// <summary>
    /// Builds the package into a fresh <see cref="MemoryStream"/>, positioned at 0. The caller owns
    /// and disposes it.
    /// </summary>
    public static MemoryStream Write(IReadOnlyList<DocxBlock> blocks)
    {
        var ms = new MemoryStream();

        try
        {
            using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var main = doc.AddMainDocumentPart();
                var body = new Body();
                main.Document = new Document(body);

                // A monotonic counter is enough here, unlike DocxEditor.ReplaceImage's NextDrawingId
                // scan: that has to coexist with drawings a template already contained, whereas this
                // document starts empty and every id in it is one we issued.
                var nextDrawingId = 1U;

                foreach (var block in blocks)
                    AppendBlock(main, body, block, ref nextDrawingId);

                AddHeadingStyles(main, blocks);

                main.Document.Save();
            }

            ms.Position = 0;
            return ms;
        }
        // Two arms, both disposing. A single arm filtered with
        // `when (ex is not DocumentConversionException)` looks equivalent and is not: a filtered
        // catch that does not match never runs its body, so a DocumentConversionException raised
        // inside the try - which AppendBlock's default case does - would escape with `ms` still
        // open. The caller cannot dispose a stream it was never handed.
        catch (DocumentConversionException)
        {
            ms.Dispose();
            throw;
        }
        catch (Exception ex)
        {
            ms.Dispose();
            throw new DocumentConversionException("Failed to create DOCX.", ex);
        }
    }

    private static void AppendBlock(
        MainDocumentPart main, Body body, DocxBlock block, ref uint nextDrawingId)
    {
        switch (block)
        {
            case ParagraphBlock p:
                body.AppendChild(TextParagraph(p.Text));
                break;

            case HeadingBlock h:
                body.AppendChild(TextParagraph(h.Text, $"Heading{h.Level}"));
                break;

            default:
                // Reached whenever a block type has no case above. That is NOT merely theoretical
                // while this feature is being built: Heading, Table and Image are already public
                // factories, so until their cases land this fires on ordinary calls. Once every
                // block type is handled it becomes a guard - a new block type added without a case
                // here fails loudly instead of silently dropping content.
                //
                // Do not describe this as unreachable. An earlier version did, and the claim was
                // false at the time it was written.
                throw new DocumentConversionException(
                    $"No writer case handles {block.GetType().Name}. This is a bug in DocToolkit.");
        }
    }

    /// <summary>
    /// A paragraph holding one run.
    ///
    /// <c>xml:space="preserve"</c> is set because without it leading and trailing spaces are
    /// stripped by the consumer, which turns "Total: " into "Total:" with nothing to show for it.
    /// </summary>
    private static Paragraph TextParagraph(string text, string? styleId = null)
    {
        var paragraph = new Paragraph();

        if (styleId is not null)
            paragraph.AppendChild(new ParagraphProperties(new ParagraphStyleId { Val = styleId }));

        paragraph.AppendChild(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
        return paragraph;
    }

    /// <summary>
    /// Half-point font sizes per heading level, matching Word's own built-in defaults closely
    /// enough that a document looks unremarkable when opened. w:sz is in HALF-points, so 32 is 16pt.
    /// </summary>
    private static readonly int[] HeadingHalfPointSizes = { 32, 28, 26, 24, 22, 20 };

    /// <summary>
    /// Defines a real style for each heading level actually used.
    ///
    /// Referencing a style without defining it is the silent failure this exists to prevent: Word
    /// renders the paragraph as ordinary text, the package stays schema-valid, and nothing reports
    /// a problem. Only the levels used are defined, so a document with one heading does not carry
    /// six unused style definitions.
    /// </summary>
    private static void AddHeadingStyles(MainDocumentPart main, IReadOnlyList<DocxBlock> blocks)
    {
        var levels = blocks.OfType<HeadingBlock>().Select(h => h.Level).Distinct().OrderBy(l => l).ToList();
        if (levels.Count == 0) return;

        var styles = new Styles();

        // Every heading is basedOn Normal, so Normal has to exist or the basedOn dangles.
        styles.AppendChild(new Style(new StyleName { Val = "Normal" })
        {
            Type = StyleValues.Paragraph,
            StyleId = "Normal",
            Default = true,
        });

        foreach (var level in levels)
        {
            styles.AppendChild(new Style(
                new StyleName { Val = $"heading {level}" },
                new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(
                    new KeepNext(),
                    new OutlineLevel { Val = level - 1 }),
                new StyleRunProperties(
                    new Bold(),
                    new FontSize { Val = HeadingHalfPointSizes[level - 1].ToString(CultureInfo.InvariantCulture) }))
            {
                Type = StyleValues.Paragraph,
                StyleId = $"Heading{level}",
            });
        }

        var part = main.AddNewPart<StyleDefinitionsPart>();
        part.Styles = styles;
        part.Styles.Save();
    }
}
