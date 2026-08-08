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

            case TableBlock t:
                body.AppendChild(BuildTable(t));
                break;

            case ImageBlock i:
                body.AppendChild(BuildImageParagraph(main, i, ref nextDrawingId));
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
    /// Adds the image as a part of <paramref name="main"/> and returns a paragraph holding an
    /// inline drawing that references it.
    ///
    /// The part is added to the part that OWNS the paragraph — here always the main document. That
    /// is not incidental: a relationship id is scoped to its container, so a part added to the wrong
    /// one still produces a schema-valid package in which Word shows nothing at all. There is no
    /// error and no exception; the image is simply absent.
    ///
    /// The content type comes from the image's own magic bytes via <see cref="ImageInspector"/>,
    /// never from a filename or a caller's assertion — a part declaring <c>image/png</c> while
    /// holding JPEG bytes renders as a blank frame, silently.
    /// </summary>
    private static Paragraph BuildImageParagraph(
        MainDocumentPart main, ImageBlock block, ref uint nextDrawingId)
    {
        var info = ImageInspector.Inspect(block.Bytes);
        var (widthEmu, heightEmu) = ImageInspector.Resolve(info, block.WidthPoints, block.HeightPoints);

        // The same content-type DERIVATION as DocxEditor.ReplaceImage, so the two paths cannot
        // disagree about what a given set of bytes is. The mechanics differ and that is not a
        // claim worth making: ReplaceImage writes through GetStream(FileMode.Create), this uses
        // FeedData. Equivalent in effect, but "exactly how ReplaceImage does it" would be false,
        // and this repo has twice been misled by a comment that overstated a shared invariant.
        var part = main.AddNewPart<ImagePart>(info.ContentType);
        using (var source = new MemoryStream(block.Bytes))
            part.FeedData(source);

        var drawing = DrawingFactory.InlineImage(
            main.GetIdOfPart(part), $"Image {nextDrawingId}", nextDrawingId, widthEmu, heightEmu,
            block.AltText);

        nextDrawingId++;
        return new Paragraph(new Run(drawing));
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

    private static Table BuildTable(TableBlock block)
    {
        var columns = block.Headers.Count;

        // Child order is Top, Left, Bottom, Right, InsideH, InsideV - the CT_TblBorders
        // sequence. It is NOT the clockwise order it reads like. An earlier draft used
        // Top, Bottom, Left, Right and OpenXmlValidator rejected it with
        // "unexpected child element 'left'". Verified by mutation: restoring that order
        // fails the table test, and correcting it passes.
        var table = new Table(
            new TableProperties(
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4 },
                    new LeftBorder { Val = BorderValues.Single, Size = 4 },
                    new BottomBorder { Val = BorderValues.Single, Size = 4 },
                    new RightBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 })));

        // REQUIRED. Without a w:tblGrid the validator rejects the first w:tr as an unexpected child.
        var grid = new TableGrid();
        for (var i = 0; i < columns; i++) grid.AppendChild(new GridColumn());
        table.AppendChild(grid);

        table.AppendChild(BuildRow(block.Headers.Select(h => (object?)h), columns, bold: true));

        foreach (var row in block.Rows)
            table.AppendChild(BuildRow(row, columns, bold: false));

        return table;
    }

    /// <summary>
    /// One row, padded to <paramref name="columns"/>.
    ///
    /// A row shorter than the header is padded rather than rejected: ragged data is normal, and a
    /// table cell count that disagrees with the grid is what makes Word show a malformed table.
    ///
    /// An over-long row is rejected by <see cref="DocxBlock.Table"/> before it can reach here, which
    /// is where the useful diagnostic lives — it names the caller's own line. The throw below is not
    /// dead code dressed as a guard: <c>TableBlock</c> is internal, so internal code can still build
    /// one directly, and <c>OpenXmlValidator</c> was measured to catch a row/grid count mismatch in
    /// NEITHER direction. Without this, that mistake would produce a silently malformed table.
    /// </summary>
    private static TableRow BuildRow(IEnumerable<object?> values, int columns, bool bold)
    {
        var row = new TableRow();
        var written = 0;

        foreach (var value in values)
        {
            if (written == columns)
                throw new InvalidOperationException(
                    $"A table row has more than {columns} " +
                    $"{(columns == 1 ? "cell" : "cells")}. DocxBlock.Table rejects this, so a " +
                    "TableBlock was built internally without going through it.");

            row.AppendChild(BuildCell(Format(value), bold));
            written++;
        }

        for (; written < columns; written++)
            row.AppendChild(BuildCell(string.Empty, bold));

        return row;
    }

    private static TableCell BuildCell(string text, bool bold)
    {
        var run = new Run();
        if (bold) run.AppendChild(new RunProperties(new Bold()));
        run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

        // A w:tc must contain at least one block-level element; an empty cell still needs a w:p.
        return new TableCell(new Paragraph(run));
    }

    /// <summary>
    /// Formats a cell value, handling the same types BY NAME as
    /// <see cref="WorkbookEditor.Create(string, System.Collections.Generic.IEnumerable{System.Collections.Generic.IEnumerable{object}})"/>
    /// but NOT always to the same text. See <see cref="DocxBlock.Table"/> for the measured
    /// divergences and why they are deliberate.
    ///
    /// One exception is not by-name at all: <see cref="XlsxFormula"/> is deliberately not one of
    /// the types handled below. A formula is meaningful only in a spreadsheet cell, so a value of
    /// that type falls to the default arm and renders as the literal text
    /// <c>DocToolkit.XlsxFormula</c> rather than the formula text.
    ///
    /// A Word table cell has no type of its own, unlike a spreadsheet cell, so every value ends up
    /// as text, and choosing that text is this library's call rather than a renderer's. What the two
    /// paths do guarantee identically is culture-invariance: otherwise the same code produces a
    /// different document depending on the machine's regional settings.
    /// </summary>
    private static string Format(object? value) => value switch
    {
        null => string.Empty,
        string s => s,
        bool b => b ? "TRUE" : "FALSE",
        DateTime d => d.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
        DateOnly d => d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
        TimeOnly t => t.ToString("HH:mm:ss", CultureInfo.InvariantCulture),
        TimeSpan ts => ts.ToString(null, CultureInfo.InvariantCulture),
        IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
        _ => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty,
    };
}
