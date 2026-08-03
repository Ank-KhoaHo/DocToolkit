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

        ReplaceTextCore(ms, replacements);
        return ms.ToArray();
    }

    /// <summary>
    /// Reads a .docx from <paramref name="source"/>, replaces every key with its value, and writes
    /// the result to <paramref name="destination"/>. See <see cref="ReplaceText"/> for exactly what
    /// counts as a match and how formatting survives it — this overload applies the identical logic
    /// via <paramref name="source"/> and <paramref name="destination"/> instead of a byte array.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought, and neither has to be seekable, so
    /// both may be sockets, files or HTTP message bodies.
    /// </summary>
    /// <param name="source">The stream the .docx package is read from.</param>
    /// <param name="replacements">Each key is replaced by its value, longest key wins per match.</param>
    /// <param name="destination">The stream the edited .docx package is written to.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="destination"/>
    /// is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or edited.</exception>
    public static async Task ReplaceTextAsync(
        Stream source, IReadOnlyDictionary<string, string> replacements, Stream destination,
        CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(replacements);
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, "DOCX content was empty.", nameof(source), "Failed to edit DOCX.", ct)
            .ConfigureAwait(false);

        ReplaceTextCore(docx, replacements);

        await StreamPipeline.EmitAsync(docx, destination, "Failed to edit DOCX.", ct).ConfigureAwait(false);
    }

    private static void ReplaceTextCore(MemoryStream ms, IReadOnlyDictionary<string, string> replacements)
    {
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

        using var ms = new MemoryStream(docx, writable: false);
        return ExtractTextCore(ms, includeHeadersAndFooters);
    }

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and returns the plain text of its body.
    /// Headers, footers, footnotes and endnotes are <b>not</b> included — call
    /// <see cref="ExtractTextAsync(Stream, bool, CancellationToken)"/> for those.
    /// <paramref name="source"/> is <b>read</b> to its end and is neither disposed, closed nor
    /// sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static Task<string> ExtractTextAsync(Stream source, CancellationToken ct = default)
        => ExtractTextAsync(source, includeHeadersAndFooters: false, ct);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and returns its plain text. When
    /// <paramref name="includeHeadersAndFooters"/> is true the body text is followed by each header
    /// part and then each footer part; footnotes and endnotes are never included.
    /// <paramref name="source"/> is <b>read</b> to its end and is neither disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static async Task<string> ExtractTextAsync(
        Stream source, bool includeHeadersAndFooters, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, "DOCX content was empty.", nameof(source), "Failed to read DOCX.", ct)
            .ConfigureAwait(false);

        return ExtractTextCore(docx, includeHeadersAndFooters);
    }

    private static string ExtractTextCore(Stream ms, bool includeHeadersAndFooters)
    {
        try
        {
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

    /// <summary>
    /// Expands a table row once per record, so a template can render a variable-length list such as
    /// invoice line items.
    ///
    /// A row is a <b>template row</b> when one of its cells contains a placeholder prefixed with
    /// <paramref name="collection"/> — <c>{{item.Desc}}</c> when <paramref name="collection"/> is
    /// <c>item</c>. Each record deep-clones that row, so every clone keeps the template's run
    /// formatting, cell shading and borders, and substitution runs through the same splicer
    /// <see cref="ReplaceText(byte[], IReadOnlyDictionary{string, string})"/> uses — a placeholder
    /// split across runs is still replaced, and a hyperlink in a cell is left intact.
    ///
    /// <b>Keys are bare field names</b> (<c>Desc</c>), not full placeholders — unlike
    /// <see cref="ReplaceText(byte[], IReadOnlyDictionary{string, string})"/>, whose keys are the
    /// placeholder text including braces. <paramref name="collection"/> is already an argument, so
    /// repeating it in every key of every record would duplicate it many times over.
    ///
    /// A placeholder with no matching key resolves to empty rather than staying visible.
    /// Placeholders for other prefixes are untouched, so a second call fills a second table. An
    /// empty <paramref name="rows"/> removes the template row, and removes the whole table when that
    /// row was its only one — an empty frame left on the page reads worse than rendering nothing.
    ///
    /// Compose with <see cref="ReplaceText(byte[], IReadOnlyDictionary{string, string})"/> for
    /// document-level scalars, expanding rows first.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="docx"/> is empty, or <paramref name="collection"/> is blank.
    /// </exception>
    /// <exception cref="DocumentConversionException">
    /// The package could not be opened or edited, or no template row was found for
    /// <paramref name="collection"/> — a mismatch between the call and the template is a bug in one
    /// of them, not a no-op.
    /// </exception>
    public static byte[] FillRows(
        byte[] docx, string collection, IEnumerable<IReadOnlyDictionary<string, string>> rows)
    {
        ArgumentNullException.ThrowIfNull(docx);
        ArgumentNullException.ThrowIfNull(collection);
        ArgumentNullException.ThrowIfNull(rows);
        if (docx.Length == 0)
            throw new ArgumentException("DOCX content was empty.", nameof(docx));
        if (string.IsNullOrWhiteSpace(collection))
            throw new ArgumentException("Collection name was blank.", nameof(collection));

        using var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        ms.Position = 0;

        FillRowsCore(ms, collection, rows);
        return ms.ToArray();
    }

    /// <summary>The one real implementation; every overload calls it so they cannot drift apart.</summary>
    private static void FillRowsCore(
        MemoryStream ms, string collection, IEnumerable<IReadOnlyDictionary<string, string>> rows)
    {
        var records = rows as IReadOnlyList<IReadOnlyDictionary<string, string>> ?? rows.ToList();
        var marker = "{{" + collection + ".";

        try
        {
            using (var doc = WordprocessingDocument.Open(ms, true))
            {
                var main = doc.MainDocumentPart
                           ?? throw new DocumentConversionException("Document has no main part.");
                var body = main.Document?.Body
                           ?? throw new DocumentConversionException("Document has no body.");

                var templates = TableRowFinder.Find(body, marker);
                if (templates.Count == 0)
                {
                    throw new DocumentConversionException(
                        $"No table row containing '{marker}' was found, so there was nothing to fill.");
                }

                foreach (var template in templates)
                    ExpandRow(template, collection, records);

                main.Document!.Save();
            }

            ms.Position = 0;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to fill table rows in the DOCX package.", ex);
        }
    }

    private static void ExpandRow(
        TableRow template, string collection,
        IReadOnlyList<IReadOnlyDictionary<string, string>> records)
    {
        var parent = template.Parent
                     ?? throw new DocumentConversionException("A template row had no parent table.");

        foreach (var record in records)
        {
            var clone = (TableRow)template.CloneNode(deep: true);
            Substitute(clone, collection, record);
            parent.InsertBefore(clone, template);
        }

        template.Remove();

        // Removing the now-empty table is a PRESENTATION choice, not a correctness fix. The design
        // assumed a w:tbl with no w:tr would be rejected; measured with OpenXmlValidator, a table
        // carrying tblPr and tblGrid but no rows validates clean. It is kept because an empty
        // one-cell frame left behind on a document whose list happened to be empty is worse than
        // rendering nothing, which is what "no records" means.
        if (parent is Table table && !table.ChildElements.OfType<TableRow>().Any())
            table.Remove();
    }

    private static void Substitute(
        TableRow clone, string collection, IReadOnlyDictionary<string, string> record)
    {
        var replacements = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var pair in record)
            replacements["{{" + collection + "." + pair.Key + "}}"] = pair.Value ?? string.Empty;

        // Deliberately the same walk ReplaceText uses, so text boxes inside a cell behave
        // identically in both methods rather than by accident.
        ReplaceIn(clone, replacements);

        ClearUnmatched(clone, collection);
    }

    /// <summary>
    /// Blanks any placeholder for this collection the record had no key for. A half-filled document
    /// showing <c>{{item.Missing}}</c> to an end user is worse than an empty cell, and the keys to
    /// clear are only knowable after reading the document.
    /// </summary>
    private static void ClearUnmatched(TableRow clone, string collection)
    {
        var marker = "{{" + collection + ".";
        var leftovers = new Dictionary<string, string>(StringComparer.Ordinal);

        foreach (var paragraph in clone.Descendants<Paragraph>())
        {
            var merged = paragraph.InnerText;
            var at = merged.IndexOf(marker, StringComparison.Ordinal);
            while (at >= 0)
            {
                var close = merged.IndexOf("}}", at, StringComparison.Ordinal);
                if (close < 0) break;
                leftovers[merged[at..(close + 2)]] = string.Empty;
                at = merged.IndexOf(marker, close, StringComparison.Ordinal);
            }
        }

        if (leftovers.Count > 0) ReplaceIn(clone, leftovers);
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
