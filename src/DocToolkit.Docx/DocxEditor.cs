using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using OfficeIMOWordWordDocument = OfficeIMO.Word.WordDocument;
namespace DocToolkit;

/// <summary>Creates, reads and edits Word (.docx) documents.</summary>
public static class DocxEditor
{
    /// <summary>
    /// Creates a document from <paramref name="blocks"/>.
    ///
    /// A DOCX can also be produced by converting HTML with <c>HtmlToDocxConverter</c>. This
    /// exists for the case where the content comes from data rather than from markup: there is no
    /// HTML to escape, so a value containing <c>&lt;</c> cannot corrupt the document's structure,
    /// and the same blocks produce the same CONTENT on every machine — nothing here consults the
    /// current culture. Not the same BYTES: the OpenXml SDK mints fresh relationship ids per
    /// package, so two calls with identical blocks in the same process differ. Do not build a cache
    /// key, a content hash or a golden-file test on the bytes.
    ///
    /// An empty sequence is valid and produces a valid empty document.
    /// </summary>
    /// <param name="blocks">The content, written in order.</param>
    /// <exception cref="ArgumentNullException"><paramref name="blocks"/> is null.</exception>
    /// <exception cref="ArgumentException">An element of <paramref name="blocks"/> is null.</exception>
    /// <exception cref="DocumentConversionException">The document could not be built.</exception>
    /// <remarks>
    /// The document is laid out on <see cref="PageSetup.A4"/>. Use
    /// <see cref="Create(IEnumerable{DocxBlock}, PageSetup)"/> for anything else.
    /// </remarks>
    /// <example>
    /// <code source="../../tests/DocToolkit.Tests/DocumentationExamples.cs" region="DocxCreate"/>
    /// </example>
    public static byte[] Create(IEnumerable<DocxBlock> blocks) => Create(blocks, PageSetup.A4);

    /// <summary>
    /// Builds a document from <paramref name="blocks"/>, laid out on <paramref name="page"/>. See
    /// <see cref="Create(IEnumerable{DocxBlock})"/> for the block semantics — this overload applies
    /// identical logic and differs only in the paper.
    /// </summary>
    /// <param name="blocks">The content, written in order.</param>
    /// <param name="page">The page size, orientation and margins.</param>
    /// <exception cref="ArgumentNullException"><paramref name="blocks"/> or <paramref name="page"/> is null.</exception>
    /// <exception cref="ArgumentException">An element of <paramref name="blocks"/> is null.</exception>
    /// <exception cref="DocumentConversionException">The document could not be built.</exception>
    public static byte[] Create(IEnumerable<DocxBlock> blocks, PageSetup page)
    {
        ArgumentNullException.ThrowIfNull(page);

        var materialised = ValidateBlocks(blocks);
        using var ms = DocxDocumentWriter.Write(materialised, page);
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a document from <paramref name="blocks"/> and writes it to
    /// <paramref name="destination"/>. See <see cref="Create(IEnumerable{DocxBlock})"/> for the block semantics — this
    /// overload applies identical logic, writing to <paramref name="destination"/> instead of
    /// returning an array.
    ///
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought — it belongs to the caller, and may be write-only and
    /// forward-only, such as an HTTP response body.
    /// </summary>
    /// <param name="blocks">The content, written in order.</param>
    /// <param name="destination">The stream the document is written to.</param>
    /// <param name="ct">Cancels the build and the write to <paramref name="destination"/>.</param>
    /// <exception cref="ArgumentNullException"><paramref name="blocks"/> or <paramref name="destination"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// An element of <paramref name="blocks"/> is null, or <paramref name="destination"/> is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be built or written.</exception>
    /// <remarks>
    /// The document is laid out on <see cref="PageSetup.A4"/>. Use
    /// <see cref="CreateAsync(IEnumerable{DocxBlock}, PageSetup, Stream, CancellationToken)"/> for
    /// anything else.
    /// </remarks>
    public static Task CreateAsync(
        IEnumerable<DocxBlock> blocks, Stream destination, CancellationToken ct = default) =>
        CreateAsync(blocks, PageSetup.A4, destination, ct);

    /// <summary>
    /// Builds a document from <paramref name="blocks"/>, laid out on <paramref name="page"/>, and
    /// writes it to <paramref name="destination"/>. See
    /// <see cref="Create(IEnumerable{DocxBlock}, PageSetup)"/> for the semantics.
    ///
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought — it belongs to the caller, and may be write-only and
    /// forward-only, such as an HTTP response body.
    /// </summary>
    /// <param name="blocks">The content, written in order.</param>
    /// <param name="page">The page size, orientation and margins.</param>
    /// <param name="destination">The stream the document is written to.</param>
    /// <param name="ct">Cancels the build and the write to <paramref name="destination"/>.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// An element of <paramref name="blocks"/> is null, or <paramref name="destination"/> is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be built or written.</exception>
    public static async Task CreateAsync(
        IEnumerable<DocxBlock> blocks, PageSetup page, Stream destination,
        CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(page);

        var materialised = ValidateBlocks(blocks);
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var ms = DocxDocumentWriter.Write(materialised, page);
        await StreamPipeline.EmitAsync(ms, destination, "Failed to create DOCX.", ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Builds a document from <paramref name="blocks"/> and writes it to
    /// <paramref name="outputPath"/>. See <see cref="Create(IEnumerable{DocxBlock})"/> for the block semantics.
    ///
    /// Named <c>CreateToFileAsync</c> rather than a third <c>CreateAsync</c> overload, matching
    /// <c>WorkbookEditor.CreateToFileAsync</c>:
    /// the distinct name keeps which kind of destination a call writes to visible at the call site,
    /// rather than resting on the argument type alone.
    ///
    /// The document is built completely before the output is opened. That ordering is the reason a
    /// failed build cannot truncate a file that was already there, and it is pinned by
    /// <c>FilePathOverloadTests</c> rather than left as a comment — it survives only as long as
    /// nobody rewrites this into a streaming write.
    /// </summary>
    /// <param name="blocks">The content, written in order.</param>
    /// <param name="outputPath">Where to write the document. Overwritten if it exists.</param>
    /// <param name="ct">Cancels the write to <paramref name="outputPath"/>.</param>
    /// <exception cref="ArgumentNullException"><paramref name="blocks"/> or <paramref name="outputPath"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="outputPath"/> is blank, or an element of <paramref name="blocks"/> is null.
    /// </exception>
    /// <exception cref="DirectoryNotFoundException"><paramref name="outputPath"/>'s directory does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be built.</exception>
    /// <remarks>
    /// The document is laid out on <see cref="PageSetup.A4"/>. Use
    /// <see cref="CreateToFileAsync(IEnumerable{DocxBlock}, PageSetup, string, CancellationToken)"/>
    /// for anything else.
    /// </remarks>
    public static Task CreateToFileAsync(
        IEnumerable<DocxBlock> blocks, string outputPath, CancellationToken ct = default) =>
        CreateToFileAsync(blocks, PageSetup.A4, outputPath, ct);

    /// <summary>
    /// Builds a document from <paramref name="blocks"/>, laid out on <paramref name="page"/>, and
    /// writes it to <paramref name="outputPath"/>. See
    /// <see cref="Create(IEnumerable{DocxBlock}, PageSetup)"/> for the semantics.
    ///
    /// The document is built completely before the output is opened, so a failed build cannot
    /// truncate a file that was already there.
    /// </summary>
    /// <param name="blocks">The content, written in order.</param>
    /// <param name="page">The page size, orientation and margins.</param>
    /// <param name="outputPath">Where to write the document. Overwritten if it exists.</param>
    /// <param name="ct">Cancels the write to <paramref name="outputPath"/>.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="outputPath"/> is blank, or an element of <paramref name="blocks"/> is null.
    /// </exception>
    /// <exception cref="DirectoryNotFoundException"><paramref name="outputPath"/>'s directory does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be built.</exception>
    public static async Task CreateToFileAsync(
        IEnumerable<DocxBlock> blocks, PageSetup page, string outputPath,
        CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var bytes = Create(blocks, page);
        await File.WriteAllBytesAsync(outputPath, bytes, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Materialises and null-checks up front, so a null block surfaces as the
    /// <see cref="ArgumentException"/> it is rather than as a <see cref="NullReferenceException"/>
    /// wrapped in a conversion failure. Mirrors <c>WorkbookEditor.ValidateRows</c>.
    /// </summary>
    private static List<DocxBlock> ValidateBlocks(IEnumerable<DocxBlock> blocks)
    {
        ArgumentNullException.ThrowIfNull(blocks);

        return blocks
            .Select((block, index) => block
                ?? throw new ArgumentException($"Block {index + 1} was null.", nameof(blocks)))
            .ToList();
    }

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
    /// <example>
    /// <code source="../../tests/DocToolkit.Tests/DocumentationExamples.cs" region="DocxReplaceText"/>
    /// </example>
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
            .DrainAsync(source, "DOCX content was empty.", nameof(source), "Failed to edit DOCX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        ReplaceTextCore(docx, replacements);

        await StreamPipeline.EmitAsync(docx, destination, "Failed to edit DOCX. See the inner exception for details.", ct).ConfigureAwait(false);
    }

    private static void ReplaceTextCore(MemoryStream ms, IReadOnlyDictionary<string, string> replacements)
    {
        try
        {
            using (var doc = WordprocessingDocument.Open(ms, true))
            {
                var main = doc.MainDocumentPart
                           ?? throw new DocumentConversionException("Document has no main part. This usually means the file is not really a .docx (for "
                           + "example it was renamed from another format) or the upload is corrupt.");
                var body = main.Document?.Body
                           ?? throw new DocumentConversionException("Document has no body. This usually means the file is not really a .docx (for "
                           + "example it was renamed from another format) or the upload is corrupt.");

                ReplaceIn(body, replacements);
                main.Document!.Save();

                // A placeholder in a header or footer used to come back unreplaced with no error
                // at all, which is a silent wrong answer for the "fill a template" use case.
                foreach (var part in main.HeaderParts.Where(p => p.Header is not null))
                {
                    ReplaceIn(part.Header!, replacements);
                    part.Header!.Save();
                }

                foreach (var part in main.FooterParts.Where(p => p.Footer is not null))
                {
                    ReplaceIn(part.Footer!, replacements);
                    part.Footer!.Save();
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
            throw new DocumentConversionException("Failed to edit DOCX. See the inner exception for details.", ex);
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
            .DrainAsync(source, "DOCX content was empty.", nameof(source), "Failed to read DOCX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        return ExtractTextCore(docx, includeHeadersAndFooters);
    }

    private static string ExtractTextCore(Stream ms, bool includeHeadersAndFooters)
    {
        try
        {
            using var doc = WordprocessingDocument.Open(ms, false);

            var main = doc.MainDocumentPart;
            var body = main?.Document?.Body;
            var bodyText = body is null ? string.Empty : BlockText(body);
            if (!includeHeadersAndFooters || main is null) return bodyText;

            var sb = new StringBuilder(bodyText);
            foreach (var text in main.HeaderParts.Select(p => p.Header is null ? null : BlockText(p.Header))
                                     .Concat(main.FooterParts.Select(p => p.Footer is null ? null : BlockText(p.Footer)))
                                     .Where(t => !string.IsNullOrEmpty(t)))
            {
                if (sb.Length > 0) sb.Append('\n');
                sb.Append(text);
            }

            return sb.ToString();
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read DOCX. See the inner exception for details.", ex);
        }
    }

    /// <summary>
    /// The text of a block container — the body, a header, a footer, or a table cell — with block
    /// boundaries preserved: <c>\n</c> between blocks, <c>\t</c> between the cells of a row.
    ///
    /// This exists because <c>InnerText</c> concatenates every descendant text node with no
    /// separator whatsoever, so a heading "Title" followed by a paragraph "Body text." came back
    /// as the single token <c>TitleBody text.</c> — indistinguishable to a substring search, wrong
    /// for anything that tokenises, indexes or diffs. Fixed 2026-08-10; see the A26 backlog row.
    ///
    /// <b>Uses <c>Elements&lt;T&gt;()</c>, never <c>Descendants&lt;T&gt;()</c>.</b> This is the
    /// same trap <see cref="TableRowFinder"/> documents: <c>Descendants</c> would yield the rows of
    /// a table nested inside a cell as though they belonged to the outer table, flattening exactly
    /// the structure this method exists to preserve. Nesting is handled by recursion instead, so
    /// each level keeps its own separators.
    /// </summary>
    private static string BlockText(OpenXmlElement container)
    {
        var blocks = new List<string>();

        foreach (var child in container.Elements())
        {
            switch (child)
            {
                case Paragraph paragraph:
                    // InnerText, deliberately, rather than walking runs. A paragraph's runs are
                    // fragments of ONE visible line — Word splits them on formatting, rsid and
                    // revision boundaries — so separating them would reintroduce the run-splitting
                    // bug that RunTextSplicer exists to solve, only in the reading direction.
                    blocks.Add(paragraph.InnerText);
                    break;
                case Table table:
                    blocks.Add(TableText(table));
                    break;
                case SdtBlock control:
                    // A content control is a THIRD block-level child, beside Paragraph and Table -
                    // and leaving it out of this switch is why every one of them was silently
                    // absent from the text. Recursing keeps the Elements-not-Descendants rule this
                    // method's summary depends on, and handles a nested control by construction.
                    if (control.SdtContentBlock is { } content)
                        blocks.Add(BlockText(content));
                    break;
            }
        }

        return string.Join("\n", blocks);
    }

    /// <summary>
    /// A table as text: cells joined by <c>\t</c> and rows by <c>\n</c>, which is what Word's own
    /// "save as plain text" writes. Cells recurse through <see cref="BlockText"/> so a cell holding
    /// several paragraphs, or a nested table, keeps its own structure.
    /// </summary>
    private static string TableText(Table table) =>
        string.Join("\n", ContentControls.Rows(table)
            .Select(row => string.Join("\t", ContentControls.Cells(row).Select(BlockText))));

    /// <summary>
    /// How many tables the document body holds.
    /// </summary>
    /// <remarks>
    /// Top-level tables only. A table nested inside a cell is part of that cell's text rather than
    /// an entry of its own, so this count and the indexes it bounds stay stable.
    ///
    /// <b>A table wrapped in a content control counts.</b> Word puts a <c>w:sdt</c> around
    /// content the author marked up, and it used to hide the table inside from this count
    /// entirely — so a document whose only table was wrapped reported <b>0</b>. It reports 1.
    ///
    /// <b>That moves indexes, deliberately.</b> Where a wrapped table precedes an ordinary one,
    /// index 0 used to return the ordinary one — the table that is physically second. Indexes are
    /// now positional over every table a reader can see, which is what
    /// <see cref="ExtractText(byte[])"/> has always reported.
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The package could not be read.</exception>
    /// <example>
    /// <code source="../../tests/DocToolkit.Tests/DocumentationExamples.cs" region="DocxReadTable"/>
    /// </example>
    public static int TableCount(byte[] docx)
    {
        ArgumentNullException.ThrowIfNull(docx);
        if (docx.Length == 0)
            throw new ArgumentException("DOCX content was empty.", nameof(docx));

        using var ms = new MemoryStream(docx, writable: false);
        return TableCountCore(ms);
    }

    /// <summary>
    /// The table at <paramref name="index"/>, as rows of cell text.
    /// </summary>
    /// <param name="docx">The .docx content to read.</param>
    /// <param name="index">
    /// <b>0-based</b>, indexing what <see cref="TableCount(byte[])"/> reports — deliberately unlike
    /// <c>PdfEditor.ExtractPages</c>, whose <c>firstPage</c> is 1-based because that is how a reader
    /// numbers pages. A table has no such reader-facing numbering, and the rows and cells this
    /// returns are 0-based, so a 1-based selector here would be the odd one out.
    /// </param>
    /// <remarks>
    /// Cell text is produced the same way <see cref="ExtractText(byte[])"/> produces it, so a cell
    /// holding several paragraphs is separated by newlines and a nested table keeps its own
    /// structure.
    ///
    /// <b>Rows are returned with the shape they have.</b> A horizontally merged cell means a row
    /// genuinely holds fewer cells than its neighbours; padding the grid to a rectangle would invent
    /// cells that are not in the document.
    ///
    /// <b>Content controls are transparent, at every level.</b> A table, a row or a cell wrapped in
    /// a <c>w:sdt</c> is read like any other. A wrapped row used to vanish from an otherwise
    /// correct table, and a wrapped cell used to shift every cell beside it — both of which look
    /// like data rather than an error. See <see cref="TableCount(byte[])"/> for what this does to
    /// indexes.
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="index"/> is negative, or at or beyond <see cref="TableCount(byte[])"/>.
    /// </exception>
    /// <exception cref="DocumentConversionException">The package could not be read.</exception>
    /// <example>
    /// <code source="../../tests/DocToolkit.Tests/DocumentationExamples.cs" region="DocxReadTable"/>
    /// </example>
    public static IReadOnlyList<IReadOnlyList<string>> ReadTable(byte[] docx, int index)
    {
        ArgumentNullException.ThrowIfNull(docx);
        if (docx.Length == 0)
            throw new ArgumentException("DOCX content was empty.", nameof(docx));
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 0);

        using var ms = new MemoryStream(docx, writable: false);
        return ReadTableCore(ms, index);
    }

    /// <inheritdoc cref="TableCount(byte[])"/>
    /// <remarks>
    /// <paramref name="source"/> is read to its end; it is not disposed, closed or sought, and does
    /// not have to be seekable.
    /// </remarks>
    public static async Task<int> TableCountAsync(Stream source, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, "DOCX content was empty.", nameof(source), "Failed to read DOCX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        return TableCountCore(docx);
    }

    /// <inheritdoc cref="TableCount(byte[])"/>
    public static async Task<int> TableCountAsync(string path, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        return TableCount(await FilePipeline.ReadAsync(path, nameof(path), ct).ConfigureAwait(false));
    }

    /// <inheritdoc cref="ReadTable(byte[], int)"/>
    /// <remarks>
    /// <paramref name="source"/> is read to its end; it is not disposed, closed or sought, and does
    /// not have to be seekable.
    /// </remarks>
    public static async Task<IReadOnlyList<IReadOnlyList<string>>> ReadTableAsync(
        Stream source, int index, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 0);
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, "DOCX content was empty.", nameof(source), "Failed to read DOCX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        return ReadTableCore(docx, index);
    }

    /// <inheritdoc cref="ReadTable(byte[], int)"/>
    public static async Task<IReadOnlyList<IReadOnlyList<string>>> ReadTableAsync(
        string path, int index, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        return ReadTable(await FilePipeline.ReadAsync(path, nameof(path), ct).ConfigureAwait(false), index);
    }

    private static int TableCountCore(Stream ms)
    {
        try
        {
            using var doc = WordprocessingDocument.Open(ms, false);

            // Elements, not Descendants: a nested table belongs to its cell's text, and counting it
            // separately would report one table twice and make indexes shift under a caller.
            var body = doc.MainDocumentPart?.Document?.Body;
            return body is null ? 0 : ContentControls.Tables(body).Count();
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read DOCX. See the inner exception for details.", ex);
        }
    }

    private static IReadOnlyList<IReadOnlyList<string>> ReadTableCore(Stream ms, int index)
    {
        try
        {
            using var doc = WordprocessingDocument.Open(ms, false);

            var body = doc.MainDocumentPart?.Document?.Body;
            var tables = body is null ? [] : ContentControls.Tables(body).ToList();

            if (index >= tables.Count)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(index), index,
                    $"The document has {tables.Count} table(s).");
            }

            return ContentControls.Rows(tables[index])
                .Select(row => (IReadOnlyList<string>)ContentControls.Cells(row)
                                                                     .Select(BlockText)
                                                                     .ToList())
                .ToList();
        }
        // ArgumentOutOfRangeException is thrown INSIDE this try and must escape as itself. The
        // sibling *Core methods only exclude DocumentConversionException, which would be wrong
        // here: a caller passing a bad index would receive "Failed to read DOCX" instead.
        catch (Exception ex)
            when (ex is not DocumentConversionException and not ArgumentOutOfRangeException)
        {
            throw new DocumentConversionException("Failed to read DOCX. See the inner exception for details.", ex);
        }
    }

    /// <summary>
    /// Expands a table row once per record, so a template can render a variable-length list such as
    /// invoice line items.
    ///
    /// A template row wrapped in a content control — or sitting in a wrapped table — is found and
    /// expanded like any other, and its clones stay inside the control the author put it in.
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

    /// <summary>
    /// Reads a .docx from <paramref name="source"/>, expands the template row once per record, and
    /// writes the result to <paramref name="destination"/>. See
    /// <see cref="FillRows"/> for what counts as a template row and how formatting survives — this
    /// overload applies the identical logic via streams instead of a byte array.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought, and neither has to be seekable, so
    /// both may be sockets, files or HTTP message bodies.
    /// </summary>
    /// <param name="source">The stream the .docx package is read from.</param>
    /// <param name="collection">The placeholder prefix marking the template row, without braces.</param>
    /// <param name="rows">One dictionary per record, keyed by bare field name.</param>
    /// <param name="destination">The stream the edited .docx package is written to.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, <paramref name="destination"/> is
    /// not writable, or <paramref name="collection"/> is blank.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The package could not be opened or edited, or no template row was found.
    /// </exception>
    public static async Task FillRowsAsync(
        Stream source, string collection, IEnumerable<IReadOnlyDictionary<string, string>> rows,
        Stream destination, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(collection);
        ArgumentNullException.ThrowIfNull(rows);
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        if (string.IsNullOrWhiteSpace(collection))
            throw new ArgumentException("Collection name was blank.", nameof(collection));

        using var buffer = await StreamPipeline
            .DrainAsync(source, "DOCX content was empty.", nameof(source), "Failed to fill table rows in the DOCX package. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        FillRowsCore(buffer, collection, rows);

        await StreamPipeline
            .EmitAsync(buffer, destination, "Failed to fill table rows in the DOCX package. See the inner exception for details.", ct)
            .ConfigureAwait(false);
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
                           ?? throw new DocumentConversionException("Document has no main part. This usually means the file is not really a .docx (for "
                           + "example it was renamed from another format) or the upload is corrupt.");
                var body = main.Document?.Body
                           ?? throw new DocumentConversionException("Document has no body. This usually means the file is not really a .docx (for "
                           + "example it was renamed from another format) or the upload is corrupt.");

                var templates = TableRowFinder.Find(body, marker);
                if (templates.Count == 0)
                {
                    throw new DocumentConversionException(
                        $"No table row containing '{marker}' was found, so there was nothing to "
                        + "fill. The marker must appear inside a table cell, not just anywhere in "
                        + "the document.");
                }

                foreach (var template in templates)
                    ExpandRow(template, collection, records);

                main.Document!.Save();
            }

            ms.Position = 0;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to fill table rows in the DOCX package. See the inner exception for details.", ex);
        }
    }

    private static void ExpandRow(
        TableRow template, string collection,
        IReadOnlyList<IReadOnlyDictionary<string, string>> records)
    {
        // The row's immediate parent is where clones go - which for a wrapped row is the
        // control's content, not the table. Inserting into the table instead would move every
        // generated row out of the control the author put it in.
        var parent = template.Parent
                     ?? throw new DocumentConversionException(
                         "A template row had no parent element. FillRows expands rows that live in "
                         + "a table, directly or inside a content control — put the marker inside a "
                         + "table row.");

        // Resolved BEFORE the template is detached. Ancestors() of a removed node is empty, so
        // computing this after the Remove() below finds nothing and the table is never cleaned up -
        // measured, and it is why the first attempt at this fix still failed its own test.
        var owner = template.Ancestors<Table>().FirstOrDefault();

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
        // The OWNING table, not the immediate parent. For a row inside a content control the
        // parent is the control's content, so `parent is Table` was silently false and the empty
        // frame stayed on the page - making DocxEditor's own shipped documentation false on the
        // very path this change opened. Ancestors, because the control may be nested.
        //
        // Emptiness is asked through ContentControls for the same reason: a table whose only
        // remaining row is wrapped is NOT empty, and a direct-child count would call it empty and
        // delete a row the caller can see.
        if (owner is not null && !ContentControls.Rows(owner).Any())
            owner.Remove();
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

    /// <summary>
    /// Replaces every occurrence of <paramref name="placeholder"/> with <paramref name="image"/>,
    /// inline, across the body, headers, footers, footnotes and endnotes.
    ///
    /// Only the matched text goes: text sharing a run with the placeholder keeps its place and its
    /// formatting, so <c>Signed: {{sig}} (authorised)</c> becomes <c>Signed: </c>, the image, then
    /// <c> (authorised)</c>.
    ///
    /// <paramref name="placeholder"/> is the literal text including braces, like
    /// <see cref="ReplaceText(byte[], IReadOnlyDictionary{string, string})"/> — and unlike
    /// <see cref="FillRows"/>, whose keys are bare field names only because the collection name is
    /// already an argument there.
    ///
    /// Size is in points. Omit both and the image's intrinsic size is used, read from its own header
    /// at 96 DPI. Give one and the other scales to preserve the aspect ratio. Give both and the
    /// image is stretched to fit — distortion is the caller's choice, not an error.
    ///
    /// PNG and JPEG only, detected from the image's magic bytes rather than any filename.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any of the three required arguments is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="docx"/> or <paramref name="image"/> is empty, or <paramref name="placeholder"/>
    /// is blank.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// A supplied size is zero or negative, or the resulting size is larger than a drawing extent can
    /// hold (2,147,483,647 EMU per side — about 2,348 inches). The upper bound also applies when the
    /// side that overflows is the one DERIVED from the aspect ratio rather than the one supplied.
    /// </exception>
    /// <exception cref="DocumentConversionException">
    /// The image is neither PNG nor JPEG, the package could not be edited, or
    /// <paramref name="placeholder"/> does not appear anywhere — a call matching nothing is a bug in
    /// the call or the template, not a no-op.
    /// </exception>
    public static byte[] ReplaceImage(
        byte[] docx, string placeholder, byte[] image,
        double? widthPoints = null, double? heightPoints = null)
    {
        ArgumentNullException.ThrowIfNull(docx);
        ArgumentNullException.ThrowIfNull(placeholder);
        ArgumentNullException.ThrowIfNull(image);
        if (docx.Length == 0) throw new ArgumentException("DOCX content was empty.", nameof(docx));
        if (image.Length == 0) throw new ArgumentException("Image content was empty.", nameof(image));
        if (string.IsNullOrWhiteSpace(placeholder))
            throw new ArgumentException("Placeholder was blank.", nameof(placeholder));

        using var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        ms.Position = 0;

        ReplaceImageCore(ms, placeholder, image, widthPoints, heightPoints);
        return ms.ToArray();
    }

    /// <summary>
    /// Reads a .docx from <paramref name="source"/>, replaces every occurrence of
    /// <paramref name="placeholder"/> with <paramref name="image"/>, and writes the result to
    /// <paramref name="destination"/>. See <see cref="ReplaceImage"/> for what is matched and how it
    /// is sized — this overload applies the identical logic via streams instead of a byte array.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought, and neither has to be seekable, so
    /// both may be sockets, files or HTTP message bodies.
    /// </summary>
    /// <param name="source">The stream the .docx package is read from.</param>
    /// <param name="placeholder">The literal placeholder text, braces included.</param>
    /// <param name="image">PNG or JPEG bytes, identified by their magic bytes.</param>
    /// <param name="destination">The stream the edited .docx package is written to.</param>
    /// <param name="widthPoints">Width in points, or null to derive it.</param>
    /// <param name="heightPoints">Height in points, or null to derive it.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, <paramref name="destination"/> is
    /// not writable, <paramref name="image"/> is empty, or <paramref name="placeholder"/> is blank.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// A supplied size is zero or negative, or the resulting size is larger than a drawing extent can
    /// hold (2,147,483,647 EMU per side — about 2,348 inches). The upper bound also applies when the
    /// side that overflows is the one DERIVED from the aspect ratio rather than the one supplied.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The image is neither PNG nor JPEG, the package could not be edited, or the placeholder was
    /// not found.
    /// </exception>
    public static async Task ReplaceImageAsync(
        Stream source, string placeholder, byte[] image, Stream destination,
        double? widthPoints = null, double? heightPoints = null, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(placeholder);
        ArgumentNullException.ThrowIfNull(image);
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        if (image.Length == 0) throw new ArgumentException("Image content was empty.", nameof(image));
        if (string.IsNullOrWhiteSpace(placeholder))
            throw new ArgumentException("Placeholder was blank.", nameof(placeholder));

        using var buffer = await StreamPipeline
            .DrainAsync(source, "DOCX content was empty.", nameof(source), "Failed to insert an image into the DOCX package. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        ReplaceImageCore(buffer, placeholder, image, widthPoints, heightPoints);

        await StreamPipeline
            .EmitAsync(buffer, destination, "Failed to insert an image into the DOCX package. See the inner exception for details.", ct)
            .ConfigureAwait(false);
    }

    /// <summary>The one real implementation; every overload calls it so they cannot drift apart.</summary>
    private static void ReplaceImageCore(
        MemoryStream ms, string placeholder, byte[] image, double? widthPoints, double? heightPoints)
    {
        var info = ImageInspector.Inspect(image);
        var (widthEmu, heightEmu) = ImageInspector.Resolve(info, widthPoints, heightPoints);
        var name = placeholder.Trim().Trim('{', '}').Trim();

        try
        {
            using (var doc = WordprocessingDocument.Open(ms, true))
            {
                var main = doc.MainDocumentPart
                           ?? throw new DocumentConversionException("Document has no main part. This usually means the file is not really a .docx (for "
                           + "example it was renamed from another format) or the upload is corrupt.");
                var body = main.Document?.Body
                           ?? throw new DocumentConversionException("Document has no body. This usually means the file is not really a .docx (for "
                           + "example it was renamed from another format) or the upload is corrupt.");

                // Unique across the WHOLE document: a duplicate wp:docPr id makes Word declare the
                // file corrupt and offer to repair it, so start above whatever is already there.
                var nextId = NextDrawingId(main);
                var replaced = 0;

                replaced += InsertImagesIn(main, body, placeholder, image, info, widthEmu, heightEmu, name, ref nextId);
                main.Document!.Save();

                foreach (var part in main.HeaderParts.Where(p => p.Header is not null))
                {
                    replaced += InsertImagesIn(part, part.Header!, placeholder, image, info, widthEmu, heightEmu, name, ref nextId);
                    part.Header!.Save();
                }

                foreach (var part in main.FooterParts.Where(p => p.Footer is not null))
                {
                    replaced += InsertImagesIn(part, part.Footer!, placeholder, image, info, widthEmu, heightEmu, name, ref nextId);
                    part.Footer!.Save();
                }

                if (main.FootnotesPart?.Footnotes is { } footnotes)
                {
                    replaced += InsertImagesIn(main.FootnotesPart, footnotes, placeholder, image, info, widthEmu, heightEmu, name, ref nextId);
                    footnotes.Save();
                }

                if (main.EndnotesPart?.Endnotes is { } endnotes)
                {
                    replaced += InsertImagesIn(main.EndnotesPart, endnotes, placeholder, image, info, widthEmu, heightEmu, name, ref nextId);
                    endnotes.Save();
                }

                if (replaced == 0)
                {
                    throw new DocumentConversionException(
                        $"The placeholder '{placeholder}' was not found, so there was nothing to "
                        + "replace. Check the placeholder text, braces included, matches the "
                        + "document exactly.");
                }
            }

            ms.Position = 0;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to insert an image into the DOCX package. See the inner exception for details.", ex);
        }
    }

    /// <summary>One above the highest wp:docPr id anywhere in the package.</summary>
    private static uint NextDrawingId(MainDocumentPart main)
    {
        var highest = AllRoots(main)
            .SelectMany(root => root.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>())
            .Select(properties => properties.Id?.Value)
            .Where(id => id.HasValue)
            .Select(id => id!.Value)
            .DefaultIfEmpty(0U)
            .Max();

        return highest + 1;

        static IEnumerable<OpenXmlElement> AllRoots(MainDocumentPart part)
        {
            if (part.Document is not null) yield return part.Document;
            foreach (var header in part.HeaderParts.Where(h => h.Header is not null))
                yield return header.Header!;
            foreach (var footer in part.FooterParts.Where(f => f.Footer is not null))
                yield return footer.Footer!;
            if (part.FootnotesPart?.Footnotes is { } footnotes) yield return footnotes;
            if (part.EndnotesPart?.Endnotes is { } endnotes) yield return endnotes;
        }
    }

    private static int InsertImagesIn(
        OpenXmlPartContainer owner, OpenXmlElement root, string placeholder, byte[] image,
        ImageInfo info, long widthEmu, long heightEmu, string name, ref uint nextId)
    {
        var inserted = 0;

        // ONE image part per owner, created on first use, shared by every occurrence in it.
        //
        // This used to be added inside the per-offset loop, so the same bytes were embedded once
        // per placeholder occurrence. Measured 2026-08-20 with a 40 KB image: one occurrence gave
        // one media part and a 41,983-byte package; THREE gave media/image.png, media/image2.png
        // and media/image3.png, three identical copies, and 122,483 bytes. It grows linearly with
        // occurrences, and a letterhead logo repeated across a body, a header and a footer is the
        // ordinary case rather than a contrived one.
        //
        // PER OWNER, not once per document, and that is the constraint that decides where this
        // line goes. An image part must belong to the container that owns the paragraph - a
        // header's image added to the main document part yields a relationship id that resolves in
        // the wrong scope, and Word then opens the file and silently shows nothing. So callers hand
        // this method one owner at a time and each gets its own part; what is removed is the
        // duplication WITHIN an owner, not the separation BETWEEN them.
        //
        // Lazy, so a root containing no occurrence does not gain an orphan part.
        string? relationshipId = null;

        foreach (var paragraph in root.Descendants<Paragraph>().ToList())
        {
            // Same scoping as ReplaceInParagraph: only the text this paragraph directly owns, so a
            // text box nested in one of its runs is visited on its own rather than folded in here.
            var texts = paragraph.Descendants<Text>()
                                 .Where(t => t.Ancestors<Paragraph>().FirstOrDefault() == paragraph)
                                 .ToList();
            if (texts.Count == 0) continue;

            var merged = string.Concat(texts.Select(t => t.Text));

            var offsets = new List<int>();
            for (var at = merged.IndexOf(placeholder, StringComparison.Ordinal);
                 at >= 0;
                 at = merged.IndexOf(placeholder, at + placeholder.Length, StringComparison.Ordinal))
            {
                offsets.Add(at);
            }

            // Right to left, so the offsets of earlier matches stay valid as later ones are spliced.
            for (var i = offsets.Count - 1; i >= 0; i--)
            {
                relationshipId ??= AddImagePart(owner, image, info);
                // The placeholder-derived name doubles as the alt text here, deliberately and as
                // shipped: "{{logo}}" gives "logo", which is a genuine if terse description. That is
                // NOT true of the create path, whose names are generated ("Image 1"), so it passes
                // real alt text or none. Do not "unify" these - they differ because the inputs do.
                var drawing = DrawingFactory.InlineImage(
                    relationshipId, name, nextId++, widthEmu, heightEmu, description: name);
                SpliceElementIn(texts, offsets[i], placeholder.Length, new Run(drawing));
                inserted++;
            }
        }

        return inserted;
    }

    /// <summary>
    /// Adds the image bytes to <paramref name="owner"/> and returns its relationship id.
    ///
    /// The part must belong to the container that owns the paragraph. A header's image added to the
    /// main document part yields a relationship id that resolves in the wrong scope: Word opens the
    /// file and simply shows nothing where the image should be.
    /// </summary>
    private static string AddImagePart(OpenXmlPartContainer owner, byte[] image, ImageInfo info)
    {
        var part = owner.AddNewPart<ImagePart>(info.ContentType);
        using (var stream = part.GetStream(FileMode.Create))
        {
            stream.Write(image, 0, image.Length);
        }

        return owner.GetIdOfPart(part);
    }

    /// <summary>
    /// Removes <paramref name="length"/> characters at <paramref name="start"/> from the
    /// concatenation of <paramref name="texts"/> and puts <paramref name="replacement"/> there
    /// instead.
    ///
    /// This cannot use <see cref="RunTextSplicer"/>: that maps match offsets back onto runs and
    /// writes <i>text</i>, whereas this has to remove a span and insert an <i>element</i> at that
    /// position. Same principle — never touch a run the match does not overlap — different
    /// mechanism. Shared by every caller that inserts a non-text element at a placeholder: an
    /// image (<see cref="ReplaceImage(byte[], string, byte[], double?, double?)"/>), a footnote or
    /// endnote reference (AddFootnote, AddEndnote).
    /// </summary>
    private static void SpliceElementIn(List<Text> texts, int start, int length, Run replacement)
    {
        var end = start + length;
        var position = 0;
        Run? anchor = null;
        var suffix = string.Empty;

        foreach (var node in texts)
        {
            var nodeStart = position;
            var nodeEnd = position + node.Text.Length;
            position = nodeEnd;

            if (nodeEnd <= start || nodeStart >= end) continue;   // untouched by this match

            var keepBefore = start > nodeStart ? node.Text[..(start - nodeStart)] : string.Empty;
            var keepAfter = end < nodeEnd ? node.Text[(end - nodeStart)..] : string.Empty;

            if (anchor is null)
            {
                node.Text = keepBefore;
                anchor = node.Ancestors<Run>().FirstOrDefault();
                suffix = keepAfter;
            }
            else
            {
                node.Text = keepAfter;
            }
        }

        if (anchor is null) return;

        anchor.InsertAfterSelf(replacement);

        // A match wholly inside one run leaves a tail that needs a run of its own after the
        // inserted element.
        if (suffix.Length > 0)
        {
            replacement.InsertAfterSelf(new Run(
                new Text(suffix) { Space = SpaceProcessingModeValues.Preserve }));
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

    /// <summary>
    /// Reads a .docx from <paramref name="inputPath"/>, substitutes placeholders, and writes the
    /// result to <paramref name="outputPath"/>. The two may be the same file: the updated bytes are
    /// computed in full before <paramref name="outputPath"/> is opened, so a document that fails to
    /// process — cannot be read, or cannot be edited — leaves <paramref name="outputPath"/>
    /// untouched. That guarantee does not extend to a failure during the write itself: a full disk,
    /// a cancellation, or the process dying mid-write can still leave a partial file, so in-place
    /// editing of an irreplaceable document is not crash-safe.
    /// </summary>
    /// <param name="inputPath">The .docx to read.</param>
    /// <param name="outputPath">Where to write the result. Overwritten if it exists.</param>
    /// <param name="replacements">Placeholder to replacement text.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    /// <exception cref="ArgumentNullException">A path or <paramref name="replacements"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// A path is blank, or the file at <paramref name="inputPath"/> is empty.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="inputPath"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException">
    /// <paramref name="inputPath"/>'s or <paramref name="outputPath"/>'s directory does not exist.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be processed.</exception>
    public static async Task ReplaceTextAsync(
        string inputPath, string outputPath,
        IReadOnlyDictionary<string, string> replacements, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        ArgumentNullException.ThrowIfNull(replacements);

        var bytes = await FilePipeline.ReadAsync(inputPath, nameof(inputPath), ct).ConfigureAwait(false);
        var result = ReplaceText(bytes, replacements);
        await File.WriteAllBytesAsync(outputPath, result, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads a .docx from <paramref name="inputPath"/>, expands one table row per record, and
    /// writes the result to <paramref name="outputPath"/>. The two may be the same file: the
    /// updated bytes are computed in full before <paramref name="outputPath"/> is opened, so a
    /// document that fails to process — cannot be read, or cannot be edited — leaves
    /// <paramref name="outputPath"/> untouched. That guarantee does not extend to a failure during
    /// the write itself: a full disk, a cancellation, or the process dying mid-write can still leave
    /// a partial file, so in-place editing of an irreplaceable document is not crash-safe.
    /// </summary>
    /// <param name="inputPath">The .docx to read.</param>
    /// <param name="outputPath">Where to write the result. Overwritten if it exists.</param>
    /// <param name="collection">The collection name used in the row's placeholders.</param>
    /// <param name="rows">One dictionary per record.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    /// <exception cref="ArgumentNullException">A path, <paramref name="collection"/> or <paramref name="rows"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// A path is blank, or the file at <paramref name="inputPath"/> is empty.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="inputPath"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException">
    /// <paramref name="inputPath"/>'s or <paramref name="outputPath"/>'s directory does not exist.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be processed.</exception>
    public static async Task FillRowsAsync(
        string inputPath, string outputPath, string collection,
        IEnumerable<IReadOnlyDictionary<string, string>> rows, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var bytes = await FilePipeline.ReadAsync(inputPath, nameof(inputPath), ct).ConfigureAwait(false);
        var result = FillRows(bytes, collection, rows);
        await File.WriteAllBytesAsync(outputPath, result, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads a .docx from <paramref name="inputPath"/>, replaces an image placeholder, and writes
    /// the result to <paramref name="outputPath"/>. The two may be the same file: the updated bytes
    /// are computed in full before <paramref name="outputPath"/> is opened, so a document that fails
    /// to process — cannot be read, or cannot be edited — leaves <paramref name="outputPath"/>
    /// untouched. That guarantee does not extend to a failure during the write itself: a full disk,
    /// a cancellation, or the process dying mid-write can still leave a partial file, so in-place
    /// editing of an irreplaceable document is not crash-safe.
    /// </summary>
    /// <param name="inputPath">The .docx to read.</param>
    /// <param name="outputPath">Where to write the result. Overwritten if it exists.</param>
    /// <param name="placeholder">The placeholder text to replace.</param>
    /// <param name="image">PNG or JPEG bytes. The format is decided by the bytes, never a filename.</param>
    /// <param name="widthPoints">Width in points. Give one dimension and the other scales.</param>
    /// <param name="heightPoints">Height in points.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    /// <exception cref="ArgumentNullException">A path, <paramref name="placeholder"/> or <paramref name="image"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// A path is blank, or the file at <paramref name="inputPath"/> is empty.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="inputPath"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException">
    /// <paramref name="inputPath"/>'s or <paramref name="outputPath"/>'s directory does not exist.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be processed.</exception>
    public static async Task ReplaceImageAsync(
        string inputPath, string outputPath, string placeholder, byte[] image,
        double? widthPoints = null, double? heightPoints = null, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var bytes = await FilePipeline.ReadAsync(inputPath, nameof(inputPath), ct).ConfigureAwait(false);
        var result = ReplaceImage(bytes, placeholder, image, widthPoints, heightPoints);
        await File.WriteAllBytesAsync(outputPath, result, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Adds a footnote at every occurrence of <paramref name="placeholder"/>, inline, across the
    /// document body.
    ///
    /// Only the matched text goes: text sharing a run with the placeholder keeps its place and its
    /// formatting, so <c>See the note{{note}} here.</c> becomes <c>See the note</c>, the footnote
    /// reference, then <c> here.</c> Each occurrence gets its own footnote entry, all carrying
    /// <paramref name="footnoteText"/>.
    ///
    /// <paramref name="placeholder"/> is the literal text including braces, like
    /// <see cref="ReplaceImage(byte[], string, byte[], double?, double?)"/>. Only the document
    /// body is searched — headers, footers, and content already inside another footnote or
    /// endnote are not.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any of the three required arguments is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="docx"/> is empty, or <paramref name="placeholder"/> is blank.
    /// </exception>
    /// <exception cref="DocumentConversionException">
    /// The package could not be edited, or <paramref name="placeholder"/> does not appear in the
    /// body — a call matching nothing is a bug in the call or the template, not a no-op.
    /// </exception>
    public static byte[] AddFootnote(byte[] docx, string placeholder, string footnoteText)
    {
        ArgumentNullException.ThrowIfNull(docx);
        ArgumentNullException.ThrowIfNull(placeholder);
        ArgumentNullException.ThrowIfNull(footnoteText);
        if (docx.Length == 0) throw new ArgumentException("DOCX content was empty.", nameof(docx));
        if (string.IsNullOrWhiteSpace(placeholder))
            throw new ArgumentException("Placeholder was blank.", nameof(placeholder));

        using var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        ms.Position = 0;

        AddFootnoteCore(ms, placeholder, footnoteText);
        return ms.ToArray();
    }

    /// <summary>
    /// Reads a .docx from <paramref name="source"/>, adds a footnote at every occurrence of
    /// <paramref name="placeholder"/>, and writes the result to <paramref name="destination"/> —
    /// see <see cref="AddFootnote(byte[], string, string)"/> for exactly what is matched.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought, and neither has to be seekable.
    /// </summary>
    /// <param name="source">The stream the .docx package is read from.</param>
    /// <param name="placeholder">The literal placeholder text, braces included.</param>
    /// <param name="footnoteText">The footnote's own text.</param>
    /// <param name="destination">The stream the edited .docx package is written to.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, <paramref name="destination"/>
    /// is not writable, or <paramref name="placeholder"/> is blank.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The package could not be edited, or the placeholder was not found.
    /// </exception>
    public static async Task AddFootnoteAsync(
        Stream source, string placeholder, string footnoteText, Stream destination,
        CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(placeholder);
        ArgumentNullException.ThrowIfNull(footnoteText);
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        if (string.IsNullOrWhiteSpace(placeholder))
            throw new ArgumentException("Placeholder was blank.", nameof(placeholder));

        using var buffer = await StreamPipeline
            .DrainAsync(source, "DOCX content was empty.", nameof(source), "Failed to add a footnote to the DOCX package. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        AddFootnoteCore(buffer, placeholder, footnoteText);

        await StreamPipeline
            .EmitAsync(buffer, destination, "Failed to add a footnote to the DOCX package. See the inner exception for details.", ct)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Reads a .docx from <paramref name="inputPath"/>, adds a footnote at every occurrence of
    /// <paramref name="placeholder"/>, and writes the result to <paramref name="outputPath"/> —
    /// see <see cref="AddFootnote(byte[], string, string)"/> for exactly what is matched. The two
    /// paths may be the same file: the updated bytes are computed in full before
    /// <paramref name="outputPath"/> is opened.
    /// </summary>
    /// <param name="inputPath">The .docx to read.</param>
    /// <param name="outputPath">Where to write the result. Overwritten if it exists.</param>
    /// <param name="placeholder">The literal placeholder text, braces included.</param>
    /// <param name="footnoteText">The footnote's own text.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    /// <exception cref="ArgumentNullException">A path or <paramref name="placeholder"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// A path is blank, the file at <paramref name="inputPath"/> is empty, or
    /// <paramref name="placeholder"/> is blank.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="inputPath"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException">
    /// <paramref name="inputPath"/>'s or <paramref name="outputPath"/>'s directory does not exist.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The package could not be edited, or the placeholder was not found.
    /// </exception>
    public static async Task AddFootnoteAsync(
        string inputPath, string outputPath, string placeholder, string footnoteText,
        CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        ArgumentNullException.ThrowIfNull(placeholder);
        ArgumentNullException.ThrowIfNull(footnoteText);

        var bytes = await FilePipeline.ReadAsync(inputPath, nameof(inputPath), ct).ConfigureAwait(false);
        var result = AddFootnote(bytes, placeholder, footnoteText);
        await File.WriteAllBytesAsync(outputPath, result, ct).ConfigureAwait(false);
    }

    /// <summary>The one real implementation; every overload calls it so they cannot drift apart.</summary>
    private static void AddFootnoteCore(MemoryStream ms, string placeholder, string footnoteText)
    {
        try
        {
            using (var doc = WordprocessingDocument.Open(ms, true))
            {
                var main = doc.MainDocumentPart
                           ?? throw new DocumentConversionException("Document has no main part. This usually means the file is not really a .docx (for "
                           + "example it was renamed from another format) or the upload is corrupt.");
                var body = main.Document?.Body
                           ?? throw new DocumentConversionException("Document has no body. This usually means the file is not really a .docx (for "
                           + "example it was renamed from another format) or the upload is corrupt.");

                var inserted = InsertFootnoteReferencesIn(main, body, placeholder, footnoteText);

                if (inserted == 0)
                {
                    throw new DocumentConversionException(
                        $"The placeholder '{placeholder}' was not found, so there was no footnote to add. "
                        + "Check the placeholder text, braces included, matches the document exactly.");
                }

                main.Document!.Save();
            }

            ms.Position = 0;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to add a footnote to the DOCX package. See the inner exception for details.", ex);
        }
    }

    private static int InsertFootnoteReferencesIn(
        MainDocumentPart main, Body body, string placeholder, string footnoteText)
    {
        var inserted = 0;
        var nextId = NextFootnoteId(main);

        foreach (var paragraph in body.Descendants<Paragraph>().ToList())
        {
            var texts = paragraph.Descendants<Text>()
                                 .Where(t => t.Ancestors<Paragraph>().FirstOrDefault() == paragraph)
                                 .ToList();
            if (texts.Count == 0) continue;

            var merged = string.Concat(texts.Select(t => t.Text));

            var offsets = new List<int>();
            for (var at = merged.IndexOf(placeholder, StringComparison.Ordinal);
                 at >= 0;
                 at = merged.IndexOf(placeholder, at + placeholder.Length, StringComparison.Ordinal))
            {
                offsets.Add(at);
            }

            // Right to left, so the offsets of earlier matches stay valid as later ones are spliced.
            for (var i = offsets.Count - 1; i >= 0; i--)
            {
                var id = nextId++;
                AddFootnoteEntry(main, id, footnoteText);

                var referenceRun = new Run(
                    new RunProperties(new RunStyle { Val = "FootnoteReference" }),
                    new FootnoteReference { Id = id });
                SpliceElementIn(texts, offsets[i], placeholder.Length, referenceRun);
                inserted++;
            }
        }

        return inserted;
    }

    /// <summary>
    /// One above the highest existing footnote id in <paramref name="main"/>'s
    /// <see cref="FootnotesPart"/>, or 1 if the part does not exist yet. Footnote and endnote ids
    /// are independent numbering spaces — measured, not assumed — so this scans only the
    /// footnotes part, never the endnotes part too.
    /// </summary>
    private static int NextFootnoteId(MainDocumentPart main)
    {
        if (main.FootnotesPart?.Footnotes is not { } footnotes) return 1;

        var highest = footnotes.Elements<Footnote>()
            .Select(f => f.Id?.Value)
            .Where(id => id.HasValue)
            .Select(id => (int)id!.Value)
            .DefaultIfEmpty(0)
            .Max();

        return highest + 1;
    }

    /// <summary>
    /// Appends one footnote entry to <paramref name="main"/>'s <see cref="FootnotesPart"/>,
    /// creating the part on first use. No separator/continuationSeparator boilerplate: measured
    /// against a real OpenXmlValidator run, a minimal part with only real footnote
    /// entries validates with zero errors — the boilerplate is a Word authoring convention, not a
    /// schema requirement.
    /// </summary>
    private static void AddFootnoteEntry(MainDocumentPart main, int id, string footnoteText)
    {
        var footnotesPart = main.FootnotesPart ?? main.AddNewPart<FootnotesPart>();
        footnotesPart.Footnotes ??= new Footnotes();

        footnotesPart.Footnotes.AppendChild(new Footnote(
            new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "FootnoteText" }),
                new Run(
                    new RunProperties(new RunStyle { Val = "FootnoteReference" }),
                    new FootnoteReferenceMark()),
                new Run(new Text(footnoteText) { Space = SpaceProcessingModeValues.Preserve })))
        { Id = id });

        footnotesPart.Footnotes.Save();
    }

    /// <summary>
    /// Adds an endnote at every occurrence of <paramref name="placeholder"/>, inline, across the
    /// document body — see <see cref="AddFootnote(byte[], string, string)"/> for exactly what is
    /// matched and how each occurrence is handled. The only difference is where the note ends up:
    /// the document's endnotes, not its footnotes.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any of the three required arguments is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="docx"/> is empty, or <paramref name="placeholder"/> is blank.
    /// </exception>
    /// <exception cref="DocumentConversionException">
    /// The package could not be edited, or <paramref name="placeholder"/> does not appear in the
    /// body — a call matching nothing is a bug in the call or the template, not a no-op.
    /// </exception>
    public static byte[] AddEndnote(byte[] docx, string placeholder, string endnoteText)
    {
        ArgumentNullException.ThrowIfNull(docx);
        ArgumentNullException.ThrowIfNull(placeholder);
        ArgumentNullException.ThrowIfNull(endnoteText);
        if (docx.Length == 0) throw new ArgumentException("DOCX content was empty.", nameof(docx));
        if (string.IsNullOrWhiteSpace(placeholder))
            throw new ArgumentException("Placeholder was blank.", nameof(placeholder));

        using var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        ms.Position = 0;

        AddEndnoteCore(ms, placeholder, endnoteText);
        return ms.ToArray();
    }

    /// <summary>
    /// Reads a .docx from <paramref name="source"/>, adds an endnote at every occurrence of
    /// <paramref name="placeholder"/>, and writes the result to <paramref name="destination"/> —
    /// see <see cref="AddEndnote(byte[], string, string)"/> for exactly what is matched.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought, and neither has to be seekable.
    /// </summary>
    /// <param name="source">The stream the .docx package is read from.</param>
    /// <param name="placeholder">The literal placeholder text, braces included.</param>
    /// <param name="endnoteText">The endnote's own text.</param>
    /// <param name="destination">The stream the edited .docx package is written to.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, <paramref name="destination"/>
    /// is not writable, or <paramref name="placeholder"/> is blank.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The package could not be edited, or the placeholder was not found.
    /// </exception>
    public static async Task AddEndnoteAsync(
        Stream source, string placeholder, string endnoteText, Stream destination,
        CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(placeholder);
        ArgumentNullException.ThrowIfNull(endnoteText);
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        if (string.IsNullOrWhiteSpace(placeholder))
            throw new ArgumentException("Placeholder was blank.", nameof(placeholder));

        using var buffer = await StreamPipeline
            .DrainAsync(source, "DOCX content was empty.", nameof(source), "Failed to add an endnote to the DOCX package. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        AddEndnoteCore(buffer, placeholder, endnoteText);

        await StreamPipeline
            .EmitAsync(buffer, destination, "Failed to add an endnote to the DOCX package. See the inner exception for details.", ct)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Reads a .docx from <paramref name="inputPath"/>, adds an endnote at every occurrence of
    /// <paramref name="placeholder"/>, and writes the result to <paramref name="outputPath"/> —
    /// see <see cref="AddEndnote(byte[], string, string)"/> for exactly what is matched. The two
    /// paths may be the same file: the updated bytes are computed in full before
    /// <paramref name="outputPath"/> is opened.
    /// </summary>
    /// <param name="inputPath">The .docx to read.</param>
    /// <param name="outputPath">Where to write the result. Overwritten if it exists.</param>
    /// <param name="placeholder">The literal placeholder text, braces included.</param>
    /// <param name="endnoteText">The endnote's own text.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    /// <exception cref="ArgumentNullException">A path or <paramref name="placeholder"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// A path is blank, the file at <paramref name="inputPath"/> is empty, or
    /// <paramref name="placeholder"/> is blank.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="inputPath"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException">
    /// <paramref name="inputPath"/>'s or <paramref name="outputPath"/>'s directory does not exist.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The package could not be edited, or the placeholder was not found.
    /// </exception>
    public static async Task AddEndnoteAsync(
        string inputPath, string outputPath, string placeholder, string endnoteText,
        CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        ArgumentNullException.ThrowIfNull(placeholder);
        ArgumentNullException.ThrowIfNull(endnoteText);

        var bytes = await FilePipeline.ReadAsync(inputPath, nameof(inputPath), ct).ConfigureAwait(false);
        var result = AddEndnote(bytes, placeholder, endnoteText);
        await File.WriteAllBytesAsync(outputPath, result, ct).ConfigureAwait(false);
    }

    /// <summary>The one real implementation; every overload calls it so they cannot drift apart.</summary>
    private static void AddEndnoteCore(MemoryStream ms, string placeholder, string endnoteText)
    {
        try
        {
            using (var doc = WordprocessingDocument.Open(ms, true))
            {
                var main = doc.MainDocumentPart
                           ?? throw new DocumentConversionException("Document has no main part. This usually means the file is not really a .docx (for "
                           + "example it was renamed from another format) or the upload is corrupt.");
                var body = main.Document?.Body
                           ?? throw new DocumentConversionException("Document has no body. This usually means the file is not really a .docx (for "
                           + "example it was renamed from another format) or the upload is corrupt.");

                var inserted = InsertEndnoteReferencesIn(main, body, placeholder, endnoteText);

                if (inserted == 0)
                {
                    throw new DocumentConversionException(
                        $"The placeholder '{placeholder}' was not found, so there was no endnote to add. "
                        + "Check the placeholder text, braces included, matches the document exactly.");
                }

                main.Document!.Save();
            }

            ms.Position = 0;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to add an endnote to the DOCX package. See the inner exception for details.", ex);
        }
    }

    private static int InsertEndnoteReferencesIn(
        MainDocumentPart main, Body body, string placeholder, string endnoteText)
    {
        var inserted = 0;
        var nextId = NextEndnoteId(main);

        foreach (var paragraph in body.Descendants<Paragraph>().ToList())
        {
            var texts = paragraph.Descendants<Text>()
                                 .Where(t => t.Ancestors<Paragraph>().FirstOrDefault() == paragraph)
                                 .ToList();
            if (texts.Count == 0) continue;

            var merged = string.Concat(texts.Select(t => t.Text));

            var offsets = new List<int>();
            for (var at = merged.IndexOf(placeholder, StringComparison.Ordinal);
                 at >= 0;
                 at = merged.IndexOf(placeholder, at + placeholder.Length, StringComparison.Ordinal))
            {
                offsets.Add(at);
            }

            // Right to left, so the offsets of earlier matches stay valid as later ones are spliced.
            for (var i = offsets.Count - 1; i >= 0; i--)
            {
                var id = nextId++;
                AddEndnoteEntry(main, id, endnoteText);

                var referenceRun = new Run(
                    new RunProperties(new RunStyle { Val = "EndnoteReference" }),
                    new EndnoteReference { Id = id });
                SpliceElementIn(texts, offsets[i], placeholder.Length, referenceRun);
                inserted++;
            }
        }

        return inserted;
    }

    /// <summary>
    /// One above the highest existing endnote id in <paramref name="main"/>'s
    /// <see cref="EndnotesPart"/>, or 1 if the part does not exist yet — independent of whatever
    /// footnote ids exist, per <see cref="NextFootnoteId"/>'s own doc comment.
    /// </summary>
    private static int NextEndnoteId(MainDocumentPart main)
    {
        if (main.EndnotesPart?.Endnotes is not { } endnotes) return 1;

        var highest = endnotes.Elements<Endnote>()
            .Select(e => e.Id?.Value)
            .Where(id => id.HasValue)
            .Select(id => (int)id!.Value)
            .DefaultIfEmpty(0)
            .Max();

        return highest + 1;
    }

    /// <summary>
    /// Appends one endnote entry to <paramref name="main"/>'s <see cref="EndnotesPart"/>, creating
    /// the part on first use — see <see cref="AddFootnoteEntry"/> for why no separator boilerplate
    /// is written.
    /// </summary>
    private static void AddEndnoteEntry(MainDocumentPart main, int id, string endnoteText)
    {
        var endnotesPart = main.EndnotesPart ?? main.AddNewPart<EndnotesPart>();
        endnotesPart.Endnotes ??= new Endnotes();

        endnotesPart.Endnotes.AppendChild(new Endnote(
            new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "EndnoteText" }),
                new Run(
                    new RunProperties(new RunStyle { Val = "EndnoteReference" }),
                    new EndnoteReferenceMark()),
                new Run(new Text(endnoteText) { Space = SpaceProcessingModeValues.Preserve })))
        { Id = id });

        endnotesPart.Endnotes.Save();
    }

    /// <summary>Reads a .docx from <paramref name="path"/> and returns its body text.</summary>
    /// <param name="path">The .docx to read.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <returns>The document's body text.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="path"/> is blank, or the file it names is empty.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="path"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException"><paramref name="path"/>'s directory does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be processed.</exception>
    public static Task<string> ExtractTextAsync(string path, CancellationToken ct = default)
        => ExtractTextAsync(path, includeHeadersAndFooters: false, ct);

    /// <summary>
    /// Reads a .docx from <paramref name="path"/> and returns its text, optionally including
    /// headers and footers.
    /// </summary>
    /// <param name="path">The .docx to read.</param>
    /// <param name="includeHeadersAndFooters">Whether to include header and footer text.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <returns>The document's text.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="path"/> is blank, or the file it names is empty.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="path"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException"><paramref name="path"/>'s directory does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be processed.</exception>
    public static async Task<string> ExtractTextAsync(
        string path, bool includeHeadersAndFooters, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        var bytes = await FilePipeline.ReadAsync(path, nameof(path), ct).ConfigureAwait(false);
        return ExtractText(bytes, includeHeadersAndFooters);
    }

    /// <summary>
    /// A copy of <paramref name="docx"/> encrypted with <paramref name="password"/>, so it cannot
    /// be opened without one.
    /// </summary>
    /// <remarks>
    /// <b>This is file encryption, not document protection.</b> Office offers both under the same
    /// menu and they are not the same thing: this scrambles the whole file, so nothing can be read
    /// without the password. The other kind - a flag asking a reader not to edit - is a request
    /// rather than a lock, and is deliberately not exposed here.
    ///
    /// <b>The result is not a DOCX package any more.</b> An encrypted Office document is a
    /// compound file with the package sealed inside it, so every other method on this class refuses
    /// it - call <see cref="Unprotect(byte[], string)"/> first. That refusal is the honest
    /// behaviour: those methods could not read the content even if they tried.
    /// </remarks>
    /// <param name="docx">The document to encrypt.</param>
    /// <param name="password">The password required to open the result. May not be empty.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The document could not be read or encrypted.</exception>
    public static byte[] Protect(byte[] docx, string password)
    {
        ArgumentNullException.ThrowIfNull(docx);
        if (docx.Length == 0)
            throw new ArgumentException("DOCX content was empty.", nameof(docx));
        OfficeCrypto.RequirePassword(password, nameof(password));

        return OfficeCrypto.TranslateWrite(() =>
        {
            using var source = new MemoryStream(docx, writable: false);
            using var document = OfficeIMOWordWordDocument.Load(source);
            using var encrypted = new MemoryStream();
            document.SaveEncrypted(encrypted, password);
            return encrypted.ToArray();
        }, "DOCX");
    }

    /// <summary>
    /// A copy of <paramref name="docx"/> with its encryption removed, so the rest of this class
    /// can work on it.
    /// </summary>
    /// <remarks>
    /// <b>The output is not protected in any way.</b> That is what was asked for, but the bytes
    /// this returns are readable by anyone who obtains them.
    ///
    /// A document that was never encrypted is reported as such rather than passed through, because
    /// silently returning the input would make a broken pipeline look like a working one.
    /// </remarks>
    /// <param name="docx">The encrypted document.</param>
    /// <param name="password">The password the document was encrypted with.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">
    /// The password was wrong, the document was not encrypted, or it could not be read.
    /// </exception>
    public static byte[] Unprotect(byte[] docx, string password)
    {
        ArgumentNullException.ThrowIfNull(docx);
        if (docx.Length == 0)
            throw new ArgumentException("DOCX content was empty.", nameof(docx));
        OfficeCrypto.RequirePassword(password, nameof(password));

        return OfficeCrypto.Translate(() =>
        {
            using var source = new MemoryStream(docx, writable: false);
            using var document = OfficeIMOWordWordDocument.LoadEncrypted(source, password);
            using var plain = new MemoryStream();
            document.Save(plain);
            return plain.ToArray();
        }, "DOCX");
    }

    /// <summary>
    /// Whether <paramref name="docx"/> is an ENCRYPTED Office document.
    /// </summary>
    /// <remarks>
    /// <b>This is not a validity check, and a <see langword="false"/> is not a promise that
    /// anything else will succeed.</b> It distinguishes an encrypted document from a plain one;
    /// input that is neither — an image, a PDF, a text file, random bytes — is not encrypted, so
    /// this answers <see langword="false"/> for it, while every other method on this class refuses
    /// it. Measured over real files: a JPEG and a PDF both return <see langword="false"/> here and
    /// both throw from <c>ExtractText</c>.
    ///
    /// <b>The summary used to say "that is, whether the other methods on this class will refuse
    /// it".</b> That reads as a guard — test it, and if false, proceed — and takes the wrong branch
    /// for every input that is not a document at all. The behaviour was always right and only the
    /// sentence was wrong, which is why the fix is here and not in the code.
    ///
    /// Reads the file signature; it does not try the password and does not need one. A plain DOCX
    /// is a ZIP package, an encrypted one is a compound file, and the two are distinguishable from
    /// their first eight bytes.
    /// </remarks>
    /// <param name="docx">The bytes to inspect.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    public static bool IsProtected(byte[] docx)
    {
        ArgumentNullException.ThrowIfNull(docx);

        return OfficeCrypto.IsEncrypted(docx);
    }

    /// <summary>
    /// Reads a document from <paramref name="source"/> and writes the encrypted copy to
    /// <paramref name="destination"/>.
    ///
    /// Neither stream is disposed, closed or sought.
    /// </summary>
    /// <inheritdoc cref="Protect(byte[], string)" path="/remarks|/exception"/>
    /// <param name="source">The stream the document is read from.</param>
    /// <param name="destination">The stream the encrypted document is written to.</param>
    /// <param name="password">The password required to open the result. May not be empty.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    public static async Task ProtectAsync(
        Stream source, Stream destination, string password, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        OfficeCrypto.RequirePassword(password, nameof(password));
        ct.ThrowIfCancellationRequested();

        using var buffer = await StreamPipeline
            .DrainAsync(source, "DOCX content was empty.", nameof(source),
                        "Failed to encrypt the DOCX.", ct)
            .ConfigureAwait(false);

        using var result = new MemoryStream(Protect(buffer.ToArray(), password), writable: false);
        await StreamPipeline
            .EmitAsync(result, destination, "Failed to encrypt the DOCX.", ct)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Reads an encrypted document from <paramref name="source"/> and writes the unprotected copy to
    /// <paramref name="destination"/>.
    ///
    /// Neither stream is disposed, closed or sought.
    /// </summary>
    /// <inheritdoc cref="Unprotect(byte[], string)" path="/remarks|/exception"/>
    /// <param name="source">The stream the encrypted document is read from.</param>
    /// <param name="destination">The stream the unprotected document is written to.</param>
    /// <param name="password">The password the document was encrypted with.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    public static async Task UnprotectAsync(
        Stream source, Stream destination, string password, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        OfficeCrypto.RequirePassword(password, nameof(password));
        ct.ThrowIfCancellationRequested();

        using var buffer = await StreamPipeline
            .DrainAsync(source, "DOCX content was empty.", nameof(source),
                        "Failed to read the encrypted DOCX.", ct)
            .ConfigureAwait(false);

        using var result = new MemoryStream(Unprotect(buffer.ToArray(), password), writable: false);
        await StreamPipeline
            .EmitAsync(result, destination, "Failed to read the encrypted DOCX.", ct)
            .ConfigureAwait(false);
    }
}
