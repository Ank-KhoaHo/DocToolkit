namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Creates, reads and edits Word (.docx) documents. Registered by <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.</summary>
public interface IDocxEditor
{
    /// <summary>
    /// Builds a document from <paramref name="blocks"/> — headings, paragraphs, tables and inline
    /// images. Content comes from data rather than markup, so there is no HTML to escape and a value
    /// containing <c>&lt;</c> cannot corrupt the document's structure. An empty sequence is valid.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="blocks"/> is null.</exception>
    /// <exception cref="ArgumentException">An element of <paramref name="blocks"/> is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The document could not be built.</exception>
    byte[] Create(IEnumerable<DocToolkit.DocxBlock> blocks);

    /// <summary>Replaces every key with its value across the document body, headers, footers, footnotes and endnotes.</summary>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or edited.</exception>
    byte[] ReplaceText(byte[] docx, IReadOnlyDictionary<string, string> replacements);

    /// <summary>Returns the plain text of the document body. Headers, footers, footnotes and endnotes are not included.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    string ExtractText(byte[] docx);

    /// <summary>Returns the plain text of the document. When <paramref name="includeHeadersAndFooters"/> is true, headers and footers follow the body text; footnotes and endnotes are never included.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    string ExtractText(byte[] docx, bool includeHeadersAndFooters);

    /// <summary>
    /// Expands a template table row once per record. A row holding <c>{{collection.Field}}</c>
    /// markers becomes one row per record, each keeping the template row's formatting.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or <paramref name="collection"/> is blank.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or edited.</exception>
    byte[] FillRows(
        byte[] docx, string collection,
        IEnumerable<IReadOnlyDictionary<string, string>> rows);

    /// <summary>
    /// Replaces a text placeholder with an image, sized from the image's own header unless a
    /// dimension is given. PNG and JPEG only, decided by magic bytes rather than by filename.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> or <paramref name="image"/> is empty, <paramref name="placeholder"/> is blank, or the image is neither PNG nor JPEG.</exception>
    /// <exception cref="ArgumentOutOfRangeException">A supplied size is zero or negative, or the resulting size is larger than a drawing extent can hold.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or edited.</exception>
    byte[] ReplaceImage(
        byte[] docx, string placeholder, byte[] image,
        double? widthPoints = null, double? heightPoints = null);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/>, replaces every key with its value, and writes
    /// the result to <paramref name="destination"/>. See <see cref="ReplaceText"/> for exactly what
    /// counts as a match. <paramref name="source"/> is <b>read</b> to its end and
    /// <paramref name="destination"/> is <b>written</b>; neither is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="destination"/>
    /// is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or edited.</exception>
    Task ReplaceTextAsync(
        Stream source, IReadOnlyDictionary<string, string> replacements, Stream destination,
        CancellationToken ct = default);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and returns the plain text of its body. Headers,
    /// footers, footnotes and endnotes are not included. <paramref name="source"/> is <b>read</b> to
    /// its end and is neither disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    Task<string> ExtractTextAsync(Stream source, CancellationToken ct = default);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and returns its plain text. See
    /// <see cref="ExtractText(byte[], bool)"/> for what <paramref name="includeHeadersAndFooters"/>
    /// controls. <paramref name="source"/> is <b>read</b> to its end and is neither disposed, closed
    /// nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    Task<string> ExtractTextAsync(Stream source, bool includeHeadersAndFooters, CancellationToken ct = default);

    /// <summary>
    /// Builds a document from <paramref name="blocks"/> and writes it to
    /// <paramref name="destination"/>. See <see cref="Create(System.Collections.Generic.IEnumerable{DocToolkit.DocxBlock})"/> for the block semantics.
    /// <paramref name="destination"/> is <b>written</b> and is neither disposed, closed nor sought,
    /// so an HTTP response body is a valid destination.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="blocks"/> or <paramref name="destination"/> is null.</exception>
    /// <exception cref="ArgumentException">An element of <paramref name="blocks"/> is null, or <paramref name="destination"/> is not writable.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The document could not be built or written.</exception>
    Task CreateAsync(
        IEnumerable<DocToolkit.DocxBlock> blocks, Stream destination, CancellationToken ct = default);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/>, expands the template row once per record, and
    /// writes the result to <paramref name="destination"/>. See <see cref="FillRows"/> for the
    /// expansion rules. Neither stream is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes, <paramref name="destination"/> is not writable, or <paramref name="collection"/> is blank.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or edited.</exception>
    Task FillRowsAsync(
        Stream source, string collection,
        IEnumerable<IReadOnlyDictionary<string, string>> rows, Stream destination,
        CancellationToken ct = default);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/>, replaces the placeholder with an image, and
    /// writes the result to <paramref name="destination"/>. See <see cref="ReplaceImage"/> for
    /// sizing and format rules. Neither stream is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes, <paramref name="destination"/> is not writable, <paramref name="placeholder"/> is blank, or the image is neither PNG nor JPEG.</exception>
    /// <exception cref="ArgumentOutOfRangeException">A supplied size is zero or negative, or the resulting size is larger than a drawing extent can hold.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or edited.</exception>
    Task ReplaceImageAsync(
        Stream source, string placeholder, byte[] image, Stream destination,
        double? widthPoints = null, double? heightPoints = null,
        CancellationToken ct = default);

    /// <summary>
    /// As above, laid out on <paramref name="page"/> rather than the A4 default.
    /// </summary>
    /// <param name="blocks">The content, written in order.</param>
    /// <param name="page">The page size, orientation and margins.</param>
    /// <exception cref="ArgumentNullException"><paramref name="blocks"/> or <paramref name="page"/> is null.</exception>
    /// <exception cref="ArgumentException">An element of <paramref name="blocks"/> is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The document could not be built.</exception>
    byte[] Create(System.Collections.Generic.IEnumerable<DocToolkit.DocxBlock> blocks, DocToolkit.PageSetup page);

    /// <summary>
    /// As above, laid out on <paramref name="page"/> rather than the A4 default.
    /// </summary>
    /// <param name="blocks">The content, written in order.</param>
    /// <param name="page">The page size, orientation and margins.</param>
    /// <param name="destination">The stream the document is written to.</param>
    /// <param name="ct">Cancels the build and the write.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">An element of <paramref name="blocks"/> is null, or <paramref name="destination"/> is not writable.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The document could not be built or written.</exception>
    Task CreateAsync(System.Collections.Generic.IEnumerable<DocToolkit.DocxBlock> blocks, DocToolkit.PageSetup page, Stream destination, CancellationToken ct = default);

    /// <summary>
    /// How many tables the document body holds.
    /// </summary>
    /// <remarks>
    /// Top-level tables only. A table nested inside a cell is part of that cell's text rather than
    /// an entry of its own, so this count and the indexes it bounds stay stable.
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be read.</exception>
    int TableCount(byte[] docx);

    /// <inheritdoc cref="TableCount(byte[])"/>
    /// <remarks>
    /// <paramref name="source"/> is read to its end; it is not disposed, closed or sought, and does
    /// not have to be seekable.
    /// </remarks>
    Task<int> TableCountAsync(Stream source, CancellationToken ct = default);

    /// <summary>
    /// The table at <paramref name="index"/>, as rows of cell text.
    /// </summary>
    /// <param name="docx">The .docx content to read.</param>
    /// <param name="index">
    /// <b>0-based</b>, indexing what <see cref="TableCount(byte[])"/> reports.
    /// </param>
    /// <remarks>
    /// Cell text is produced the same way <see cref="ExtractText(byte[])"/> produces it, so a cell
    /// holding several paragraphs is separated by newlines and a nested table keeps its own
    /// structure.
    ///
    /// <b>Rows are returned with the shape they have.</b> A horizontally merged cell means a row
    /// genuinely holds fewer cells than its neighbours; padding the grid to a rectangle would invent
    /// cells that are not in the document.
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="index"/> is negative, or at or beyond <see cref="TableCount(byte[])"/>.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be read.</exception>
    IReadOnlyList<IReadOnlyList<string>> ReadTable(byte[] docx, int index);

    /// <inheritdoc cref="ReadTable(byte[], int)"/>
    /// <remarks>
    /// <paramref name="source"/> is read to its end; it is not disposed, closed or sought, and does
    /// not have to be seekable.
    /// </remarks>
    Task<IReadOnlyList<IReadOnlyList<string>>> ReadTableAsync(Stream source, int index, CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="docx"/> encrypted with <paramref name="password"/>.
    /// </summary>
    /// <remarks>
    /// <b>File encryption, not the "restrict editing" flag.</b> The result is a compound file rather
    /// than a DOCX package, so every other member here refuses it - call
    /// <see cref="Unprotect(byte[], string)"/> first.
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be read or encrypted.</exception>
    byte[] Protect(byte[] docx, string password);

    /// <summary>A copy of <paramref name="docx"/> with its encryption removed.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The password was wrong, the document was not encrypted, or it could not be read.
    /// </exception>
    byte[] Unprotect(byte[] docx, string password);

    /// <summary>
    /// Whether <paramref name="docx"/> is encrypted - that is, whether the other members here
    /// will refuse it. Reads the file signature; needs no password.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    bool IsProtected(byte[] docx);

    /// <summary>
    /// Reads a document from <paramref name="source"/> and writes the encrypted copy to
    /// <paramref name="destination"/>. Neither stream is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either stream is null, or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be encrypted.</exception>
    Task ProtectAsync(Stream source, Stream destination, string password, CancellationToken ct = default);

    /// <summary>
    /// Reads an encrypted document from <paramref name="source"/> and writes the unprotected copy to
    /// <paramref name="destination"/>. Neither stream is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either stream is null, or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The password was wrong, or it could not be read.</exception>
    Task UnprotectAsync(Stream source, Stream destination, string password, CancellationToken ct = default);

    /// <summary>
    /// Adds a footnote at every occurrence of <paramref name="placeholder"/>, inline, across the
    /// document body.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any of the three required arguments is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="docx"/> is empty, or <paramref name="placeholder"/> is blank.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The package could not be edited, or <paramref name="placeholder"/> does not appear in the
    /// body.
    /// </exception>
    byte[] AddFootnote(byte[] docx, string placeholder, string footnoteText);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/>, adds a footnote at every occurrence of
    /// <paramref name="placeholder"/>, and writes the result to <paramref name="destination"/> —
    /// see <see cref="AddFootnote"/>. Neither stream is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, <paramref name="destination"/>
    /// is not writable, or <paramref name="placeholder"/> is blank.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The package could not be edited, or the placeholder was not found.
    /// </exception>
    Task AddFootnoteAsync(
        Stream source, string placeholder, string footnoteText, Stream destination,
        CancellationToken ct = default);

    /// <summary>
    /// Adds an endnote at every occurrence of <paramref name="placeholder"/>, inline, across the
    /// document body.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any of the three required arguments is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="docx"/> is empty, or <paramref name="placeholder"/> is blank.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The package could not be edited, or <paramref name="placeholder"/> does not appear in the
    /// body.
    /// </exception>
    byte[] AddEndnote(byte[] docx, string placeholder, string endnoteText);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/>, adds an endnote at every occurrence of
    /// <paramref name="placeholder"/>, and writes the result to <paramref name="destination"/> —
    /// see <see cref="AddEndnote"/>. Neither stream is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, <paramref name="destination"/>
    /// is not writable, or <paramref name="placeholder"/> is blank.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The package could not be edited, or the placeholder was not found.
    /// </exception>
    Task AddEndnoteAsync(
        Stream source, string placeholder, string endnoteText, Stream destination,
        CancellationToken ct = default);

    /// <summary>
    /// Replaces the paragraph containing only <paramref name="placeholder"/> with a table of
    /// contents spanning heading levels <paramref name="minLevel"/> through
    /// <paramref name="maxLevel"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either required argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="docx"/> is empty, <paramref name="placeholder"/> is blank, or
    /// <paramref name="minLevel"/> is greater than <paramref name="maxLevel"/>.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="minLevel"/> or <paramref name="maxLevel"/> is outside 1-9.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The package could not be edited; no paragraph containing only the placeholder was found; the
    /// paragraph holding it also holds content other than plain text; that paragraph's
    /// <c>w:pPr</c> carries a <c>w:sectPr</c>, so replacing it would discard a section break; or
    /// more than one paragraph's text exactly matches the placeholder.
    /// </exception>
    byte[] AddTableOfContents(byte[] docx, string placeholder, int minLevel = 1, int maxLevel = 3);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/>, replaces the paragraph containing only
    /// <paramref name="placeholder"/> with a table of contents, and writes the result to
    /// <paramref name="destination"/> — see <see cref="AddTableOfContents"/>. Neither stream is
    /// disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, <paramref name="destination"/>
    /// is not writable, <paramref name="placeholder"/> is blank, or <paramref name="minLevel"/> is
    /// greater than <paramref name="maxLevel"/>.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="minLevel"/> or <paramref name="maxLevel"/> is outside 1-9.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The package could not be edited; no matching paragraph was found; the paragraph holding the
    /// placeholder also holds content other than plain text; that paragraph's <c>w:pPr</c>
    /// carries a <c>w:sectPr</c>, so replacing it would discard a section break; or more than one
    /// paragraph's text exactly matches the placeholder.
    /// </exception>
    Task AddTableOfContentsAsync(
        Stream source, string placeholder, Stream destination,
        int minLevel = 1, int maxLevel = 3, CancellationToken ct = default);

    /// <summary>
    /// Inspects <paramref name="docx"/> for digital signatures — whether it carries one, how
    /// many, and who claims to have signed it. Does not validate anything cryptographically; see
    /// <see cref="ValidateSignatures"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The document could not be inspected.</exception>
    DocToolkit.DocumentSignatureInfo InspectSignatures(byte[] docx);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and inspects it for digital signatures — see
    /// <see cref="InspectSignatures"/>. <paramref name="source"/> is read to its end and is
    /// neither disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The document could not be inspected.</exception>
    Task<DocToolkit.DocumentSignatureInfo> InspectSignaturesAsync(Stream source, CancellationToken ct = default);

    /// <summary>
    /// Validates every digital signature <paramref name="docx"/> carries, returning the
    /// report-level tamper-detection verdict alongside each signature's own certificate chain
    /// trust and revocation status. Never performs revocation checking or certificate downloads
    /// over the network, regardless of <paramref name="options"/> — see
    /// <see cref="DocToolkit.DocumentSignatureValidationOptions"/>'s own remarks.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The document could not be validated.</exception>
    DocToolkit.DocumentSignatureValidationReport ValidateSignatures(byte[] docx, DocToolkit.DocumentSignatureValidationOptions? options = null);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and validates its digital signatures — see
    /// <see cref="ValidateSignatures"/>. <paramref name="source"/> is read to its end and is
    /// neither disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The document could not be validated.</exception>
    Task<DocToolkit.DocumentSignatureValidationReport> ValidateSignaturesAsync(
        Stream source, DocToolkit.DocumentSignatureValidationOptions? options = null, CancellationToken ct = default);

    /// <summary>The document properties <paramref name="docx"/> carries.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The document could not be read.</exception>
    DocToolkit.DocumentMetadata ReadMetadata(byte[] docx);

    /// <summary>
    /// A copy of <paramref name="docx"/> carrying <paramref name="metadata"/>.
    /// </summary>
    /// <remarks>
    /// A <see langword="null"/> property leaves what the document already had in place, so
    /// stamping a title does not silently erase an author. Pass an empty string to clear one.
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> or <paramref name="metadata"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The document could not be read or written.</exception>
    byte[] WithMetadata(byte[] docx, DocToolkit.DocumentMetadata metadata);

    /// <inheritdoc cref="IsProtected(byte[])" path="/summary"/>
    /// <remarks>
    /// <paramref name="source"/> is <b>read</b> to its end and is neither disposed, closed nor
    /// sought. Unlike <see cref="IsProtected(byte[])"/>, which answers <see langword="false"/>
    /// for an empty array, an empty <paramref name="source"/> is rejected - every <c>Stream</c>
    /// overload in this package treats a source that held no bytes as a caller error rather
    /// than as content.
    /// </remarks>
    /// <param name="source">The stream the document is read from.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    Task<bool> IsProtectedAsync(Stream source, CancellationToken ct = default);

    /// <inheritdoc cref="ReadMetadata(byte[])" path="/summary"/>
    /// <remarks>
    /// <paramref name="source"/> is <b>read</b> to its end and is neither disposed, closed nor sought.
    /// </remarks>
    /// <param name="source">The stream the document is read from.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The document could not be read.</exception>
    Task<DocToolkit.DocumentMetadata> ReadMetadataAsync(Stream source, CancellationToken ct = default);

    /// <inheritdoc cref="WithMetadata(byte[], DocToolkit.DocumentMetadata)" path="/summary"/>
    /// <remarks>
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought, and neither has to be seekable.
    /// </remarks>
    /// <param name="source">The stream the document is read from.</param>
    /// <param name="metadata">The properties to stamp.</param>
    /// <param name="destination">The stream the updated document is written to.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="destination"/>
    /// is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The document could not be read or written.</exception>
    Task WithMetadataAsync(Stream source, DocToolkit.DocumentMetadata metadata, Stream destination, CancellationToken ct = default);
}
