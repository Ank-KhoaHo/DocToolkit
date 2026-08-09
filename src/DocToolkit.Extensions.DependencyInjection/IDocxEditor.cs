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
}
