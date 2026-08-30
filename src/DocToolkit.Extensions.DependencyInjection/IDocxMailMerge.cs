namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Fills a Word mail-merge template — a document carrying <c>MERGEFIELD</c> instructions — from a set
/// of named values. Registered by <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.
/// </summary>
/// <remarks>
/// <b>Not <see cref="IDocxEditor"/>'s placeholders and not <see cref="IDocxForm"/>'s content
/// controls.</b> The difference is who authored the template: <c>{{placeholder}}</c> is a convention
/// this library invented, a <c>MERGEFIELD</c> is what Word writes from <i>Insert → Merge Field</i>,
/// and a content control is a named region Word protects. A caller has whichever one their document
/// was built with.
/// </remarks>
public interface IDocxMailMerge
{
    /// <summary>Reads what <paramref name="docx"/> asks for, without merging anything.</summary>
    /// <param name="docx">The template to read.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be opened or read.</exception>
    DocToolkit.DocxMailMergeTemplate InspectTemplate(byte[] docx);

    /// <inheritdoc cref="InspectTemplate(byte[])" path="/summary|/remarks|/exception"/>
    /// <param name="source">The template to read. Read to its end; never disposed or sought.</param>
    /// <param name="ct">Cancels before the document is read.</param>
    Task<DocToolkit.DocxMailMergeTemplate> InspectTemplateAsync(
        Stream source, CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="docx"/> with every merge field filled.
    /// </summary>
    /// <remarks>
    /// <b>Refuses to produce a document with an unfilled field</b>, naming every one. Measured: an
    /// unfilled field survives as a live field and the document reads <c>«Balance»</c> — valid,
    /// opening cleanly, and looking finished. Use
    /// <see cref="MergeWithReport(byte[], IReadOnlyDictionary{string, string})"/> when you want it
    /// anyway.
    /// </remarks>
    /// <param name="docx">The template to fill.</param>
    /// <param name="values">The value for each field, matched case-insensitively.</param>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a value is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// A field received no value, or the document could not be read or written.
    /// </exception>
    byte[] Merge(byte[] docx, IReadOnlyDictionary<string, string> values);

    /// <inheritdoc cref="Merge(byte[], IReadOnlyDictionary{string, string})" path="/summary|/remarks|/exception"/>
    /// <param name="source">The template to fill. Read to its end; never disposed or sought.</param>
    /// <param name="destination">Receives the filled document. Written; never disposed or sought.</param>
    /// <param name="values">The value for each field, matched case-insensitively.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    Task MergeAsync(
        Stream source, Stream destination, IReadOnlyDictionary<string, string> values,
        CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="docx"/> with every merge field filled, <b>together with what
    /// happened to each one</b>. Always produces a document, complete or not.
    /// </summary>
    /// <param name="docx">The template to fill.</param>
    /// <param name="values">The value for each field, matched case-insensitively.</param>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a value is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be read or written.</exception>
    DocToolkit.DocxMailMergeResult MergeWithReport(
        byte[] docx, IReadOnlyDictionary<string, string> values);

    /// <inheritdoc cref="MergeWithReport(byte[], IReadOnlyDictionary{string, string})" path="/summary|/remarks|/exception"/>
    /// <remarks>
    /// Returns the report alone, because the document went to <paramref name="destination"/>.
    /// </remarks>
    /// <param name="source">The template to fill. Read to its end; never disposed or sought.</param>
    /// <param name="destination">Receives the filled document. Written; never disposed or sought.</param>
    /// <param name="values">The value for each field, matched case-insensitively.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    Task<DocToolkit.DocxMailMergeReport> MergeWithReportAsync(
        Stream source, Stream destination, IReadOnlyDictionary<string, string> values,
        CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="docx"/> with every conditional block resolved, refusing if the
    /// template asks for a condition <paramref name="conditions"/> did not supply.
    /// </summary>
    /// <param name="docx">The template to resolve.</param>
    /// <param name="conditions">Whether to include each named block.</param>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// A condition the template asks for was not supplied, the marker structure is unbalanced, or
    /// the document could not be read or written.
    /// </exception>
    byte[] MergeConditional(byte[] docx, IReadOnlyDictionary<string, bool> conditions);

    /// <inheritdoc cref="MergeConditional(byte[], IReadOnlyDictionary{string, bool})" path="/summary|/exception"/>
    /// <param name="source">The template to resolve. Read to its end; never disposed or sought.</param>
    /// <param name="destination">Receives the resolved document. Written; never disposed or sought.</param>
    /// <param name="conditions">Whether to include each named block.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    Task MergeConditionalAsync(
        Stream source, Stream destination, IReadOnlyDictionary<string, bool> conditions,
        CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="docx"/> with every conditional block resolved, <b>together with
    /// which condition names the template asked for that <paramref name="conditions"/> did not
    /// supply</b>. Never refuses for a missing name — an unsupplied condition defaults to
    /// <see langword="false"/>.
    /// </summary>
    /// <param name="docx">The template to resolve.</param>
    /// <param name="conditions">Whether to include each named block.</param>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The marker structure is unbalanced, or the document could not be read or written.
    /// </exception>
    DocToolkit.DocxMailMergeBlockResult MergeConditionalWithReport(
        byte[] docx, IReadOnlyDictionary<string, bool> conditions);

    /// <inheritdoc cref="MergeConditionalWithReport(byte[], IReadOnlyDictionary{string, bool})" path="/summary|/exception"/>
    /// <remarks>Returns the report alone, because the document went to <paramref name="destination"/>.</remarks>
    /// <param name="source">The template to resolve. Read to its end; never disposed or sought.</param>
    /// <param name="destination">Receives the resolved document. Written; never disposed or sought.</param>
    /// <param name="conditions">Whether to include each named block.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    Task<DocToolkit.DocxMailMergeBlockReport> MergeConditionalWithReportAsync(
        Stream source, Stream destination, IReadOnlyDictionary<string, bool> conditions,
        CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="docx"/> with every repeating block (<c>{{#each Name}}</c> …
    /// <c>{{/each Name}}</c>) expanded once per entry in its region, refusing if the template asks
    /// for a region <paramref name="regions"/> did not supply.
    /// </summary>
    /// <param name="docx">The template to expand.</param>
    /// <param name="regions">One sequence of value sets per named repeating region.</param>
    /// <exception cref="ArgumentNullException">An argument is null, or an individual record in it is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a record's value is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// A region the template asks for was not supplied, the marker structure is unbalanced, or the
    /// document could not be read or written.
    /// </exception>
    byte[] MergeRepeating(
        byte[] docx, IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> regions);

    /// <inheritdoc cref="MergeRepeating(byte[], IReadOnlyDictionary{string, IEnumerable{IReadOnlyDictionary{string, string}}})" path="/summary|/exception"/>
    /// <param name="source">The template to expand. Read to its end; never disposed or sought.</param>
    /// <param name="destination">Receives the expanded document. Written; never disposed or sought.</param>
    /// <param name="regions">One sequence of value sets per named repeating region.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    Task MergeRepeatingAsync(
        Stream source, Stream destination,
        IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> regions,
        CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="docx"/> with every repeating block expanded, <b>together with
    /// which region names the template asked for that <paramref name="regions"/> did not supply</b>.
    /// An unsupplied region defaults to zero rows rather than refusing.
    /// </summary>
    /// <param name="docx">The template to expand.</param>
    /// <param name="regions">One sequence of value sets per named repeating region.</param>
    /// <exception cref="ArgumentNullException">An argument is null, or an individual record in it is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a record's value is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The marker structure is unbalanced, or the document could not be read or written.
    /// </exception>
    DocToolkit.DocxMailMergeBlockResult MergeRepeatingWithReport(
        byte[] docx, IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> regions);

    /// <inheritdoc cref="MergeRepeatingWithReport(byte[], IReadOnlyDictionary{string, IEnumerable{IReadOnlyDictionary{string, string}}})" path="/summary|/exception"/>
    /// <remarks>Returns the report alone, because the document went to <paramref name="destination"/>.</remarks>
    /// <param name="source">The template to expand. Read to its end; never disposed or sought.</param>
    /// <param name="destination">Receives the expanded document. Written; never disposed or sought.</param>
    /// <param name="regions">One sequence of value sets per named repeating region.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    Task<DocToolkit.DocxMailMergeBlockReport> MergeRepeatingWithReportAsync(
        Stream source, Stream destination,
        IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> regions,
        CancellationToken ct = default);

    /// <summary>
    /// The nested-region twin of <see cref="MergeRepeating"/> — each entry may itself carry further
    /// nested regions, for a template whose repeating blocks are nested inside one another.
    /// </summary>
    /// <param name="docx">The template to expand.</param>
    /// <param name="regions">One sequence of block rows per named top-level repeating region.</param>
    /// <exception cref="ArgumentNullException">An argument is null, or an individual block row in it is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a block row's value is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// A region the template asks for was not supplied at any nesting level, the marker structure is
    /// unbalanced, or the document could not be read or written.
    /// </exception>
    byte[] MergeRepeatingRegions(
        byte[] docx, IReadOnlyDictionary<string, IEnumerable<DocToolkit.DocxMailMergeBlockData>> regions);

    /// <inheritdoc cref="MergeRepeatingRegions(byte[], IReadOnlyDictionary{string, IEnumerable{DocToolkit.DocxMailMergeBlockData}})" path="/summary|/exception"/>
    /// <param name="source">The template to expand. Read to its end; never disposed or sought.</param>
    /// <param name="destination">Receives the expanded document. Written; never disposed or sought.</param>
    /// <param name="regions">One sequence of block rows per named top-level repeating region.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    Task MergeRepeatingRegionsAsync(
        Stream source, Stream destination,
        IReadOnlyDictionary<string, IEnumerable<DocToolkit.DocxMailMergeBlockData>> regions,
        CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="docx"/> with every nested repeating region expanded, <b>together
    /// with which region names the template asked for — at any nesting level — that
    /// <paramref name="regions"/> did not supply</b>. An unsupplied region, at any level, defaults
    /// to zero rows rather than refusing.
    /// </summary>
    /// <param name="docx">The template to expand.</param>
    /// <param name="regions">One sequence of block rows per named top-level repeating region.</param>
    /// <exception cref="ArgumentNullException">An argument is null, or an individual block row in it is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a block row's value is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The marker structure is unbalanced, or the document could not be read or written.
    /// </exception>
    DocToolkit.DocxMailMergeBlockResult MergeRepeatingRegionsWithReport(
        byte[] docx, IReadOnlyDictionary<string, IEnumerable<DocToolkit.DocxMailMergeBlockData>> regions);

    /// <inheritdoc cref="MergeRepeatingRegionsWithReport(byte[], IReadOnlyDictionary{string, IEnumerable{DocToolkit.DocxMailMergeBlockData}})" path="/summary|/exception"/>
    /// <remarks>Returns the report alone, because the document went to <paramref name="destination"/>.</remarks>
    /// <param name="source">The template to expand. Read to its end; never disposed or sought.</param>
    /// <param name="destination">Receives the expanded document. Written; never disposed or sought.</param>
    /// <param name="regions">One sequence of block rows per named top-level repeating region.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    Task<DocToolkit.DocxMailMergeBlockReport> MergeRepeatingRegionsWithReportAsync(
        Stream source, Stream destination,
        IReadOnlyDictionary<string, IEnumerable<DocToolkit.DocxMailMergeBlockData>> regions,
        CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="docx"/> with the row at <paramref name="templateRowIndex"/> in the
    /// table at <paramref name="tableIndex"/> repeated once per entry in <paramref name="rows"/>.
    /// Index-based rather than marker-based — no strict/lenient split, since there is no name that
    /// could go unsupplied.
    /// </summary>
    /// <param name="docx">The template to expand.</param>
    /// <param name="tableIndex">Zero-based index of the table, in document order.</param>
    /// <param name="templateRowIndex">Zero-based row index within that table to clone and bind.</param>
    /// <param name="rows">One value set per generated row.</param>
    /// <exception cref="ArgumentNullException">An argument is null, or an individual row in it is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a row's value is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// <paramref name="tableIndex"/> or <paramref name="templateRowIndex"/> is out of range, or the
    /// document could not be read or written.
    /// </exception>
    byte[] MergeTableRows(
        byte[] docx, int tableIndex, int templateRowIndex,
        IEnumerable<IReadOnlyDictionary<string, string>> rows);

    /// <inheritdoc cref="MergeTableRows(byte[], int, int, IEnumerable{IReadOnlyDictionary{string, string}})" path="/summary|/exception"/>
    /// <param name="source">The template to expand. Read to its end; never disposed or sought.</param>
    /// <param name="destination">Receives the expanded document. Written; never disposed or sought.</param>
    /// <param name="tableIndex">Zero-based index of the table, in document order.</param>
    /// <param name="templateRowIndex">Zero-based row index within that table to clone and bind.</param>
    /// <param name="rows">One value set per generated row.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    Task MergeTableRowsAsync(
        Stream source, Stream destination, int tableIndex, int templateRowIndex,
        IEnumerable<IReadOnlyDictionary<string, string>> rows, CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="docx"/> with a group/header row and its detail row template, both
    /// in the table at <paramref name="tableIndex"/>, repeated once per group in
    /// <paramref name="groups"/>. Index-based, like <see cref="MergeTableRows"/>; no strict/lenient
    /// split.
    /// </summary>
    /// <param name="docx">The template to expand.</param>
    /// <param name="tableIndex">Zero-based index of the table, in document order.</param>
    /// <param name="groupTemplateRowIndex">Zero-based row index of the group/header row template.</param>
    /// <param name="detailTemplateRowIndex">Zero-based row index of the detail row template.</param>
    /// <param name="groups">One group/header value set, with its detail rows, per generated group.</param>
    /// <exception cref="ArgumentNullException">An argument is null, or an individual group or detail row in it is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a group or detail row's value is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// <paramref name="tableIndex"/>, <paramref name="groupTemplateRowIndex"/> or
    /// <paramref name="detailTemplateRowIndex"/> is out of range, or the document could not be read
    /// or written.
    /// </exception>
    byte[] MergeTableRowGroups(
        byte[] docx, int tableIndex, int groupTemplateRowIndex, int detailTemplateRowIndex,
        IEnumerable<DocToolkit.DocxMailMergeTableRowGroup> groups);

    /// <inheritdoc cref="MergeTableRowGroups(byte[], int, int, int, IEnumerable{DocToolkit.DocxMailMergeTableRowGroup})" path="/summary|/exception"/>
    /// <param name="source">The template to expand. Read to its end; never disposed or sought.</param>
    /// <param name="destination">Receives the expanded document. Written; never disposed or sought.</param>
    /// <param name="tableIndex">Zero-based index of the table, in document order.</param>
    /// <param name="groupTemplateRowIndex">Zero-based row index of the group/header row template.</param>
    /// <param name="detailTemplateRowIndex">Zero-based row index of the detail row template.</param>
    /// <param name="groups">One group/header value set, with its detail rows, per generated group.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    Task MergeTableRowGroupsAsync(
        Stream source, Stream destination, int tableIndex, int groupTemplateRowIndex,
        int detailTemplateRowIndex, IEnumerable<DocToolkit.DocxMailMergeTableRowGroup> groups,
        CancellationToken ct = default);

    /// <summary>
    /// Fills <paramref name="docx"/> once per entry in <paramref name="records"/>, yielding each
    /// filled document in order. Strict — refuses the moment a record is incomplete, mid-sequence;
    /// everything already yielded before that point is unaffected.
    /// </summary>
    /// <param name="docx">The template to fill, once per record.</param>
    /// <param name="records">
    /// One dictionary of values per output document, matched case-insensitively. An empty sequence
    /// yields no documents.
    /// </param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="docx"/> or <paramref name="records"/> is null, or an individual record in
    /// <paramref name="records"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a record's value is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The package could not be read or written, or a record is missing a value for a field the
    /// template requires — the message names the record's position (0-based) and the missing
    /// field(s).
    /// </exception>
    IEnumerable<byte[]> MergeBatch(byte[] docx, IEnumerable<IReadOnlyDictionary<string, string>> records);

    /// <inheritdoc cref="MergeBatch(byte[], IEnumerable{IReadOnlyDictionary{string, string}})" path="/summary|/exception"/>
    /// <remarks>
    /// Argument validation is not thrown until the caller starts enumerating the result — inherent
    /// to how an <see cref="IAsyncEnumerable{T}"/> iterator method defers its whole body.
    /// </remarks>
    /// <param name="docx">The template to fill, once per record.</param>
    /// <param name="records">
    /// One dictionary of values per output document, matched case-insensitively. An empty sequence
    /// yields no documents.
    /// </param>
    /// <param name="ct">Cancels before the next record's merge runs.</param>
    IAsyncEnumerable<byte[]> MergeBatchAsync(
        byte[] docx, IEnumerable<IReadOnlyDictionary<string, string>> records,
        CancellationToken ct = default);

    /// <summary>
    /// Fills <paramref name="docx"/> once per entry in <paramref name="records"/>, yielding each
    /// record's document <b>together with what happened to every field in it</b>. The lenient
    /// half of the pair — always produces a document for every record, complete or not, and never
    /// throws for an incomplete one.
    /// </summary>
    /// <param name="docx">The template to fill, once per record.</param>
    /// <param name="records">
    /// One dictionary of values per output document, matched case-insensitively. An empty sequence
    /// yields no items.
    /// </param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="docx"/> or <paramref name="records"/> is null, or an individual record in
    /// <paramref name="records"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a record's value is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be read or written.</exception>
    IEnumerable<DocToolkit.DocxMailMergeBatchItem> MergeBatchWithReport(
        byte[] docx, IEnumerable<IReadOnlyDictionary<string, string>> records);

    /// <inheritdoc cref="MergeBatchWithReport(byte[], IEnumerable{IReadOnlyDictionary{string, string}})" path="/summary|/exception"/>
    /// <remarks>
    /// Argument validation is not thrown until the caller starts enumerating the result — the same
    /// <see cref="IAsyncEnumerable{T}"/> deferral <see cref="MergeBatchAsync"/> has.
    /// </remarks>
    /// <param name="docx">The template to fill, once per record.</param>
    /// <param name="records">
    /// One dictionary of values per output document, matched case-insensitively. An empty sequence
    /// yields no items.
    /// </param>
    /// <param name="ct">Cancels before the next record's merge runs.</param>
    IAsyncEnumerable<DocToolkit.DocxMailMergeBatchItem> MergeBatchWithReportAsync(
        byte[] docx, IEnumerable<IReadOnlyDictionary<string, string>> records,
        CancellationToken ct = default);
}
