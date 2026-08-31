namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Creates, reads and edits Excel (.xlsx) workbooks. Registered by <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.</summary>
public interface IWorkbookEditor
{
    /// <summary>Creates a workbook with one sheet populated from <paramref name="rows"/>.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="rows"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="sheetName"/> is blank, or a row is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be built.</exception>
    byte[] Create(string sheetName, IEnumerable<IEnumerable<object?>> rows);

    /// <summary>
    /// Builds a workbook from <paramref name="sheets"/>, one worksheet each, in sequence order.
    /// Content comes from data rather than a template, so there is no source file to edit.
    ///
    /// <para>A cell holding a <see cref="DocToolkit.XlsxFormula"/> is written as a formula. No
    /// cached result is stored, so a reader that only reads cached values sees an empty cell until
    /// Excel has opened and saved the file; this package's own readers compute on read.</para>
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="sheets"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="sheets"/> is empty, contains a null element, or names the same sheet twice.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be built.</exception>
    byte[] Create(IEnumerable<DocToolkit.XlsxSheet> sheets);

    /// <summary>
    /// Appends <paramref name="rows"/> to <paramref name="sheetName"/>, after its last used row,
    /// leaving every other sheet and all existing formatting as it was.
    ///
    /// <para>"Last used" counts a cell comment or a merged range even where the cell has no value,
    /// so a stray comment far below the data pushes the append below it.</para>
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> or <paramref name="rows"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty, <paramref name="sheetName"/> is blank, longer than 31 characters or contains one of <c>: \ / ? * [ ]</c>, or an element of <paramref name="rows"/> is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The sheet was not found, or the package could not be opened or edited.</exception>
    byte[] AppendRows(byte[] xlsx, string sheetName, IEnumerable<IEnumerable<object?>> rows);

    /// <summary>Reads a cell as a string. <paramref name="cellRef"/> is an A1-style reference.</summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty, or a name is blank.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be opened, the sheet does not exist, or the reference is not valid.</exception>
    string ReadCell(byte[] xlsx, string sheetName, string cellRef);

    /// <summary>
    /// Lists every sheet in the workbook, in tab order, including hidden sheets — hiding a sheet
    /// is a presentation choice, not a privacy boundary, and a caller who cannot see a hidden sheet
    /// listed has no way to discover it exists.
    /// </summary>
    /// <param name="xlsx">The workbook bytes.</param>
    /// <returns>The sheet names, in tab order.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be opened.</exception>
    IReadOnlyList<string> SheetNames(byte[] xlsx);

    /// <summary>
    /// Reads a whole sheet as strings, anchored at A1: if the data starts at C3, its first value
    /// is at <c>rows[2][2]</c>. Every row is padded to the last used column, so all rows have the
    /// same length; blank cells — and entirely blank rows inside the range — come back as empty
    /// strings rather than being dropped, which keeps <c>rows[r][c]</c> positionally meaningful.
    ///
    /// Values are produced exactly as <see cref="ReadCell"/> produces them, so the two can never
    /// disagree about what a cell says. A formula cell yields its cached value: nothing in this
    /// library evaluates formulas.
    ///
    /// Text follows the calling thread's <see cref="System.Globalization.CultureInfo.CurrentCulture"/>
    /// — the same rule <see cref="ReadCell"/> uses, and asymmetric with <see cref="Create(string, IEnumerable{IEnumerable{object}})"/>, which
    /// deliberately writes with <see cref="System.Globalization.CultureInfo.InvariantCulture"/> so
    /// the same code produces the same file everywhere. A number such as <c>1234.5</c> reads back
    /// as <c>"1234.5"</c> under an invariant or en-US culture but <c>"1234,5"</c> under de-DE;
    /// callers who parse the returned text as a number should account for that, e.g. by parsing
    /// with an explicit <see cref="System.Globalization.CultureInfo"/> rather than the default.
    /// </summary>
    /// <param name="xlsx">The workbook bytes.</param>
    /// <param name="sheetName">The sheet to read.</param>
    /// <returns>
    /// The sheet's used range, anchored at A1 and padded rectangular. Empty only if the sheet
    /// holds no values and no cell comments: formatting alone never widens the range, but a
    /// comment on an otherwise-blank cell does, because ClosedXML's <c>LastCellUsed()</c> counts
    /// it as used.
    /// </returns>
    /// <remarks>
    /// The whole range is materialised into memory at once, so its cost is proportional to
    /// <c>rows &#215; columns</c>, not to how much of that rectangle actually holds data. To keep
    /// one far-flung stray value from exhausting memory, <see cref="ReadSheet"/> throws
    /// <see cref="DocToolkit.DocumentConversionException"/> rather than allocate when the used
    /// range exceeds 2,000,000 cells.
    /// </remarks>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="xlsx"/> is empty, or <paramref name="sheetName"/> is blank.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the sheet's used range
    /// exceeds the 2,000,000-cell limit <see cref="ReadSheet"/> will materialise.
    /// </exception>
    IReadOnlyList<IReadOnlyList<string>> ReadSheet(byte[] xlsx, string sheetName);

    /// <summary>Sets a cell and returns the updated workbook bytes.</summary>
    /// <exception cref="ArgumentNullException">Any argument other than <paramref name="value"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty, or a name is blank.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be opened, the sheet does not exist, or the reference is not valid.</exception>
    byte[] SetCell(byte[] xlsx, string sheetName, string cellRef, object? value);

    /// <summary>
    /// Builds a workbook with one sheet populated from <paramref name="rows"/> and writes it to
    /// <paramref name="destination"/>. See <see cref="Create(string, IEnumerable{IEnumerable{object}})"/> for the exact typing and culture
    /// rules applied to each cell. <paramref name="destination"/> is <b>written</b> and is not
    /// disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="rows"/> or <paramref name="destination"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="sheetName"/> is blank, a row is null, or <paramref name="destination"/> is
    /// not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be built or written.</exception>
    Task CreateAsync(
        string sheetName, IEnumerable<IEnumerable<object?>> rows, Stream destination,
        CancellationToken ct = default);

    /// <summary>
    /// Reads a workbook from <paramref name="source"/> and returns a cell as a string.
    /// <paramref name="cellRef"/> is an A1-style reference. <paramref name="source"/> is
    /// <b>read</b> to its end and is neither disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or a name is blank.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    Task<string> ReadCellAsync(Stream source, string sheetName, string cellRef, CancellationToken ct = default);

    /// <summary>
    /// Reads a workbook from <paramref name="source"/> and lists every sheet in tab order,
    /// including hidden sheets. <paramref name="source"/> is <b>read</b> to its end and is neither
    /// disposed, closed nor sought.
    /// </summary>
    /// <param name="source">The stream the workbook is read from.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <returns>The sheet names, in tab order.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be opened.</exception>
    Task<IReadOnlyList<string>> SheetNamesAsync(Stream source, CancellationToken ct = default);

    /// <summary>
    /// Builds a workbook from <paramref name="sheets"/> and writes it to
    /// <paramref name="destination"/>. See <see cref="Create(IEnumerable{DocToolkit.XlsxSheet})"/>
    /// for the semantics.
    ///
    /// <para><paramref name="destination"/> is <b>written</b> and is neither disposed, closed nor
    /// sought, so an HTTP response body is a valid destination.</para>
    /// </summary>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="sheets"/> is invalid as above, or <paramref name="destination"/> is not writable.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be built or written.</exception>
    Task CreateAsync(
        IEnumerable<DocToolkit.XlsxSheet> sheets, Stream destination, CancellationToken ct = default);

    /// <summary>
    /// Reads a workbook from <paramref name="source"/>, appends <paramref name="rows"/> to
    /// <paramref name="sheetName"/>, and writes the result to <paramref name="destination"/>. See
    /// <see cref="AppendRows"/> for the semantics.
    ///
    /// <para><paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/>
    /// is <b>written</b>; neither is disposed, closed nor sought.</para>
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="rows"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes, <paramref name="destination"/> is not writable, <paramref name="sheetName"/> is invalid as above, or an element of <paramref name="rows"/> is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The sheet was not found, or the package could not be opened or edited.</exception>
    Task AppendRowsAsync(
        Stream source, string sheetName, IEnumerable<IEnumerable<object?>> rows, Stream destination,
        CancellationToken ct = default);

    /// <summary>
    /// Reads a workbook from <paramref name="source"/> and returns a whole sheet as strings. See
    /// <see cref="ReadSheet"/> for the anchoring, padding, culture and formula rules — this
    /// overload applies the identical logic. <paramref name="source"/> is <b>read</b> to its end
    /// and is neither disposed, closed nor sought.
    /// </summary>
    /// <param name="source">The stream the workbook is read from.</param>
    /// <param name="sheetName">The sheet to read.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <returns>
    /// The sheet's used range, anchored at A1 and padded rectangular. Empty only if the sheet
    /// holds no values and no cell comments — see <see cref="ReadSheet"/>.
    /// </returns>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="sheetName"/>
    /// is blank.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the sheet's used range
    /// exceeds the 2,000,000-cell limit <see cref="ReadSheet"/> will materialise.
    /// </exception>
    Task<IReadOnlyList<IReadOnlyList<string>>> ReadSheetAsync(
        Stream source, string sheetName, CancellationToken ct = default);

    /// <summary>
    /// Reads a workbook from <paramref name="source"/>, sets one cell, and writes the result to
    /// <paramref name="destination"/>. <paramref name="cellRef"/> is an A1-style reference.
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="source"/> or <paramref name="destination"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, a name is blank, or
    /// <paramref name="destination"/> is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    Task SetCellAsync(
        Stream source, string sheetName, string cellRef, object? value, Stream destination,
        CancellationToken ct = default);

    /// <summary>
    /// Applies <paramref name="format"/> to <paramref name="sheetName"/> and returns the workbook.
    /// </summary>
    /// <remarks>
    /// Formatting is applied to an existing workbook rather than being an argument to
    /// <c>Create</c>, so it composes with every way a workbook can arrive - built here, appended
    /// to, or handed in by a caller who never used this library. See
    /// <see cref="DocToolkit.XlsxFormat"/> for the boundary: a CLOSED vocabulary rather than a
    /// small one, and what is still deliberately outside it.
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> or <paramref name="format"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty, or the sheet name is blank.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The workbook could not be opened, or the sheet does not exist.
    /// </exception>
    byte[] Format(byte[] xlsx, string sheetName, DocToolkit.XlsxFormat format);

    /// <summary>
    /// Reads a workbook from <paramref name="source"/>, applies <paramref name="format"/>, and
    /// writes the result to <paramref name="destination"/>. Neither stream is disposed, closed or
    /// sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="format"/> is null, or a stream is null.</exception>
    /// <exception cref="ArgumentException">
    /// A stream is unusable, <paramref name="source"/> held no bytes, or the sheet name is blank.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The workbook could not be opened, or the sheet does not exist.
    /// </exception>
    Task FormatAsync(
        Stream source, string sheetName, DocToolkit.XlsxFormat format, Stream destination,
        CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="xlsx"/> encrypted with <paramref name="password"/>.
    /// </summary>
    /// <remarks>
    /// <b>File encryption, not the "restrict editing" flag.</b> The result is a compound file rather
    /// than a XLSX package, so every other member here refuses it - call
    /// <see cref="Unprotect(byte[], string)"/> first.
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be read or encrypted.</exception>
    byte[] Protect(byte[] xlsx, string password);

    /// <summary>A copy of <paramref name="xlsx"/> with its encryption removed.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The password was wrong, the workbook was not encrypted, or it could not be read.
    /// </exception>
    byte[] Unprotect(byte[] xlsx, string password);

    /// <summary>
    /// Whether <paramref name="xlsx"/> is encrypted - that is, whether the other members here
    /// will refuse it. Reads the file signature; needs no password.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    bool IsProtected(byte[] xlsx);

    /// <summary>
    /// Reads a workbook from <paramref name="source"/> and writes the encrypted copy to
    /// <paramref name="destination"/>. Neither stream is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either stream is null, or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be encrypted.</exception>
    Task ProtectAsync(Stream source, Stream destination, string password, CancellationToken ct = default);

    /// <summary>
    /// Reads an encrypted workbook from <paramref name="source"/> and writes the unprotected copy to
    /// <paramref name="destination"/>. Neither stream is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either stream is null, or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The password was wrong, or it could not be read.</exception>
    Task UnprotectAsync(Stream source, Stream destination, string password, CancellationToken ct = default);

    /// <summary>
    /// Inspects <paramref name="xlsx"/> for digital signatures — whether it carries one, how
    /// many, and who claims to have signed it. Does not validate anything cryptographically; see
    /// <see cref="ValidateSignatures"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be inspected.</exception>
    DocToolkit.DocumentSignatureInfo InspectSignatures(byte[] xlsx);

    /// <summary>
    /// Reads an .xlsx from <paramref name="source"/> and inspects it for digital signatures — see
    /// <see cref="InspectSignatures"/>. <paramref name="source"/> is read to its end and is
    /// neither disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be inspected.</exception>
    Task<DocToolkit.DocumentSignatureInfo> InspectSignaturesAsync(Stream source, CancellationToken ct = default);

    /// <summary>
    /// Validates every digital signature <paramref name="xlsx"/> carries, returning the
    /// report-level tamper-detection verdict alongside each signature's own certificate chain
    /// trust and revocation status. Never performs revocation checking or certificate downloads
    /// over the network, regardless of <paramref name="options"/> — see
    /// <see cref="DocToolkit.DocumentSignatureValidationOptions"/>'s own remarks.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be validated.</exception>
    DocToolkit.DocumentSignatureValidationReport ValidateSignatures(byte[] xlsx, DocToolkit.DocumentSignatureValidationOptions? options = null);

    /// <summary>
    /// Reads an .xlsx from <paramref name="source"/> and validates its digital signatures — see
    /// <see cref="ValidateSignatures"/>. <paramref name="source"/> is read to its end and is
    /// neither disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be validated.</exception>
    Task<DocToolkit.DocumentSignatureValidationReport> ValidateSignaturesAsync(
        Stream source, DocToolkit.DocumentSignatureValidationOptions? options = null, CancellationToken ct = default);

    /// <summary>The document properties <paramref name="xlsx"/> carries.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be read.</exception>
    DocToolkit.DocumentMetadata ReadMetadata(byte[] xlsx);

    /// <summary>
    /// A copy of <paramref name="xlsx"/> carrying <paramref name="metadata"/>.
    /// </summary>
    /// <remarks>
    /// A <see langword="null"/> property leaves what the workbook already had in place, so
    /// stamping a title does not silently erase an author. Pass an empty string to clear one.
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> or <paramref name="metadata"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be read or written.</exception>
    byte[] WithMetadata(byte[] xlsx, DocToolkit.DocumentMetadata metadata);

    /// <summary>
    /// Every formula <paramref name="xlsx"/> carries, and whether each one is understood well
    /// enough to trust its value. See <see cref="DocToolkit.XlsxFormulaInspection"/> for why this
    /// asks rather than assumes.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be read.</exception>
    DocToolkit.XlsxFormulaInspection InspectFormulas(byte[] xlsx);

    /// <summary>
    /// A copy of <paramref name="xlsx"/> with every formula's computed value written into the
    /// file, not just held in memory.
    /// </summary>
    /// <remarks>
    /// A formula <see cref="DocToolkit.XlsxFormulaInspection"/> would report as unsupported is left
    /// exactly as it was — no plausible-looking value is invented for it. Call
    /// <see cref="InspectFormulas"/> first if that distinction matters to the caller.
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be read or written.</exception>
    byte[] EvaluateFormulas(byte[] xlsx);

    /// <summary>
    /// Adds a chart to <paramref name="sheetName"/>, anchored at <paramref name="cellRef"/>, and
    /// returns the updated workbook.
    /// </summary>
    /// <param name="xlsx">The workbook to add the chart to. It is not modified.</param>
    /// <param name="sheetName">The sheet to add the chart to.</param>
    /// <param name="cellRef">
    /// An A1-style cell reference for the chart's top-left corner, e.g. <c>"B2"</c>.
    /// </param>
    /// <param name="type">The chart's shape.</param>
    /// <param name="data">The chart's categories and value series.</param>
    /// <param name="title">The chart's title. Empty for no title.</param>
    /// <param name="widthPixels">The chart's width, in pixels.</param>
    /// <param name="heightPixels">The chart's height, in pixels.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="xlsx"/>, <paramref name="data"/> or another required argument is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="xlsx"/> is empty, or <paramref name="sheetName"/>/<paramref name="cellRef"/>
    /// is blank.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    byte[] AddChart(
        byte[] xlsx, string sheetName, string cellRef, DocToolkit.ChartType type, DocToolkit.ChartData data,
        string title = "", int widthPixels = 640, int heightPixels = 360);

    /// <summary>
    /// Reads a workbook from <paramref name="source"/>, adds a chart, and writes the result to
    /// <paramref name="destination"/> — see <see cref="AddChart"/> for the parameters.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought, and neither has to be seekable.
    /// </summary>
    /// <param name="source">The stream the workbook is read from.</param>
    /// <param name="sheetName">The sheet to add the chart to.</param>
    /// <param name="cellRef">
    /// An A1-style cell reference for the chart's top-left corner, e.g. <c>"B2"</c>.
    /// </param>
    /// <param name="type">The chart's shape.</param>
    /// <param name="data">The chart's categories and value series.</param>
    /// <param name="destination">The stream the updated workbook is written to.</param>
    /// <param name="title">The chart's title. Empty for no title.</param>
    /// <param name="widthPixels">The chart's width, in pixels.</param>
    /// <param name="heightPixels">The chart's height, in pixels.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="source"/>, <paramref name="destination"/> or <paramref name="data"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, a name is blank, or
    /// <paramref name="destination"/> is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    Task AddChartAsync(
        Stream source, string sheetName, string cellRef, DocToolkit.ChartType type, DocToolkit.ChartData data, Stream destination,
        string title = "", int widthPixels = 640, int heightPixels = 360, CancellationToken ct = default);

    /// <summary>
    /// Adds a pivot table to <paramref name="sheetName"/> and returns the updated workbook.
    /// </summary>
    /// <remarks>
    /// <b>The result grid is empty until Excel opens and recalculates it.</b> A pivot table's
    /// aggregated values are computed by whichever application opens the file — nothing that
    /// writes it (this method included) populates the grid. Open the result in Excel (or an
    /// equivalent) to see it populated.
    /// </remarks>
    /// <param name="xlsx">The workbook to add the pivot table to. It is not modified.</param>
    /// <param name="sheetName">The sheet to add the pivot table to.</param>
    /// <param name="sourceRange">An A1-style range naming the source data, e.g. <c>"A1:C10"</c>.</param>
    /// <param name="destinationCell">
    /// An A1-style cell reference for the pivot table's top-left corner, e.g. <c>"E1"</c>.
    /// </param>
    /// <param name="name">The pivot table's name.</param>
    /// <param name="rowFields">Source column headers to group by, down the rows. At least one.</param>
    /// <param name="dataFields">The aggregated value columns. At least one.</param>
    /// <param name="columnFields">Source column headers to group by, across the columns. Optional.</param>
    /// <param name="pageFields">Source column headers used as report filters. Optional.</param>
    /// <param name="showRowGrandTotals">Whether to show a grand total row.</param>
    /// <param name="showColumnGrandTotals">Whether to show a grand total column.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="xlsx"/>, <paramref name="rowFields"/>, <paramref name="dataFields"/> or
    /// another required argument is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="xlsx"/> is empty, a name argument is blank, or
    /// <paramref name="rowFields"/>/<paramref name="dataFields"/> is empty.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or
    /// <paramref name="destinationCell"/> is not a valid cell reference.
    /// </exception>
    byte[] AddPivotTable(
        byte[] xlsx, string sheetName, string sourceRange, string destinationCell, string name,
        IEnumerable<string> rowFields, IEnumerable<DocToolkit.PivotDataField> dataFields,
        IEnumerable<string>? columnFields = null, IEnumerable<string>? pageFields = null,
        bool showRowGrandTotals = true, bool showColumnGrandTotals = true);

    /// <summary>
    /// Reads a workbook from <paramref name="source"/>, adds a pivot table, and writes the
    /// result to <paramref name="destination"/> — see <see cref="AddPivotTable"/> for the
    /// parameters.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought, and neither has to be seekable.
    /// </summary>
    /// <param name="source">The stream the workbook is read from.</param>
    /// <param name="sheetName">The sheet to add the pivot table to.</param>
    /// <param name="sourceRange">An A1-style range naming the source data, e.g. <c>"A1:C10"</c>.</param>
    /// <param name="destinationCell">
    /// An A1-style cell reference for the pivot table's top-left corner, e.g. <c>"E1"</c>.
    /// </param>
    /// <param name="name">The pivot table's name.</param>
    /// <param name="rowFields">Source column headers to group by, down the rows. At least one.</param>
    /// <param name="dataFields">The aggregated value columns. At least one.</param>
    /// <param name="destination">The stream the updated workbook is written to.</param>
    /// <param name="columnFields">Source column headers to group by, across the columns. Optional.</param>
    /// <param name="pageFields">Source column headers used as report filters. Optional.</param>
    /// <param name="showRowGrandTotals">Whether to show a grand total row.</param>
    /// <param name="showColumnGrandTotals">Whether to show a grand total column.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="source"/>, <paramref name="destination"/>, <paramref name="rowFields"/> or
    /// <paramref name="dataFields"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, a name is blank, or
    /// <paramref name="rowFields"/>/<paramref name="dataFields"/> is empty.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or
    /// <paramref name="destinationCell"/> is not a valid cell reference.
    /// </exception>
    Task AddPivotTableAsync(
        Stream source, string sheetName, string sourceRange, string destinationCell, string name,
        IEnumerable<string> rowFields, IEnumerable<DocToolkit.PivotDataField> dataFields, Stream destination,
        IEnumerable<string>? columnFields = null, IEnumerable<string>? pageFields = null,
        bool showRowGrandTotals = true, bool showColumnGrandTotals = true, CancellationToken ct = default);
}
