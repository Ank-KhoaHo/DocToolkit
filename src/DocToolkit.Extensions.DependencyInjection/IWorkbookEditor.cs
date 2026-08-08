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
}
