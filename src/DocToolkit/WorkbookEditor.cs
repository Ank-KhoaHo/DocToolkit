using System.Globalization;
using ClosedXML.Excel;

namespace DocToolkit;

/// <summary>Creates, reads and edits Excel (.xlsx) workbooks. Legacy .xls is not supported.</summary>
public static class WorkbookEditor
{
    /// <summary>
    /// Creates a workbook with one sheet populated from <paramref name="rows"/>.
    ///
    /// Every built-in numeric type is written as a number so formulas such as SUM() pick it up;
    /// <see cref="DateTime"/> and <see cref="DateOnly"/> become dates, <see cref="TimeOnly"/> and
    /// <see cref="TimeSpan"/> become durations. Anything else is formatted with
    /// <see cref="CultureInfo.InvariantCulture"/>, so the same code produces the same spreadsheet
    /// on every machine.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="rows"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="sheetName"/> is blank, or a row is null.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be built.</exception>
    public static byte[] Create(string sheetName, IEnumerable<IEnumerable<object?>> rows)
    {
        var materialised = ValidateRows(sheetName, rows);
        using var ms = CreateCore(sheetName, materialised);
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a workbook with one sheet populated from <paramref name="rows"/> and writes it to
    /// <paramref name="destination"/>. See <see cref="Create"/> for the exact typing and culture
    /// rules applied to each cell — this overload applies the identical logic, writing to
    /// <paramref name="destination"/> instead of returning an array.
    ///
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought — it belongs to the caller, and may be write-only and
    /// forward-only, such as an HTTP response body.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to create.</param>
    /// <param name="rows">The rows to populate it with.</param>
    /// <param name="destination">The stream the workbook is written to.</param>
    /// <param name="ct">Cancels the build and the write to <paramref name="destination"/>.</param>
    /// <exception cref="ArgumentNullException"><paramref name="rows"/> or <paramref name="destination"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="sheetName"/> is blank, a row is null, or <paramref name="destination"/> is
    /// not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be built or written.</exception>
    public static async Task CreateAsync(
        string sheetName, IEnumerable<IEnumerable<object?>> rows, Stream destination,
        CancellationToken ct = default)
    {
        var materialised = ValidateRows(sheetName, rows);
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var ms = CreateCore(sheetName, materialised);
        await StreamPipeline.EmitAsync(ms, destination, "Failed to create XLSX.", ct).ConfigureAwait(false);
    }

    private static List<IEnumerable<object?>> ValidateRows(string sheetName, IEnumerable<IEnumerable<object?>> rows)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentNullException.ThrowIfNull(rows);

        // Validated up front so a null row surfaces as the ArgumentException it is rather than as
        // a NullReferenceException wrapped in a conversion failure.
        return rows
            .Select((row, index) => row
                ?? throw new ArgumentException($"Row {index + 1} was null.", nameof(rows)))
            .ToList();
    }

    private static MemoryStream CreateCore(string sheetName, List<IEnumerable<object?>> rows)
    {
        try
        {
            using var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add(sheetName);

            var r = 1;
            foreach (var row in rows)
            {
                var c = 1;
                foreach (var value in row)
                    SetCellValue(sheet.Cell(r, c++), value);
                r++;
            }

            var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to create XLSX.", ex);
        }
    }

    /// <summary>
    /// Lists every sheet in the workbook, in tab order, including hidden sheets — hiding a sheet
    /// is a presentation choice, not a privacy boundary, and a caller who cannot see a hidden sheet
    /// listed has no way to discover it exists.
    /// </summary>
    /// <param name="xlsx">The workbook bytes.</param>
    /// <returns>The sheet names, in tab order.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be opened.</exception>
    public static IReadOnlyList<string> SheetNames(byte[] xlsx)
    {
        ValidateWorkbook(xlsx);

        try
        {
            using var workbook = Open(xlsx);
            return SheetNamesCore(workbook);
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read XLSX.", ex);
        }
    }

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
    /// <exception cref="DocumentConversionException">The workbook could not be opened.</exception>
    public static async Task<IReadOnlyList<string>> SheetNamesAsync(
        Stream source, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var xlsx = await StreamPipeline
            .DrainAsync(source, "Workbook content was empty.", nameof(source), "Failed to read XLSX.", ct)
            .ConfigureAwait(false);

        try
        {
            using var workbook = new XLWorkbook(xlsx);
            return SheetNamesCore(workbook);
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read XLSX.", ex);
        }
    }

    // Ordered by Position explicitly: Position is the tab order, whereas the enumeration order of
    // Worksheets is not documented to be. Sorting makes the guarantee true by construction.
    private static List<string> SheetNamesCore(XLWorkbook workbook)
        => workbook.Worksheets.OrderBy(sheet => sheet.Position).Select(sheet => sheet.Name).ToList();

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
    /// — the same rule <see cref="ReadCell"/> uses, and asymmetric with <see cref="Create"/>, which
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
    /// <see cref="DocumentConversionException"/> rather than allocate when the used range exceeds
    /// 2,000,000 cells.
    /// </remarks>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="xlsx"/> is empty, or <paramref name="sheetName"/> is blank.
    /// </exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the sheet's used range
    /// exceeds the 2,000,000-cell limit <see cref="ReadSheet"/> will materialise.
    /// </exception>
    public static IReadOnlyList<IReadOnlyList<string>> ReadSheet(byte[] xlsx, string sheetName)
    {
        ValidateWorkbook(xlsx);
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);

        try
        {
            using var workbook = Open(xlsx);
            return ReadSheetCore(workbook, sheetName);
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read XLSX.", ex);
        }
    }

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
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the sheet's used range
    /// exceeds the 2,000,000-cell limit <see cref="ReadSheet"/> will materialise.
    /// </exception>
    public static async Task<IReadOnlyList<IReadOnlyList<string>>> ReadSheetAsync(
        Stream source, string sheetName, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var xlsx = await StreamPipeline
            .DrainAsync(source, "Workbook content was empty.", nameof(source), "Failed to read XLSX.", ct)
            .ConfigureAwait(false);

        try
        {
            using var workbook = new XLWorkbook(xlsx);
            return ReadSheetCore(workbook, sheetName);
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read XLSX.", ex);
        }
    }

    // Not part of any spec: the design only said "read a sheet without knowing its shape in
    // advance". ReadSheetCore materialises the whole rows x columns rectangle up front, so its
    // memory cost tracks the *rectangle*, not how much of it actually holds data — a single stray
    // value (or even just a cell comment; see the LastCellUsed() note below) way out at
    // XFD1048576, Excel's own maximum address, describes a 1,048,576 x 16,384 grid, ~17.2 billion
    // string-array slots, from a workbook that can be a few KB on disk. 2,000,000 is chosen as
    // comfortably above any sheet a caller would actually want back as an in-memory jagged array,
    // while still catching that case before a single byte of it is allocated.
    private const long ReadSheetCellLimit = 2_000_000;

    private static List<IReadOnlyList<string>> ReadSheetCore(XLWorkbook workbook, string sheetName)
    {
        var sheet = Sheet(workbook, sheetName);

        // The extent comes from LastCellUsed() rather than LastRowUsed()/LastColumnUsed(): those
        // return range rows/columns whose RowNumber()/ColumnNumber() are documented as positions
        // *within the range*, which is an off-by-origin waiting to happen. A cell's Address is
        // absolute. LastCellUsed() ignores formatting, so one bolded empty cell out at Z1 cannot
        // pad every row out to it — but it does count a cell comment as "used" even with no value,
        // so a comment out at the far corner widens the range exactly like a stray value would.
        // Null means the sheet holds no values and no comments at all.
        var last = sheet.LastCellUsed();
        if (last is null)
            return new List<IReadOnlyList<string>>();

        var lastRow = last.Address.RowNumber;
        var lastColumn = last.Address.ColumnNumber;

        // long, not int: lastRow * lastColumn as plain int arithmetic overflows (wraps, possibly
        // negative) well before it reaches Excel's real maximum of ~17.2 billion, which would
        // silently defeat this exact check.
        var cellCount = (long)lastRow * lastColumn;
        if (cellCount > ReadSheetCellLimit)
        {
            throw new DocumentConversionException(
                $"Sheet '{sheetName}' spans {lastRow} rows x {lastColumn} columns ({cellCount} " +
                $"cells), which exceeds the {ReadSheetCellLimit}-cell limit ReadSheet will materialise.");
        }

        // From row 1 and column 1, not from the first used cell: the result is anchored at A1 so
        // rows[r][c] addresses the sheet the way the caller sees it in Excel.
        var rows = new List<IReadOnlyList<string>>(lastRow);
        for (var r = 1; r <= lastRow; r++)
        {
            var row = new string[lastColumn];
            for (var c = 1; c <= lastColumn; c++)
                row[c - 1] = sheet.Cell(r, c).GetString();
            rows.Add(row);
        }

        return rows;
    }

    /// <summary>
    /// Reads a cell as a string. <paramref name="cellRef"/> is an A1-style reference. Text follows
    /// the calling thread's <see cref="System.Globalization.CultureInfo.CurrentCulture"/> — see
    /// <see cref="ReadSheet"/> for the full rule.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty, or a name is blank.</exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    public static string ReadCell(byte[] xlsx, string sheetName, string cellRef)
    {
        ValidateArguments(xlsx, sheetName, cellRef);

        try
        {
            using var workbook = Open(xlsx);
            return Sheet(workbook, sheetName).Cell(cellRef).GetString();
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read XLSX.", ex);
        }
    }

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
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    public static async Task<string> ReadCellAsync(
        Stream source, string sheetName, string cellRef, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellRef);
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var xlsx = await StreamPipeline
            .DrainAsync(source, "Workbook content was empty.", nameof(source), "Failed to read XLSX.", ct)
            .ConfigureAwait(false);

        try
        {
            using var workbook = new XLWorkbook(xlsx);
            return Sheet(workbook, sheetName).Cell(cellRef).GetString();
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read XLSX.", ex);
        }
    }

    /// <summary>Sets a cell and returns the updated workbook bytes.</summary>
    /// <exception cref="ArgumentNullException">Any argument other than <paramref name="value"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty, or a name is blank.</exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    public static byte[] SetCell(byte[] xlsx, string sheetName, string cellRef, object? value)
    {
        ValidateArguments(xlsx, sheetName, cellRef);

        try
        {
            using var workbook = Open(xlsx);
            SetCellValue(Sheet(workbook, sheetName).Cell(cellRef), value);

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to edit XLSX.", ex);
        }
    }

    /// <summary>
    /// Reads a workbook from <paramref name="source"/>, sets one cell, and writes the result to
    /// <paramref name="destination"/>. <paramref name="cellRef"/> is an A1-style reference.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought, and neither has to be seekable.
    /// </summary>
    /// <param name="source">The stream the workbook is read from.</param>
    /// <param name="sheetName">The sheet containing the cell.</param>
    /// <param name="cellRef">An A1-style cell reference, e.g. <c>"B2"</c>.</param>
    /// <param name="value">The value to write. <c>null</c> clears the cell.</param>
    /// <param name="destination">The stream the updated workbook is written to.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="source"/> or <paramref name="destination"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, a name is blank, or
    /// <paramref name="destination"/> is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    public static async Task SetCellAsync(
        Stream source, string sheetName, string cellRef, object? value, Stream destination,
        CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellRef);
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var xlsx = await StreamPipeline
            .DrainAsync(source, "Workbook content was empty.", nameof(source), "Failed to edit XLSX.", ct)
            .ConfigureAwait(false);

        using var result = SetCellCore(xlsx, sheetName, cellRef, value);
        await StreamPipeline.EmitAsync(result, destination, "Failed to edit XLSX.", ct).ConfigureAwait(false);
    }

    private static MemoryStream SetCellCore(Stream xlsx, string sheetName, string cellRef, object? value)
    {
        try
        {
            using var workbook = new XLWorkbook(xlsx);
            SetCellValue(Sheet(workbook, sheetName).Cell(cellRef), value);

            var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to edit XLSX.", ex);
        }
    }

    private static void ValidateArguments(byte[] xlsx, string sheetName, string cellRef)
    {
        ArgumentNullException.ThrowIfNull(xlsx);
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellRef);
        if (xlsx.Length == 0)
            throw new ArgumentException("Workbook content was empty.", nameof(xlsx));
    }

    private static void ValidateWorkbook(byte[] xlsx)
    {
        ArgumentNullException.ThrowIfNull(xlsx);
        if (xlsx.Length == 0)
            throw new ArgumentException("Workbook content was empty.", nameof(xlsx));
    }

    private static XLWorkbook Open(byte[] xlsx)
    {
        var ms = new MemoryStream();
        ms.Write(xlsx, 0, xlsx.Length);
        ms.Position = 0;
        return new XLWorkbook(ms);
    }

    private static IXLWorksheet Sheet(XLWorkbook workbook, string sheetName)
    {
        if (!workbook.Worksheets.TryGetWorksheet(sheetName, out var sheet))
            throw new DocumentConversionException($"Worksheet '{sheetName}' was not found.");
        return sheet;
    }

    private static void SetCellValue(IXLCell cell, object? value)
    {
        switch (value)
        {
            case null: cell.Clear(XLClearOptions.Contents); break;
            case string s: cell.Value = s; break;
            case bool b: cell.Value = b; break;
            case DateTime d: cell.Value = d; break;
            case DateOnly d: cell.Value = d.ToDateTime(TimeOnly.MinValue); break;
            case TimeOnly t: cell.Value = t.ToTimeSpan(); break;
            case TimeSpan ts: cell.Value = ts; break;

            // The unsigned types and sbyte used to fall through to ToString() and land as text,
            // which silently excluded them from SUM() and the rest of Excel's numeric functions.
            case sbyte or byte or short or ushort or int or uint or long or ulong
                 or float or double or decimal:
                cell.Value = Convert.ToDouble(value, CultureInfo.InvariantCulture); break;

            // Invariant, not CurrentCulture: otherwise the same code writes a different
            // spreadsheet depending on the machine's regional settings.
            default: cell.Value = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty; break;
        }
    }

    /// <summary>
    /// Builds a workbook with one sheet populated from <paramref name="rows"/> and writes it to
    /// <paramref name="outputPath"/>. See <see cref="Create"/> for the exact typing and culture
    /// rules applied to each cell — this overload applies the identical logic, writing to
    /// <paramref name="outputPath"/> instead of returning an array.
    ///
    /// Named <c>CreateToFileAsync</c> rather than a <c>CreateAsync</c> overload: <see cref="CreateAsync"/>
    /// already starts with <c>string sheetName</c>, so a plain overload taking <c>string outputPath</c>
    /// first would leave two <c>CreateAsync</c> methods whose first argument means different things.
    /// </summary>
    /// <param name="outputPath">Where to write the workbook. Overwritten if it exists.</param>
    /// <param name="sheetName">The name of the sheet to create.</param>
    /// <param name="rows">The rows to populate it with.</param>
    /// <param name="ct">Cancels the build and the write to <paramref name="outputPath"/>.</param>
    /// <exception cref="ArgumentNullException"><paramref name="rows"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="outputPath"/> or <paramref name="sheetName"/> is blank, or a row is null.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="outputPath"/>'s directory does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be built.</exception>
    public static async Task CreateToFileAsync(
        string outputPath, string sheetName, IEnumerable<IEnumerable<object?>> rows,
        CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var bytes = Create(sheetName, rows);
        await File.WriteAllBytesAsync(outputPath, bytes, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads a workbook from <paramref name="inputPath"/>, sets one cell, and writes the result to
    /// <paramref name="outputPath"/>. <paramref name="cellRef"/> is an A1-style reference. The two
    /// paths may be the same file: the input is read completely before the output is written, so
    /// editing in place is safe, and a workbook that fails to process leaves whatever was at
    /// <paramref name="outputPath"/> untouched.
    /// </summary>
    /// <param name="inputPath">The workbook to read.</param>
    /// <param name="outputPath">Where to write the result. Overwritten if it exists.</param>
    /// <param name="sheetName">The sheet containing the cell.</param>
    /// <param name="cellRef">An A1-style cell reference, e.g. <c>"B2"</c>.</param>
    /// <param name="value">The value to write. <c>null</c> clears the cell.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">A path or a name is null.</exception>
    /// <exception cref="ArgumentException">A path or a name is blank.</exception>
    /// <exception cref="FileNotFoundException"><paramref name="inputPath"/> does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    public static async Task SetCellAsync(
        string inputPath, string outputPath, string sheetName, string cellRef, object? value,
        CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var bytes = await File.ReadAllBytesAsync(inputPath, ct).ConfigureAwait(false);
        var result = SetCell(bytes, sheetName, cellRef, value);
        await File.WriteAllBytesAsync(outputPath, result, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads a workbook from <paramref name="path"/> and returns a cell as a string.
    /// <paramref name="cellRef"/> is an A1-style reference. See <see cref="ReadCell"/> for the
    /// culture rule applied to the text.
    /// </summary>
    /// <param name="path">The workbook to read.</param>
    /// <param name="sheetName">The sheet containing the cell.</param>
    /// <param name="cellRef">An A1-style cell reference, e.g. <c>"B2"</c>.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <returns>The cell's value as a string.</returns>
    /// <exception cref="ArgumentNullException">A name is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="path"/> or a name is blank.</exception>
    /// <exception cref="FileNotFoundException"><paramref name="path"/> does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    public static async Task<string> ReadCellAsync(
        string path, string sheetName, string cellRef, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        var bytes = await File.ReadAllBytesAsync(path, ct).ConfigureAwait(false);
        return ReadCell(bytes, sheetName, cellRef);
    }

    /// <summary>
    /// Reads a workbook from <paramref name="path"/> and lists every sheet in tab order, including
    /// hidden sheets. See <see cref="SheetNames"/> for the full rule.
    /// </summary>
    /// <param name="path">The workbook to read.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <returns>The sheet names, in tab order.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="path"/> is blank.</exception>
    /// <exception cref="FileNotFoundException"><paramref name="path"/> does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be opened.</exception>
    public static async Task<IReadOnlyList<string>> SheetNamesAsync(
        string path, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        var bytes = await File.ReadAllBytesAsync(path, ct).ConfigureAwait(false);
        return SheetNames(bytes);
    }

    /// <summary>
    /// Reads a workbook from <paramref name="path"/> and returns a whole sheet as strings. See
    /// <see cref="ReadSheet"/> for the anchoring, padding, culture and formula rules — this
    /// overload applies the identical logic.
    /// </summary>
    /// <param name="path">The workbook to read.</param>
    /// <param name="sheetName">The sheet to read.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <returns>
    /// The sheet's used range, anchored at A1 and padded rectangular. Empty only if the sheet
    /// holds no values and no cell comments — see <see cref="ReadSheet"/>.
    /// </returns>
    /// <exception cref="ArgumentNullException"><paramref name="sheetName"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="path"/> or <paramref name="sheetName"/> is blank.</exception>
    /// <exception cref="FileNotFoundException"><paramref name="path"/> does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the sheet's used range
    /// exceeds the 2,000,000-cell limit <see cref="ReadSheet"/> will materialise.
    /// </exception>
    public static async Task<IReadOnlyList<IReadOnlyList<string>>> ReadSheetAsync(
        string path, string sheetName, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        var bytes = await File.ReadAllBytesAsync(path, ct).ConfigureAwait(false);
        return ReadSheet(bytes, sheetName);
    }
}
