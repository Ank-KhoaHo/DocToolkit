using System.Globalization;
using ClosedXML.Excel;

using OfficeIMOExcelExcelDocument = OfficeIMO.Excel.ExcelDocument;
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
    /// <exception cref="ArgumentException">
    /// <paramref name="sheetName"/> is blank, is longer than 31 characters, or contains one of
    /// <c>: \ / ? * [ ]</c>; or a row is null.
    /// </exception>
    /// <exception cref="DocumentConversionException">The workbook could not be built.</exception>
    public static byte[] Create(string sheetName, IEnumerable<IEnumerable<object?>> rows)
    {
        var materialised = ValidateRows(sheetName, rows);
        using var ms = CreateCore(sheetName, materialised);
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a workbook with one sheet populated from <paramref name="rows"/> and writes it to
    /// <paramref name="destination"/>. See <see cref="Create(string, IEnumerable{IEnumerable{object}})"/>
    /// for the exact typing and culture rules applied to each cell — this overload applies the
    /// identical logic, writing to <paramref name="destination"/> instead of returning an array.
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
    /// <paramref name="sheetName"/> is blank, is longer than 31 characters, or contains one of
    /// <c>: \ / ? * [ ]</c>; a row is null; or <paramref name="destination"/> is not writable.
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
        await StreamPipeline.EmitAsync(ms, destination, "Failed to create XLSX. See the inner exception for details.", ct).ConfigureAwait(false);
    }

    // Excel's own rules. Enforced here rather than left to ClosedXML so an invalid name fails fast,
    // naming the parameter, instead of surfacing as a DocumentConversionException wrapping someone
    // else's message. One helper, used by every path that names a sheet, so the rules cannot drift.
    private static readonly char[] InvalidSheetNameChars = [':', '\\', '/', '?', '*', '[', ']'];

    internal static string ValidateSheetName(string sheetName, string paramName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName, paramName);

        if (sheetName.Length > 31)
            throw new ArgumentException(
                $"Sheet name '{sheetName}' is {sheetName.Length} characters; Excel allows at most 31.",
                paramName);

        var bad = sheetName.IndexOfAny(InvalidSheetNameChars);
        if (bad >= 0)
            throw new ArgumentException(
                $"Sheet name '{sheetName}' contains '{sheetName[bad]}'. Excel does not allow : \\ / ? * [ ] in a sheet name.",
                paramName);

        return sheetName;
    }

    private static List<XlsxSheet> ValidateSheets(IEnumerable<XlsxSheet> sheets)
    {
        ArgumentNullException.ThrowIfNull(sheets);

        var materialised = sheets
            .Select((sheet, index) => sheet
                ?? throw new ArgumentException($"Sheet {index + 1} was null.", nameof(sheets)))
            .ToList();

        if (materialised.Count == 0)
            throw new ArgumentException(
                "At least one sheet is required; a workbook with no worksheets is not a valid .xlsx.",
                nameof(sheets));

        // Excel compares sheet names case-insensitively, so "Data" and "data" collide.
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in materialised)
            if (!seen.Add(sheet.Name))
                throw new ArgumentException(
                    $"Sheet name '{sheet.Name}' appears more than once; Excel requires unique sheet names.",
                    nameof(sheets));

        return materialised;
    }

    private static MemoryStream CreateCore(List<XlsxSheet> sheets)
    {
        try
        {
            using var workbook = new XLWorkbook();

            foreach (var sheet in sheets)
            {
                var worksheet = workbook.Worksheets.Add(sheet.Name);

                var r = 1;
                foreach (var row in sheet.Rows)
                {
                    var c = 1;
                    foreach (var value in row)
                        SetCellValue(worksheet.Cell(r, c++), value);
                    r++;
                }
            }

            var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to create XLSX. See the inner exception for details.", ex);
        }
    }

    private static List<IEnumerable<object?>> ValidateRows(string sheetName, IEnumerable<IEnumerable<object?>> rows)
    {
        ValidateSheetName(sheetName, nameof(sheetName));
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
            throw new DocumentConversionException("Failed to create XLSX. See the inner exception for details.", ex);
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
            throw new DocumentConversionException("Failed to read XLSX. See the inner exception for details.", ex);
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
            .DrainAsync(source, "Workbook content was empty.", nameof(source), "Failed to read XLSX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        try
        {
            using var workbook = new XLWorkbook(xlsx);
            return SheetNamesCore(workbook);
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read XLSX. See the inner exception for details.", ex);
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
    /// disagree about what a cell says. A formula cell reads back one of two ways, and which one
    /// depends on the file rather than on this library:
    ///
    /// <list type="bullet">
    /// <item><description>
    /// <b>If the file carries a cached value</b>, as one Excel has saved does, that cached value is
    /// returned and the formula is <b>not</b> evaluated. It can therefore be stale — a workbook
    /// whose inputs were edited by something that did not recalculate reports the old result.
    /// </description></item>
    /// <item><description>
    /// <b>If it does not</b> — which is what this library writes, see <see cref="XlsxFormula"/> —
    /// ClosedXML evaluates the formula on read, so a cell holding <c>=A1+A2</c> over 1 and 2 reads
    /// back as <c>"3"</c>, and one that cannot be evaluated reads back as its Excel error string
    /// (<c>#DIV/0!</c>, <c>#NAME?</c>, <c>#REF!</c>) rather than throwing.
    /// </description></item>
    /// </list>
    ///
    /// Text follows the calling thread's <see cref="System.Globalization.CultureInfo.CurrentCulture"/>
    /// — the same rule <see cref="ReadCell"/> uses, and asymmetric with
    /// <see cref="Create(string, IEnumerable{IEnumerable{object}})"/>, which
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
    /// <example>
    /// <code source="../../tests/DocToolkit.Tests/DocumentationExamples.cs" region="WorkbookReadSheet"/>
    /// </example>
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
            throw new DocumentConversionException("Failed to read XLSX. See the inner exception for details.", ex);
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
            .DrainAsync(source, "Workbook content was empty.", nameof(source), "Failed to read XLSX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        try
        {
            using var workbook = new XLWorkbook(xlsx);
            return ReadSheetCore(workbook, sheetName);
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read XLSX. See the inner exception for details.", ex);
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

    private static List<IReadOnlyList<string>> ReadSheetCore(
        XLWorkbook workbook, string sheetName, bool invariant = false)
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
                $"cells), which exceeds the {ReadSheetCellLimit}-cell limit ReadSheet will " +
                "materialise. Read specific cells with ReadCell instead.");
        }

        // From row 1 and column 1, not from the first used cell: the result is anchored at A1 so
        // rows[r][c] addresses the sheet the way the caller sees it in Excel.
        var rows = new List<IReadOnlyList<string>>(lastRow);
        for (var r = 1; r <= lastRow; r++)
        {
            var row = new string[lastColumn];
            for (var c = 1; c <= lastColumn; c++)
                row[c - 1] = CellText(sheet.Cell(r, c), invariant);
            rows.Add(row);
        }

        return rows;
    }

    /// <summary>
    /// The same grid <see cref="ReadSheet(byte[], string)"/> returns, rendered
    /// <b>culture-invariantly</b>, for the exporters.
    /// </summary>
    /// <remarks>
    /// One grid reader with a flag rather than a second reader: two ways of turning a sheet into
    /// text is exactly the drift the <c>*Core</c> convention and <c>SetCellValue</c>'s single site
    /// exist to prevent.
    /// </remarks>
    internal static IReadOnlyList<IReadOnlyList<string>> ReadSheetInvariant(byte[] xlsx, string sheetName)
    {
        ValidateWorkbook(xlsx);
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);

        try
        {
            using var workbook = Open(xlsx);
            return ReadSheetCore(workbook, sheetName, invariant: true);
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read XLSX. See the inner exception for details.", ex);
        }
    }

    /// <summary>
    /// A cell as text, optionally independent of the machine's regional settings.
    /// </summary>
    /// <remarks>
    /// <b>Why the invariant form exists.</b> <c>GetString()</c> follows
    /// <see cref="CultureInfo.CurrentCulture"/>, so the same workbook reads back as <c>1234.5</c>
    /// on one machine and <c>1234,5</c> on another. That is tolerable for
    /// <see cref="ReadSheet(byte[], string)"/>, whose result a caller inspects — and **corrupting**
    /// for CSV, where a decimal comma collides with the delimiter itself. Measured 2026-08-13
    /// across en-US, de-DE and fr-FR.
    ///
    /// Invariant output is already this class's convention on the way in: <c>SetCellValue</c>
    /// writes invariantly so that "the same code writes a different spreadsheet depending on the
    /// machine's regional settings" cannot happen. The exporters simply hold the same line on the
    /// way out.
    ///
    /// <b>A date-only value renders as <c>yyyy-MM-dd</c>, not with a midnight time.</b> Excel has
    /// no date-without-time type, so every date carries 00:00:00; emitting it produces
    /// <c>2026-08-13 00:00:00</c> for what the author entered as a date. A time component is kept
    /// only when there is one.
    ///
    /// A formula cell renders its computed VALUE, which is what both exporters want — the formula
    /// text is a spreadsheet concern, not a CSV or HTML one.
    /// </remarks>
    private static string CellText(IXLCell cell, bool invariant)
    {
        if (!invariant) return cell.GetString();

        var value = cell.Value;

        if (value.IsDateTime)
        {
            var d = value.GetDateTime();
            return d.TimeOfDay == TimeSpan.Zero
                ? d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                : d.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        if (value.IsNumber)
            return value.GetNumber().ToString(CultureInfo.InvariantCulture);

        if (value.IsBoolean)
            return value.GetBoolean() ? "TRUE" : "FALSE";

        if (value.IsTimeSpan)
            return value.GetTimeSpan().ToString("c", CultureInfo.InvariantCulture);

        return cell.GetString();
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
            throw new DocumentConversionException("Failed to read XLSX. See the inner exception for details.", ex);
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
            .DrainAsync(source, "Workbook content was empty.", nameof(source), "Failed to read XLSX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        try
        {
            using var workbook = new XLWorkbook(xlsx);
            return Sheet(workbook, sheetName).Cell(cellRef).GetString();
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read XLSX. See the inner exception for details.", ex);
        }
    }

    /// <summary>
    /// Sets a cell and returns the updated workbook bytes. A cell holding an
    /// <see cref="XlsxFormula"/> is written as a formula instead of a literal value — see that
    /// type for the one limit worth knowing about cached values.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument other than <paramref name="value"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty, or a name is blank.</exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    public static byte[] SetCell(byte[] xlsx, string sheetName, string cellRef, object? value)
    {
        ValidateArguments(xlsx, sheetName, cellRef);

        using var source = new MemoryStream(xlsx, writable: false);
        using var result = SetCellCore(source, sheetName, cellRef, value);
        return result.ToArray();
    }

    /// <summary>
    /// Reads a workbook from <paramref name="source"/>, sets one cell, and writes the result to
    /// <paramref name="destination"/>. <paramref name="cellRef"/> is an A1-style reference. A cell
    /// holding an <see cref="XlsxFormula"/> is written as a formula instead of a literal value —
    /// see that type for the one limit worth knowing about cached values.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought, and neither has to be seekable.
    /// </summary>
    /// <param name="source">The stream the workbook is read from.</param>
    /// <param name="sheetName">The sheet containing the cell.</param>
    /// <param name="cellRef">An A1-style cell reference, e.g. <c>"B2"</c>.</param>
    /// <param name="value">
    /// The value to write. <c>null</c> clears the cell; an <see cref="XlsxFormula"/> writes a
    /// formula.
    /// </param>
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
            .DrainAsync(source, "Workbook content was empty.", nameof(source), "Failed to edit XLSX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        using var result = SetCellCore(xlsx, sheetName, cellRef, value);
        await StreamPipeline.EmitAsync(result, destination, "Failed to edit XLSX. See the inner exception for details.", ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Applies <paramref name="format"/> to <paramref name="sheetName"/> and returns the workbook.
    /// </summary>
    /// <remarks>
    /// Formatting is applied to an existing workbook rather than being an argument to
    /// <c>Create</c>, so it composes with every way a workbook can arrive here — built by
    /// <see cref="Create(string, IEnumerable{IEnumerable{object}})"/>, appended to, or handed in by
    /// a caller who never used this library to make it.
    ///
    /// See <see cref="XlsxFormat"/> for the boundary: a CLOSED vocabulary rather than a small
    /// one, and what is still deliberately outside it.
    /// </remarks>
    /// <param name="xlsx">The workbook to format. It is not modified.</param>
    /// <param name="sheetName">The sheet to format.</param>
    /// <param name="format">The formatting to apply.</param>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> or <paramref name="format"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="xlsx"/> is empty, or <paramref name="sheetName"/> is blank.
    /// </exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or a rule or validation named a
    /// range the sheet rejects — a malformed one such as <c>"B2:B1"</c>, or a column letter beyond
    /// the sheet's width. Ranges are checked by the library beneath rather than here, deliberately:
    /// a second range parser would be a second source of truth about what a range is.
    /// </exception>
    public static byte[] Format(byte[] xlsx, string sheetName, XlsxFormat format)
    {
        ValidateWorkbook(xlsx);
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentNullException.ThrowIfNull(format);

        using var source = new MemoryStream(xlsx, writable: false);
        using var result = FormatCore(source, sheetName, format);
        return result.ToArray();
    }

    /// <summary>
    /// Reads a workbook from <paramref name="source"/>, applies <paramref name="format"/> to
    /// <paramref name="sheetName"/>, and writes the result to <paramref name="destination"/>.
    ///
    /// Neither stream is disposed, closed or sought.
    /// </summary>
    /// <param name="source">The stream the workbook is read from.</param>
    /// <param name="sheetName">The sheet to format.</param>
    /// <param name="format">The formatting to apply.</param>
    /// <param name="destination">The stream the workbook is written to.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    /// <exception cref="ArgumentNullException"><paramref name="format"/> is null, or a stream is null.</exception>
    /// <exception cref="ArgumentException">
    /// A stream is unusable, <paramref name="source"/> held no bytes, or <paramref name="sheetName"/>
    /// is blank.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, or the sheet does not exist.
    /// </exception>
    public static async Task FormatAsync(
        Stream source, string sheetName, XlsxFormat format, Stream destination,
        CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentNullException.ThrowIfNull(format);
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var xlsx = await StreamPipeline
            .DrainAsync(source, "Workbook content was empty.", nameof(source), "Failed to edit XLSX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        using var result = FormatCore(xlsx, sheetName, format);
        await StreamPipeline.EmitAsync(result, destination, "Failed to edit XLSX. See the inner exception for details.", ct).ConfigureAwait(false);
    }

    private static MemoryStream FormatCore(Stream xlsx, string sheetName, XlsxFormat format)
    {
        try
        {
            using var workbook = new XLWorkbook(xlsx);
            ApplyFormat(Sheet(workbook, sheetName), format);

            var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to edit XLSX. See the inner exception for details.", ex);
        }
    }

    /// <summary>
    /// The one place an <see cref="XlsxFormat"/> becomes spreadsheet styling, so the
    /// <c>byte[]</c> and <c>Stream</c> paths cannot disagree about what a format means — the same
    /// rule <see cref="SetCellValue"/> and <c>SectionPropertiesFactory</c> follow.
    /// </summary>
    private static void ApplyFormat(IXLWorksheet sheet, XlsxFormat format)
    {
        // Number formats first: they apply to whole columns, and doing them before AutoFit means
        // the widths account for the formatted text rather than the raw value. "1234.5" and
        // "1,234.50" are different widths.
        foreach (var (column, numberFormat) in format.ColumnNumberFormats)
            sheet.Column(column).Style.NumberFormat.Format = numberFormat;

        if (format.BoldHeaderRow)
            sheet.Row(1).Style.Font.Bold = true;

        if (format.AutoFitColumns)
            sheet.Columns().AdjustToContents();

        // AFTER AutoFit, so a named column takes the width the caller asked for while the rest stay
        // auto-fitted. A specific instruction beats a blanket one, and the ordering is asserted.
        foreach (var (column, width) in format.ColumnWidths)
            sheet.Column(column).Width = width;

        // Freezing splits at the TOP of the row after the frozen ones, which is what "freeze the
        // header" means for (1, 0). Applied after AutoFit because adjusting a frozen pane's columns
        // is the sort of interaction worth not relying on.
        if (format.FreezeAt is { } freeze)
        {
            if (freeze.Row > 0) sheet.SheetView.FreezeRows(freeze.Row);
            if (freeze.Column > 0) sheet.SheetView.FreezeColumns(freeze.Column);
        }

        // RangeUsed() is null on an empty sheet, and SetAutoFilter on nothing would throw - so a
        // caller asking for a filter on a sheet with no data gets a sheet with no filter rather
        // than an exception.
        if (format.AutoFilter && sheet.RangeUsed() is { } used)
            used.SetAutoFilter();

        foreach (XlsxRule rule in format.Rules)
            ApplyRule(sheet, rule);

        foreach (XlsxValidation validation in format.Validations)
            ApplyValidation(sheet, validation);
    }

    /// <summary>
    /// The one place a rule becomes a conditional format.
    /// </summary>
    /// <remarks>
    /// Kept here beside <see cref="ApplyFormat"/> rather than on <see cref="XlsxRule"/>, so no
    /// ClosedXML type appears in this library's public API — the same reason
    /// <see cref="XlsxHighlight"/> names an intent instead of carrying a colour.
    /// </remarks>
    private static void ApplyRule(IXLWorksheet sheet, XlsxRule rule)
    {
        // `var`, not a named type: what the When... methods return is ClosedXML's business, and
        // naming it here would be a guess that compiles until it does not. Text is dereferenced
        // with `!` because the two kinds that read it are the two whose factories require it.
        var style = rule.Kind switch
        {
            XlsxRuleKind.GreaterThan => sheet.Range(rule.Range).AddConditionalFormat().WhenGreaterThan(rule.Value),
            XlsxRuleKind.LessThan => sheet.Range(rule.Range).AddConditionalFormat().WhenLessThan(rule.Value),
            XlsxRuleKind.Between => sheet.Range(rule.Range).AddConditionalFormat().WhenBetween(rule.Value, rule.High),
            XlsxRuleKind.EqualTo => sheet.Range(rule.Range).AddConditionalFormat().WhenEquals(rule.Text!),
            XlsxRuleKind.Contains => sheet.Range(rule.Range).AddConditionalFormat().WhenContains(rule.Text!),
            XlsxRuleKind.Blank => sheet.Range(rule.Range).AddConditionalFormat().WhenIsBlank(),

            // NOT a fall-through arm. C# lets any int be cast to an enum, so (XlsxRuleKind)99
            // reaches here - and a `_` arm silently turned it into a Blank rule. Measured. For a
            // type whose whole premise is a CLOSED vocabulary, quietly answering a value outside
            // that vocabulary is the one thing it must not do.
            _ => throw new ArgumentOutOfRangeException(
                nameof(rule), rule.Kind,
                "Not a defined XlsxRuleKind. The vocabulary is closed; see XlsxRule's factories."),
        };

        style.Fill.SetBackgroundColor(Colour(rule.Highlight));
    }

    /// <summary>
    /// Four intents to four colours, deliberately not a caller-supplied one — see
    /// <see cref="XlsxHighlight"/>. They must stay DISTINCT: a mapping that collapsed two onto one
    /// colour would make two intents indistinguishable on the page, and a test asserts the set.
    /// </summary>
    private static XLColor Colour(XlsxHighlight highlight) => highlight switch
    {
        XlsxHighlight.Red => XLColor.Red,
        XlsxHighlight.Amber => XLColor.Orange,
        XlsxHighlight.Green => XLColor.LightGreen,
        XlsxHighlight.Grey => XLColor.LightGray,
        _ => throw new ArgumentOutOfRangeException(
            nameof(highlight), highlight,
            "Not a defined XlsxHighlight. Measured: (XlsxHighlight)99 used to come back grey, "
            + "which makes an out-of-range cast indistinguishable from a deliberate Grey."),
    };

    /// <summary>The one place a validation becomes a data validation.</summary>
    private static void ApplyValidation(IXLWorksheet sheet, XlsxValidation validation)
    {
        IXLDataValidation target = sheet.Range(validation.Range).CreateDataValidation();

        switch (validation.Kind)
        {
            case XlsxValidationKind.WholeNumber:
                target.WholeNumber.Between((int)validation.Min, (int)validation.Max);
                break;
            case XlsxValidationKind.Decimal:
                target.Decimal.Between(validation.Min, validation.Max);
                break;
            case XlsxValidationKind.TextLength:
                target.TextLength.Between((int)validation.Min, (int)validation.Max);
                break;
            case XlsxValidationKind.Date:
                target.Date.Between(validation.MinDate, validation.MaxDate);
                break;
            default:
                // An inline list is quoted and comma-joined, which is the form the file expects.
                target.List($"\"{string.Join(",", validation.Options)}\"");
                break;
        }
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
            throw new DocumentConversionException("Failed to edit XLSX. See the inner exception for details.", ex);
        }
    }

    private static MemoryStream AppendRowsCore(
        Stream xlsx, string sheetName, List<IEnumerable<object?>> rows)
    {
        try
        {
            using var workbook = new XLWorkbook(xlsx);
            var sheet = Sheet(workbook, sheetName);

            // LastRowUsed() is null for an empty sheet, in which case appending starts at row 1.
            var r = sheet.LastRowUsed()?.RowNumber() ?? 0;

            foreach (var row in rows)
            {
                r++;
                var c = 1;
                foreach (var value in row)
                    SetCellValue(sheet.Cell(r, c++), value);
            }

            var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to edit XLSX. See the inner exception for details.", ex);
        }
    }

    /// <summary>
    /// Appends <paramref name="rows"/> to <paramref name="sheetName"/>, starting immediately after
    /// its last used row, and returns the updated workbook. Every other sheet, and all existing
    /// formatting, is left as it was.
    ///
    /// <para>"Last used row" comes from ClosedXML's <c>LastRowUsed()</c>, which — like
    /// <c>LastCellUsed()</c> in <see cref="ReadSheet"/> — ignores formatting but counts a cell
    /// comment as used even with no value. A comment on an otherwise-blank row far below the real
    /// data therefore pushes the append down to start after that row, leaving a gap rather than
    /// continuing immediately after the last row a caller would see as holding data.</para>
    ///
    /// <para>Cell typing and culture rules are identical to
    /// <see cref="Create(string, IEnumerable{IEnumerable{object}})"/>. A cell holding an
    /// <see cref="XlsxFormula"/> is written as a formula. An empty <paramref name="rows"/> is a
    /// no-op that still returns a valid workbook.</para>
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> or <paramref name="rows"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="xlsx"/> is empty; <paramref name="sheetName"/> is blank, is longer than 31
    /// characters, or contains one of <c>: \ / ? * [ ]</c>; or an element of <paramref name="rows"/>
    /// is null.
    /// </exception>
    /// <exception cref="DocumentConversionException">The sheet was not found, or the package could not be opened or edited.</exception>
    public static byte[] AppendRows(byte[] xlsx, string sheetName, IEnumerable<IEnumerable<object?>> rows)
    {
        ValidateWorkbook(xlsx);
        var materialised = ValidateRows(sheetName, rows);

        using var source = new MemoryStream(xlsx, writable: false);
        using var ms = AppendRowsCore(source, sheetName, materialised);
        return ms.ToArray();
    }

    /// <summary>
    /// Reads a workbook from <paramref name="source"/>, appends <paramref name="rows"/> to
    /// <paramref name="sheetName"/>, and writes the result to <paramref name="destination"/>. See
    /// <see cref="AppendRows(byte[], string, IEnumerable{IEnumerable{object}})"/> for the
    /// semantics, including exactly what "last used row" means.
    ///
    /// <para><paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/>
    /// is <b>written</b>; neither is disposed, closed or sought.</para>
    /// </summary>
    /// <param name="source">The stream the workbook is read from.</param>
    /// <param name="sheetName">The sheet to append to.</param>
    /// <param name="rows">The rows to append.</param>
    /// <param name="destination">The stream the updated workbook is written to.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException"><paramref name="rows"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes; <paramref name="destination"/>
    /// is not writable; <paramref name="sheetName"/> is blank, is longer than 31 characters, or
    /// contains one of <c>: \ / ? * [ ]</c>; or an element of <paramref name="rows"/> is null.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The sheet was not found, or the package could not be opened or edited.</exception>
    public static async Task AppendRowsAsync(
        Stream source, string sheetName, IEnumerable<IEnumerable<object?>> rows, Stream destination,
        CancellationToken ct = default)
    {
        var materialised = ValidateRows(sheetName, rows);
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var xlsx = await StreamPipeline
            .DrainAsync(source, "Workbook content was empty.", nameof(source), "Failed to edit XLSX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        using var result = AppendRowsCore(xlsx, sheetName, materialised);
        await StreamPipeline.EmitAsync(result, destination, "Failed to edit XLSX. See the inner exception for details.", ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads a workbook from <paramref name="inputPath"/>, appends <paramref name="rows"/> to
    /// <paramref name="sheetName"/>, and writes the result to <paramref name="outputPath"/>,
    /// overwriting any existing file. See
    /// <see cref="AppendRows(byte[], string, IEnumerable{IEnumerable{object}})"/> for the
    /// semantics, including exactly what "last used row" means.
    /// </summary>
    /// <param name="inputPath">The workbook to read.</param>
    /// <param name="outputPath">Where to write the result. Overwritten if it exists.</param>
    /// <param name="sheetName">The sheet to append to.</param>
    /// <param name="rows">The rows to append.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    /// <exception cref="ArgumentNullException"><paramref name="rows"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="inputPath"/> or <paramref name="outputPath"/> is blank; <paramref name="sheetName"/>
    /// is blank, is longer than 31 characters, or contains one of <c>: \ / ? * [ ]</c>; or an
    /// element of <paramref name="rows"/> is null.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="inputPath"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException">
    /// <paramref name="inputPath"/>'s or <paramref name="outputPath"/>'s directory does not exist.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The sheet was not found, or the package could not be opened or edited.</exception>
    public static async Task AppendRowsAsync(
        string inputPath, string outputPath, string sheetName, IEnumerable<IEnumerable<object?>> rows,
        CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var bytes = await FilePipeline.ReadAsync(inputPath, nameof(inputPath), ct).ConfigureAwait(false);
        var result = AppendRows(bytes, sheetName, rows);
        await File.WriteAllBytesAsync(outputPath, result, ct).ConfigureAwait(false);
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
        // Checked BEFORE handing the bytes to ClosedXML, because ClosedXML reports a legacy .xls as
        // "File contains corrupted data" - which is false, and sends the caller to check a file,
        // a disk and an upload path that are all fine. Measured 2026-08-17 across 62 real .xls
        // files from a .gov crawl: every one reported as corrupt, every one a valid compound file
        // that Excel opens.
        //
        // This is the same defect this repository has recorded twice before in other places: a
        // message must not name a cause it cannot distinguish. Here it can distinguish, from the
        // first eight bytes, so it should.
        if (OfficeCrypto.IsEncrypted(xlsx))
        {
            throw new DocumentConversionException(
                "This is not an .xlsx package. The bytes are a compound file, which means either a "
                + "legacy Excel 97-2003 .xls workbook - save it as .xlsx to read it here - or an "
                + "encrypted .xlsx, which WorkbookEditor.Unprotect will open with its password.");
        }

        // Zero-copy and disposed, matching every other call site in this file. It used to copy the
        // whole workbook into a growable MemoryStream and never dispose it - on the 1.9 MB,
        // 40,000-row workbook this repository has measured at ~120 MB peak, a full extra copy plus
        // the doubling garbage a growable stream makes, for no benefit.
        //
        // Disposing here is safe for the reason PdfEditor.Open records for PdfSharp: ClosedXML
        // reads the whole package during construction, so the stream is not needed afterwards.
        // Asserted by the WorkbookEditor suite rather than assumed - every read, write and export
        // path goes through this method.
        using var source = new MemoryStream(xlsx, writable: false);
        return new XLWorkbook(source);
    }

    private static IXLWorksheet Sheet(XLWorkbook workbook, string sheetName)
    {
        if (!workbook.Worksheets.TryGetWorksheet(sheetName, out var sheet))
        {
            throw new DocumentConversionException(
                $"Worksheet '{sheetName}' was not found. Call WorkbookEditor.SheetNames to see "
                + "what is available.");
        }
        return sheet;
    }

    private static void SetCellValue(IXLCell cell, object? value)
    {
        switch (value)
        {
            case null: cell.Clear(XLClearOptions.Contents); break;

            // Placed in the ONE shared cell-writing path so Create, AppendRows and SetCell cannot
            // disagree about what an XlsxFormula means - the same reason every capability has a
            // single *Core method behind its overloads.
            case XlsxFormula f: cell.FormulaA1 = f.Formula; break;
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
    /// <paramref name="outputPath"/>. See <see cref="Create(string, IEnumerable{IEnumerable{object}})"/>
    /// for the exact typing and culture rules applied to each cell — this overload applies the
    /// identical logic, writing to <paramref name="outputPath"/> instead of returning an array.
    ///
    /// Named <c>CreateToFileAsync</c> rather than a third <c>CreateAsync</c> overload: <paramref
    /// name="sheetName"/> and <paramref name="rows"/> come first, same as
    /// <see cref="CreateAsync(string, IEnumerable{IEnumerable{object}}, Stream, CancellationToken)"/>,
    /// but the destination is a <c>string</c> path instead of a <c>Stream</c> — the distinct name
    /// keeps which kind of destination a call writes to visible at the call site, rather than
    /// resting on the argument type alone.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to create.</param>
    /// <param name="rows">The rows to populate it with.</param>
    /// <param name="outputPath">Where to write the workbook. Overwritten if it exists.</param>
    /// <param name="ct">Cancels the write to <paramref name="outputPath"/>.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="outputPath"/>, <paramref name="sheetName"/> or <paramref name="rows"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="outputPath"/> is blank; <paramref name="sheetName"/> is blank, is longer
    /// than 31 characters, or contains one of <c>: \ / ? * [ ]</c>; or a row is null.
    /// </exception>
    /// <exception cref="DirectoryNotFoundException"><paramref name="outputPath"/>'s directory does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be built.</exception>
    public static async Task CreateToFileAsync(
        string sheetName, IEnumerable<IEnumerable<object?>> rows, string outputPath,
        CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var bytes = Create(sheetName, rows);
        await File.WriteAllBytesAsync(outputPath, bytes, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Builds a workbook from <paramref name="sheets"/>, one worksheet each, in sequence order.
    /// Content comes from data rather than a template, so there is no source file to edit.
    ///
    /// <para>Cell typing and culture rules are identical to
    /// <see cref="Create(string, IEnumerable{IEnumerable{object}})"/>. A cell holding an
    /// <see cref="XlsxFormula"/> is written as a formula — see that type for the one limit worth
    /// knowing about cached values.</para>
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="sheets"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="sheets"/> is empty, contains a null element, or names the same sheet twice.
    /// </exception>
    /// <exception cref="DocumentConversionException">The workbook could not be built.</exception>
    /// <example>
    /// <code source="../../tests/DocToolkit.Tests/DocumentationExamples.cs" region="WorkbookCreate"/>
    /// </example>
    public static byte[] Create(IEnumerable<XlsxSheet> sheets)
    {
        var materialised = ValidateSheets(sheets);
        using var ms = CreateCore(materialised);
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a workbook from <paramref name="sheets"/> and writes it to
    /// <paramref name="destination"/>. See <see cref="Create(IEnumerable{XlsxSheet})"/> for the
    /// semantics — this overload applies identical logic.
    ///
    /// <para><paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought.</para>
    /// </summary>
    /// <param name="sheets">The sheets to build the workbook from, one worksheet each.</param>
    /// <param name="destination">The stream the workbook is written to.</param>
    /// <param name="ct">Cancels the build and the write to <paramref name="destination"/>.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="sheets"/> or <paramref name="destination"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="sheets"/> is empty, contains a null element, or names the same sheet
    /// twice; or <paramref name="destination"/> is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be built or written.</exception>
    public static async Task CreateAsync(
        IEnumerable<XlsxSheet> sheets, Stream destination, CancellationToken ct = default)
    {
        var materialised = ValidateSheets(sheets);
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var ms = CreateCore(materialised);
        await StreamPipeline.EmitAsync(ms, destination, "Failed to create XLSX. See the inner exception for details.", ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Builds a workbook from <paramref name="sheets"/> and writes it to
    /// <paramref name="outputPath"/>, overwriting any existing file. See
    /// <see cref="Create(IEnumerable{XlsxSheet})"/> for the semantics.
    /// </summary>
    /// <param name="sheets">The sheets to build the workbook from, one worksheet each.</param>
    /// <param name="outputPath">Where to write the workbook. Overwritten if it exists.</param>
    /// <param name="ct">Cancels the write to <paramref name="outputPath"/>.</param>
    /// <exception cref="ArgumentNullException"><paramref name="sheets"/> or <paramref name="outputPath"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="sheets"/> is empty, contains a null element, or names the same sheet
    /// twice; or <paramref name="outputPath"/> is blank.
    /// </exception>
    /// <exception cref="DirectoryNotFoundException"><paramref name="outputPath"/>'s directory does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be built or written.</exception>
    public static async Task CreateToFileAsync(
        IEnumerable<XlsxSheet> sheets, string outputPath, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var bytes = Create(sheets);
        await File.WriteAllBytesAsync(outputPath, bytes, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads a workbook from <paramref name="inputPath"/>, sets one cell, and writes the result to
    /// <paramref name="outputPath"/>. <paramref name="cellRef"/> is an A1-style reference. A cell
    /// holding an <see cref="XlsxFormula"/> is written as a formula instead of a literal value —
    /// see that type for the one limit worth knowing about cached values. The two paths may be
    /// the same file: the updated bytes are computed in full before
    /// <paramref name="outputPath"/> is opened, so a workbook that fails to process — cannot be
    /// read, or cannot be edited — leaves <paramref name="outputPath"/> untouched. That guarantee
    /// does not extend to a failure during the write itself: a full disk, a cancellation, or the
    /// process dying mid-write can still leave a partial file, so in-place editing of an
    /// irreplaceable workbook is not crash-safe.
    /// </summary>
    /// <param name="inputPath">The workbook to read.</param>
    /// <param name="outputPath">Where to write the result. Overwritten if it exists.</param>
    /// <param name="sheetName">The sheet containing the cell.</param>
    /// <param name="cellRef">An A1-style cell reference, e.g. <c>"B2"</c>.</param>
    /// <param name="value">
    /// The value to write. <c>null</c> clears the cell; an <see cref="XlsxFormula"/> writes a
    /// formula.
    /// </param>
    /// <param name="ct">Cancels the read and the write.</param>
    /// <exception cref="ArgumentNullException">A path or a name is null.</exception>
    /// <exception cref="ArgumentException">
    /// A path or a name is blank, or the file at <paramref name="inputPath"/> is empty.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="inputPath"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException">
    /// <paramref name="inputPath"/>'s or <paramref name="outputPath"/>'s directory does not exist.
    /// </exception>
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

        var bytes = await FilePipeline.ReadAsync(inputPath, nameof(inputPath), ct).ConfigureAwait(false);
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
    /// <exception cref="ArgumentNullException">A path or a name is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="path"/> or a name is blank, or the file it names is empty.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="path"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException"><paramref name="path"/>'s directory does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    public static async Task<string> ReadCellAsync(
        string path, string sheetName, string cellRef, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        var bytes = await FilePipeline.ReadAsync(path, nameof(path), ct).ConfigureAwait(false);
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
    /// <exception cref="ArgumentException">
    /// <paramref name="path"/> is blank, or the file it names is empty.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="path"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException"><paramref name="path"/>'s directory does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be opened.</exception>
    public static async Task<IReadOnlyList<string>> SheetNamesAsync(
        string path, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        var bytes = await FilePipeline.ReadAsync(path, nameof(path), ct).ConfigureAwait(false);
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
    /// <exception cref="ArgumentNullException">A path or <paramref name="sheetName"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="path"/> or <paramref name="sheetName"/> is blank, or the file at
    /// <paramref name="path"/> is empty.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="path"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException"><paramref name="path"/>'s directory does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the sheet's used range
    /// exceeds the 2,000,000-cell limit <see cref="ReadSheet"/> will materialise.
    /// </exception>
    public static async Task<IReadOnlyList<IReadOnlyList<string>>> ReadSheetAsync(
        string path, string sheetName, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        var bytes = await FilePipeline.ReadAsync(path, nameof(path), ct).ConfigureAwait(false);
        return ReadSheet(bytes, sheetName);
    }

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
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    public static byte[] AddChart(
        byte[] xlsx, string sheetName, string cellRef, ChartType type, ChartData data,
        string title = "", int widthPixels = 640, int heightPixels = 360)
    {
        ValidateArguments(xlsx, sheetName, cellRef);
        ArgumentNullException.ThrowIfNull(data);

        using var source = new MemoryStream(xlsx, writable: false);
        using var result = AddChartCore(source, sheetName, cellRef, type, data, title, widthPixels, heightPixels);
        return result.ToArray();
    }

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
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    public static async Task AddChartAsync(
        Stream source, string sheetName, string cellRef, ChartType type, ChartData data, Stream destination,
        string title = "", int widthPixels = 640, int heightPixels = 360, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(cellRef);
        ArgumentNullException.ThrowIfNull(data);
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var xlsx = await StreamPipeline
            .DrainAsync(source, "Workbook content was empty.", nameof(source), "Failed to add a chart to the XLSX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        using var result = AddChartCore(xlsx, sheetName, cellRef, type, data, title, widthPixels, heightPixels);
        await StreamPipeline.EmitAsync(result, destination, "Failed to add a chart to the XLSX. See the inner exception for details.", ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads a workbook from <paramref name="inputPath"/>, adds a chart, and writes the result to
    /// <paramref name="outputPath"/> — see <see cref="AddChart"/> for the parameters. The two
    /// paths may be the same file: the updated bytes are computed in full before
    /// <paramref name="outputPath"/> is opened.
    /// </summary>
    /// <param name="inputPath">The workbook to read.</param>
    /// <param name="outputPath">Where to write the result. Overwritten if it exists.</param>
    /// <param name="sheetName">The sheet to add the chart to.</param>
    /// <param name="cellRef">
    /// An A1-style cell reference for the chart's top-left corner, e.g. <c>"B2"</c>.
    /// </param>
    /// <param name="type">The chart's shape.</param>
    /// <param name="data">The chart's categories and value series.</param>
    /// <param name="title">The chart's title. Empty for no title.</param>
    /// <param name="widthPixels">The chart's width, in pixels.</param>
    /// <param name="heightPixels">The chart's height, in pixels.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    /// <exception cref="ArgumentNullException">A path, a name or <paramref name="data"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// A path or a name is blank, or the file at <paramref name="inputPath"/> is empty.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="inputPath"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException">
    /// <paramref name="inputPath"/>'s or <paramref name="outputPath"/>'s directory does not exist.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    public static async Task AddChartAsync(
        string inputPath, string outputPath, string sheetName, string cellRef, ChartType type, ChartData data,
        string title = "", int widthPixels = 640, int heightPixels = 360, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var bytes = await FilePipeline.ReadAsync(inputPath, nameof(inputPath), ct).ConfigureAwait(false);
        var result = AddChart(bytes, sheetName, cellRef, type, data, title, widthPixels, heightPixels);
        await File.WriteAllBytesAsync(outputPath, result, ct).ConfigureAwait(false);
    }

    private static MemoryStream AddChartCore(
        Stream xlsx, string sheetName, string cellRef, ChartType type, ChartData data,
        string title, int widthPixels, int heightPixels)
    {
        try
        {
            if (!XLHelper.IsValidA1Address(cellRef))
                throw new DocumentConversionException($"\"{cellRef}\" is not a valid A1-style cell reference.");
            // IsValidA1Address accepts absolute references ("$B$2"), but GetColumnNumberFromAddress
            // throws on the literal "$" - measured directly. SetCell/ReadCell accept "$"-prefixed
            // refs fine (ClosedXML's own Cell(cellRef) handles them), so AddChart strips "$" before
            // parsing to match that same convention rather than rejecting a reference those methods
            // already accept.
            var cleanedRef = cellRef.Replace("$", string.Empty);
            var column = XLHelper.GetColumnNumberFromAddress(cleanedRef);
            var row = int.Parse(
                new string(cleanedRef.SkipWhile(char.IsLetter).ToArray()), CultureInfo.InvariantCulture);

            // xlsx is typically a non-writable MemoryStream (new MemoryStream(bytes, writable:
            // false)) and OfficeIMO's ExcelDocument.Load needs an editable package, so this copy -
            // unlike the one CLAUDE.md records removing elsewhere - is load-bearing, not incidental.
            using var source = new MemoryStream();
            xlsx.CopyTo(source);
            source.Position = 0;
            using var document = OfficeIMOExcelExcelDocument.Load(source);

            var sheet = document.Sheets.FirstOrDefault(s => s.Name == sheetName)
                ?? throw new DocumentConversionException($"Sheet \"{sheetName}\" was not found.");

            sheet.AddChart(
                ToOfficeChartKind(type), ToOfficeChartData(data), row, column, widthPixels, heightPixels, title);

            var ms = new MemoryStream();
            document.Save(ms);
            return ms;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to add a chart to the XLSX. See the inner exception for details.", ex);
        }
    }

    private static OfficeIMO.Drawing.OfficeChartKind ToOfficeChartKind(ChartType type) => type switch
    {
        ChartType.ColumnClustered => OfficeIMO.Drawing.OfficeChartKind.ColumnClustered,
        ChartType.ColumnStacked => OfficeIMO.Drawing.OfficeChartKind.ColumnStacked,
        ChartType.ColumnStacked100 => OfficeIMO.Drawing.OfficeChartKind.ColumnStacked100,
        ChartType.BarClustered => OfficeIMO.Drawing.OfficeChartKind.BarClustered,
        ChartType.BarStacked => OfficeIMO.Drawing.OfficeChartKind.BarStacked,
        ChartType.BarStacked100 => OfficeIMO.Drawing.OfficeChartKind.BarStacked100,
        ChartType.Line => OfficeIMO.Drawing.OfficeChartKind.Line,
        ChartType.LineStacked => OfficeIMO.Drawing.OfficeChartKind.LineStacked,
        ChartType.LineStacked100 => OfficeIMO.Drawing.OfficeChartKind.LineStacked100,
        ChartType.Area => OfficeIMO.Drawing.OfficeChartKind.Area,
        ChartType.AreaStacked => OfficeIMO.Drawing.OfficeChartKind.AreaStacked,
        ChartType.AreaStacked100 => OfficeIMO.Drawing.OfficeChartKind.AreaStacked100,
        ChartType.Radar => OfficeIMO.Drawing.OfficeChartKind.Radar,
        ChartType.Pie => OfficeIMO.Drawing.OfficeChartKind.Pie,
        ChartType.Doughnut => OfficeIMO.Drawing.OfficeChartKind.Doughnut,
        _ => throw new ArgumentOutOfRangeException(nameof(type), type, null),
    };

    private static OfficeIMO.Drawing.OfficeChartData ToOfficeChartData(ChartData data) => new(
        data.Categories,
        data.Series.Select(s => new OfficeIMO.Drawing.OfficeChartSeries(s.Name, s.Values)));

    /// <summary>
    /// Adds a pivot table to <paramref name="sheetName"/> and returns the updated workbook.
    /// </summary>
    /// <remarks>
    /// <para><b>The result grid is empty until Excel opens and recalculates it.</b> A pivot
    /// table's aggregated values are computed by whichever application opens the file — nothing
    /// that WRITES it (this method included) populates the grid. This is a HARDER version of the
    /// limitation this package already documents for <see cref="XlsxFormula"/>: a formula's value
    /// <b>is</b> computed by <see cref="ReadCell"/>/<see cref="ReadSheet"/> on read, because this
    /// library's own engine evaluates it — but there is no equivalent pivot-evaluation engine
    /// here, so reading the pivot's own cells back with <see cref="ReadCell"/> immediately after
    /// calling this method returns empty strings, and <c>XlsxToPdfConverter</c> renders nothing
    /// where the pivot's results would be — for the identical reason it renders a formula's
    /// literal text rather than its computed value. Open the result in Excel (or an equivalent)
    /// to see it populated.
    /// </para>
    /// <para>Further edits to the workbook through this class's other methods (all
    /// ClosedXML-based) re-serialize the pivot table's XML — measured directly — but its field
    /// structure and aggregation choices survive that re-serialization correctly.</para>
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
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or
    /// <paramref name="destinationCell"/> is not a valid cell reference.
    /// </exception>
    public static byte[] AddPivotTable(
        byte[] xlsx, string sheetName, string sourceRange, string destinationCell, string name,
        IEnumerable<string> rowFields, IEnumerable<PivotDataField> dataFields,
        IEnumerable<string>? columnFields = null, IEnumerable<string>? pageFields = null,
        bool showRowGrandTotals = true, bool showColumnGrandTotals = true)
    {
        // ValidateArguments already covers xlsx (null/empty), sheetName and destinationCell
        // (blank) in one place shared with AddChart, so the two methods cannot disagree about
        // which argument a mistake gets blamed on.
        ValidateArguments(xlsx, sheetName, destinationCell);
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceRange);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        var rowFieldList = ValidatePivotFields(rowFields, nameof(rowFields));
        var dataFieldList = (dataFields ?? throw new ArgumentNullException(nameof(dataFields))).ToList();
        if (dataFieldList.Count == 0)
            throw new ArgumentException("Data fields were empty.", nameof(dataFields));

        using var source = new MemoryStream(xlsx, writable: false);
        using var result = AddPivotTableCore(
            source, sheetName, sourceRange, destinationCell, name, rowFieldList, dataFieldList,
            columnFields?.ToList(), pageFields?.ToList(), showRowGrandTotals, showColumnGrandTotals);
        return result.ToArray();
    }

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
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or
    /// <paramref name="destinationCell"/> is not a valid cell reference.
    /// </exception>
    public static async Task AddPivotTableAsync(
        Stream source, string sheetName, string sourceRange, string destinationCell, string name,
        IEnumerable<string> rowFields, IEnumerable<PivotDataField> dataFields, Stream destination,
        IEnumerable<string>? columnFields = null, IEnumerable<string>? pageFields = null,
        bool showRowGrandTotals = true, bool showColumnGrandTotals = true, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceRange);
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationCell);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        var rowFieldList = ValidatePivotFields(rowFields, nameof(rowFields));
        var dataFieldList = (dataFields ?? throw new ArgumentNullException(nameof(dataFields))).ToList();
        if (dataFieldList.Count == 0)
            throw new ArgumentException("Data fields were empty.", nameof(dataFields));
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var xlsx = await StreamPipeline
            .DrainAsync(source, "Workbook content was empty.", nameof(source), "Failed to add a pivot table to the XLSX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        using var result = AddPivotTableCore(
            xlsx, sheetName, sourceRange, destinationCell, name, rowFieldList, dataFieldList,
            columnFields?.ToList(), pageFields?.ToList(), showRowGrandTotals, showColumnGrandTotals);
        await StreamPipeline.EmitAsync(result, destination, "Failed to add a pivot table to the XLSX. See the inner exception for details.", ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads a workbook from <paramref name="inputPath"/>, adds a pivot table, and writes the
    /// result to <paramref name="outputPath"/> — see <see cref="AddPivotTable"/> for the
    /// parameters. The two paths may be the same file: the updated bytes are computed in full
    /// before <paramref name="outputPath"/> is opened.
    /// </summary>
    /// <param name="inputPath">The workbook to read.</param>
    /// <param name="outputPath">Where to write the result. Overwritten if it exists.</param>
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
    /// <param name="ct">Cancels the read and the write.</param>
    /// <exception cref="ArgumentNullException">
    /// A path, a name, <paramref name="rowFields"/> or <paramref name="dataFields"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// A path or a name is blank, the file at <paramref name="inputPath"/> is empty, or
    /// <paramref name="rowFields"/>/<paramref name="dataFields"/> is empty.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="inputPath"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException">
    /// <paramref name="inputPath"/>'s or <paramref name="outputPath"/>'s directory does not exist.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or
    /// <paramref name="destinationCell"/> is not a valid cell reference.
    /// </exception>
    public static async Task AddPivotTableAsync(
        string inputPath, string outputPath, string sheetName, string sourceRange,
        string destinationCell, string name, IEnumerable<string> rowFields,
        IEnumerable<PivotDataField> dataFields, IEnumerable<string>? columnFields = null,
        IEnumerable<string>? pageFields = null, bool showRowGrandTotals = true,
        bool showColumnGrandTotals = true, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var bytes = await FilePipeline.ReadAsync(inputPath, nameof(inputPath), ct).ConfigureAwait(false);
        var result = AddPivotTable(
            bytes, sheetName, sourceRange, destinationCell, name, rowFields, dataFields,
            columnFields, pageFields, showRowGrandTotals, showColumnGrandTotals);
        await File.WriteAllBytesAsync(outputPath, result, ct).ConfigureAwait(false);
    }

    private static List<string> ValidatePivotFields(IEnumerable<string> fields, string paramName)
    {
        var list = (fields ?? throw new ArgumentNullException(paramName)).ToList();
        if (list.Count == 0) throw new ArgumentException($"{paramName} were empty.", paramName);
        return list;
    }

    private static MemoryStream AddPivotTableCore(
        Stream xlsx, string sheetName, string sourceRange, string destinationCell, string name,
        IReadOnlyList<string> rowFields, IReadOnlyList<PivotDataField> dataFields,
        IReadOnlyList<string>? columnFields, IReadOnlyList<string>? pageFields,
        bool showRowGrandTotals, bool showColumnGrandTotals)
    {
        try
        {
            // Cleans then validates - the reverse order from AddChartCore's cellRef handling,
            // which validates the original string first. Both accept every reference the other
            // does; the one difference is a malformed-but-strippable input like "D1$", which this
            // order accepts and AddChartCore's rejects. Not unified deliberately: AddChartCore
            // decomposes cellRef into row/column ints afterward, while this passes the string
            // straight to OfficeIMO, so the two validate different things even where they agree.
            var cleanedDestinationCell = destinationCell.Replace("$", string.Empty);
            if (!XLHelper.IsValidA1Address(cleanedDestinationCell))
                throw new DocumentConversionException($"\"{destinationCell}\" is not a valid A1-style cell reference.");

            // xlsx is typically a non-writable MemoryStream and OfficeIMO's ExcelDocument.Load
            // needs an editable package, so this copy is load-bearing - see AddChartCore's own
            // identical comment for the full reasoning.
            using var source = new MemoryStream();
            xlsx.CopyTo(source);
            source.Position = 0;
            using var document = OfficeIMOExcelExcelDocument.Load(source);

            var sheet = document.Sheets.FirstOrDefault(s => s.Name == sheetName)
                ?? throw new DocumentConversionException($"Sheet \"{sheetName}\" was not found.");

            sheet.AddPivotTable(
                sourceRange: sourceRange,
                destinationCell: cleanedDestinationCell,
                name: name,
                rowFields: rowFields,
                columnFields: columnFields,
                pageFields: pageFields,
                dataFields: dataFields.Select(f => new OfficeIMO.Excel.ExcelPivotDataField(f.FieldName, ToPivotDataFunction(f.Function))),
                showRowGrandTotals: showRowGrandTotals,
                showColumnGrandTotals: showColumnGrandTotals);

            var ms = new MemoryStream();
            document.Save(ms);
            return ms;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to add a pivot table to the XLSX. See the inner exception for details.", ex);
        }
    }

    private static OfficeIMO.Excel.ExcelPivotDataFunction ToPivotDataFunction(PivotFunction function) => function switch
    {
        PivotFunction.Sum => OfficeIMO.Excel.ExcelPivotDataFunction.Sum,
        PivotFunction.Average => OfficeIMO.Excel.ExcelPivotDataFunction.Average,
        PivotFunction.Count => OfficeIMO.Excel.ExcelPivotDataFunction.Count,
        PivotFunction.CountNumbers => OfficeIMO.Excel.ExcelPivotDataFunction.CountNumbers,
        PivotFunction.Maximum => OfficeIMO.Excel.ExcelPivotDataFunction.Maximum,
        PivotFunction.Minimum => OfficeIMO.Excel.ExcelPivotDataFunction.Minimum,
        PivotFunction.Product => OfficeIMO.Excel.ExcelPivotDataFunction.Product,
        PivotFunction.StandardDeviation => OfficeIMO.Excel.ExcelPivotDataFunction.StandardDeviation,
        PivotFunction.StandardDeviationP => OfficeIMO.Excel.ExcelPivotDataFunction.StandardDeviationP,
        PivotFunction.Variance => OfficeIMO.Excel.ExcelPivotDataFunction.Variance,
        PivotFunction.VarianceP => OfficeIMO.Excel.ExcelPivotDataFunction.VarianceP,
        _ => throw new ArgumentOutOfRangeException(nameof(function), function, null),
    };

    /// <summary>
    /// A copy of <paramref name="xlsx"/> encrypted with <paramref name="password"/>, so it cannot
    /// be opened without one.
    /// </summary>
    /// <remarks>
    /// <b>This is file encryption, not workbook protection.</b> Office offers both under the same
    /// menu and they are not the same thing: this scrambles the whole file, so nothing can be read
    /// without the password. The other kind - a flag asking a reader not to edit - is a request
    /// rather than a lock, and is deliberately not exposed here.
    ///
    /// <b>The result is not a XLSX package any more.</b> An encrypted Office document is a
    /// compound file with the package sealed inside it, so every other method on this class refuses
    /// it - call <see cref="Unprotect(byte[], string)"/> first. That refusal is the honest
    /// behaviour: those methods could not read the content even if they tried.
    /// </remarks>
    /// <param name="xlsx">The workbook to encrypt.</param>
    /// <param name="password">The password required to open the result. May not be empty.</param>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be read or encrypted.</exception>
    public static byte[] Protect(byte[] xlsx, string password)
    {
        ArgumentNullException.ThrowIfNull(xlsx);
        if (xlsx.Length == 0)
            throw new ArgumentException("XLSX content was empty.", nameof(xlsx));
        OfficeCrypto.RequirePassword(password, nameof(password));

        return OfficeCrypto.TranslateWrite(() =>
        {
            using var source = new MemoryStream(xlsx, writable: false);
            using var document = OfficeIMOExcelExcelDocument.Load(source);
            using var encrypted = new MemoryStream();
            document.SaveEncrypted(encrypted, password);
            return encrypted.ToArray();
        }, "XLSX");
    }

    /// <summary>
    /// A copy of <paramref name="xlsx"/> with its encryption removed, so the rest of this class
    /// can work on it.
    /// </summary>
    /// <remarks>
    /// <b>The output is not protected in any way.</b> That is what was asked for, but the bytes
    /// this returns are readable by anyone who obtains them.
    ///
    /// A workbook that was never encrypted is reported as such rather than passed through, because
    /// silently returning the input would make a broken pipeline look like a working one.
    /// </remarks>
    /// <param name="xlsx">The encrypted workbook.</param>
    /// <param name="password">The password the workbook was encrypted with.</param>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">
    /// The password was wrong, the workbook was not encrypted, or it could not be read.
    /// </exception>
    public static byte[] Unprotect(byte[] xlsx, string password)
    {
        ArgumentNullException.ThrowIfNull(xlsx);
        if (xlsx.Length == 0)
            throw new ArgumentException("XLSX content was empty.", nameof(xlsx));
        OfficeCrypto.RequirePassword(password, nameof(password));

        return OfficeCrypto.Translate(() =>
        {
            using var source = new MemoryStream(xlsx, writable: false);
            using var document = OfficeIMOExcelExcelDocument.LoadEncrypted(source, password);
            using var plain = new MemoryStream();
            document.Save(plain);
            return plain.ToArray();
        }, "XLSX");
    }

    /// <summary>
    /// Whether <paramref name="xlsx"/> is an ENCRYPTED Office document.
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
    /// Reads the file signature; it does not try the password and does not need one. A plain XLSX
    /// is a ZIP package, an encrypted one is a compound file, and the two are distinguishable from
    /// their first eight bytes.
    /// </remarks>
    /// <param name="xlsx">The bytes to inspect.</param>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    public static bool IsProtected(byte[] xlsx)
    {
        ArgumentNullException.ThrowIfNull(xlsx);

        return OfficeCrypto.IsEncrypted(xlsx);
    }

    /// <summary>
    /// Reads a workbook from <paramref name="source"/> and writes the encrypted copy to
    /// <paramref name="destination"/>.
    ///
    /// Neither stream is disposed, closed or sought.
    /// </summary>
    /// <inheritdoc cref="Protect(byte[], string)" path="/remarks|/exception"/>
    /// <param name="source">The stream the workbook is read from.</param>
    /// <param name="destination">The stream the encrypted workbook is written to.</param>
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
            .DrainAsync(source, "XLSX content was empty.", nameof(source),
                        "Failed to encrypt the XLSX.", ct)
            .ConfigureAwait(false);

        using var result = new MemoryStream(Protect(buffer.ToArray(), password), writable: false);
        await StreamPipeline
            .EmitAsync(result, destination, "Failed to encrypt the XLSX.", ct)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Reads an encrypted workbook from <paramref name="source"/> and writes the unprotected copy to
    /// <paramref name="destination"/>.
    ///
    /// Neither stream is disposed, closed or sought.
    /// </summary>
    /// <inheritdoc cref="Unprotect(byte[], string)" path="/remarks|/exception"/>
    /// <param name="source">The stream the encrypted workbook is read from.</param>
    /// <param name="destination">The stream the unprotected workbook is written to.</param>
    /// <param name="password">The password the workbook was encrypted with.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    public static async Task UnprotectAsync(
        Stream source, Stream destination, string password, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        OfficeCrypto.RequirePassword(password, nameof(password));
        ct.ThrowIfCancellationRequested();

        using var buffer = await StreamPipeline
            .DrainAsync(source, "XLSX content was empty.", nameof(source),
                        "Failed to read the encrypted XLSX.", ct)
            .ConfigureAwait(false);

        using var result = new MemoryStream(Unprotect(buffer.ToArray(), password), writable: false);
        await StreamPipeline
            .EmitAsync(result, destination, "Failed to read the encrypted XLSX.", ct)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Inspects <paramref name="xlsx"/> for digital signatures — whether it carries one, how
    /// many, and who claims to have signed it. Does not validate anything cryptographically; see
    /// <see cref="ValidateSignatures"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be inspected.</exception>
    public static DocumentSignatureInfo InspectSignatures(byte[] xlsx)
    {
        ArgumentNullException.ThrowIfNull(xlsx);
        if (xlsx.Length == 0) throw new ArgumentException("XLSX content was empty.", nameof(xlsx));

        return OfficeSignature.Inspect(xlsx, ".xlsx", OfficeIMOExcelExcelDocument.InspectPackageSignatures, "XLSX");
    }

    /// <summary>
    /// Reads an .xlsx from <paramref name="source"/> and inspects it for digital signatures — see
    /// <see cref="InspectSignatures"/>. <paramref name="source"/> is read to its end and is
    /// neither disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be inspected.</exception>
    public static async Task<DocumentSignatureInfo> InspectSignaturesAsync(Stream source, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var xlsx = await StreamPipeline
            .DrainAsync(source, "Workbook content was empty.", nameof(source), "Failed to inspect XLSX signatures. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        return OfficeSignature.Inspect(xlsx.ToArray(), ".xlsx", OfficeIMOExcelExcelDocument.InspectPackageSignatures, "XLSX");
    }

    /// <summary>
    /// Validates every digital signature <paramref name="xlsx"/> carries — cryptographic
    /// integrity (tamper detection), certificate chain trust, and revocation, each reported
    /// independently. Never performs revocation checking or certificate downloads over the
    /// network, regardless of <paramref name="options"/> — see
    /// <see cref="DocumentSignatureValidationOptions"/>'s own remarks.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be validated.</exception>
    public static DocumentSignatureValidationReport ValidateSignatures(byte[] xlsx, DocumentSignatureValidationOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(xlsx);
        if (xlsx.Length == 0) throw new ArgumentException("XLSX content was empty.", nameof(xlsx));

        return OfficeSignature.Validate(xlsx, ".xlsx", OfficeIMOExcelExcelDocument.ValidatePackageSignatures, options, "XLSX");
    }

    /// <summary>
    /// Reads an .xlsx from <paramref name="source"/> and validates its digital signatures — see
    /// <see cref="ValidateSignatures"/>. <paramref name="source"/> is read to its end and is
    /// neither disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be validated.</exception>
    public static async Task<DocumentSignatureValidationReport> ValidateSignaturesAsync(
        Stream source, DocumentSignatureValidationOptions? options = null, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var xlsx = await StreamPipeline
            .DrainAsync(source, "Workbook content was empty.", nameof(source), "Failed to validate XLSX signatures. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        return OfficeSignature.Validate(xlsx.ToArray(), ".xlsx", OfficeIMOExcelExcelDocument.ValidatePackageSignatures, options, "XLSX");
    }
}
