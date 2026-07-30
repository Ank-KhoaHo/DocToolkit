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

    /// <summary>Reads a cell as a string. <paramref name="cellRef"/> is an A1-style reference.</summary>
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
}
