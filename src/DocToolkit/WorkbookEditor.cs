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
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentNullException.ThrowIfNull(rows);

        // Validated up front so a null row surfaces as the ArgumentException it is rather than as
        // a NullReferenceException wrapped in a conversion failure.
        var materialised = rows
            .Select((row, index) => row
                ?? throw new ArgumentException($"Row {index + 1} was null.", nameof(rows)))
            .ToList();

        try
        {
            using var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add(sheetName);

            var r = 1;
            foreach (var row in materialised)
            {
                var c = 1;
                foreach (var value in row)
                    SetCellValue(sheet.Cell(r, c++), value);
                r++;
            }

            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            return ms.ToArray();
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
