using ClosedXML.Excel;

namespace DocToolkit;

/// <summary>Creates, reads and edits Excel (.xlsx) workbooks. Legacy .xls is not supported.</summary>
public static class WorkbookEditor
{
    /// <summary>Creates a workbook with one sheet populated from <paramref name="rows"/>.</summary>
    public static byte[] Create(string sheetName, IEnumerable<IEnumerable<object?>> rows)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentNullException.ThrowIfNull(rows);

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

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        return ms.ToArray();
    }

    /// <summary>Reads a cell as a string. <paramref name="cellRef"/> is an A1-style reference.</summary>
    public static string ReadCell(byte[] xlsx, string sheetName, string cellRef)
    {
        ArgumentNullException.ThrowIfNull(xlsx);
        using var workbook = Open(xlsx);
        return Sheet(workbook, sheetName).Cell(cellRef).GetString();
    }

    /// <summary>Sets a cell and returns the updated workbook bytes.</summary>
    public static byte[] SetCell(byte[] xlsx, string sheetName, string cellRef, object? value)
    {
        ArgumentNullException.ThrowIfNull(xlsx);
        using var workbook = Open(xlsx);
        SetCellValue(Sheet(workbook, sheetName).Cell(cellRef), value);

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        return ms.ToArray();
    }

    private static XLWorkbook Open(byte[] xlsx)
    {
        if (xlsx.Length == 0)
            throw new ArgumentException("Workbook content was empty.", nameof(xlsx));
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
            case int or long or short or byte or double or float or decimal:
                cell.Value = Convert.ToDouble(value); break;
            default: cell.Value = value.ToString(); break;
        }
    }
}
