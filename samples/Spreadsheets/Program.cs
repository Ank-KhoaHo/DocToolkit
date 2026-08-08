using DocToolkit;

Console.WriteLine("Spreadsheets");
Console.WriteLine("============");

// --- Create, read one cell, edit one cell -------------------------------------------------

byte[] xlsx = WorkbookEditor.Create("Sales", new object?[][]
{
    new object?[] { "Region", "Total" },
    new object?[] { "North",  1200 },
    new object?[] { "South",  950 },
});

string before = WorkbookEditor.ReadCell(xlsx, "Sales", "B2");
byte[] updated = WorkbookEditor.SetCell(xlsx, "Sales", "B2", 1500);
string after = WorkbookEditor.ReadCell(updated, "Sales", "B2");

Console.WriteLine($"\nB2 before {before}, after SetCell {after}");

// --- Reading a workbook you were handed ---------------------------------------------------
// The point of these two: you do not need to know the workbook's shape in advance.

IReadOnlyList<string> sheets = WorkbookEditor.SheetNames(updated);
IReadOnlyList<IReadOnlyList<string>> grid = WorkbookEditor.ReadSheet(updated, sheets[0]);

Console.WriteLine($"Sheets       : {string.Join(", ", sheets)}");
Console.WriteLine($"Shape        : {grid.Count} rows x {grid[0].Count} columns");

foreach (var row in grid)
    Console.WriteLine($"  | {string.Join(" | ", row)}");

// --- A workbook with more than one sheet, and a formula across them -----------------------
// Content comes from data rather than a template, so there is no source file to edit. A cell
// holding an XlsxFormula is written as a formula rather than as text.

byte[] workbook = WorkbookEditor.Create(new[]
{
    XlsxSheet.Named("Q1", new[]
    {
        new object?[] { "Region", "Revenue" },
        new object?[] { "EMEA", 1200 },
        new object?[] { "APAC", 980 },
    }),
    XlsxSheet.Named("Summary", new[]
    {
        new object?[] { "Grand total", XlsxFormula.From("SUM(Q1!B2:B4)") },
    }),
});

// Rows append after the sheet's last used row, leaving every other sheet untouched.
workbook = WorkbookEditor.AppendRows(workbook, "Q1", new[]
{
    new object?[] { "AMER", 1450 },
});

Console.WriteLine($"\nSheets       : {string.Join(", ", WorkbookEditor.SheetNames(workbook))}");
Console.WriteLine($"Appended row : {string.Join(" | ", WorkbookEditor.ReadSheet(workbook, "Q1")[3])}");

// Read back through this package: ClosedXML evaluates the formula, because the file carries no
// cached value. A reader that only reads cached values would see an empty cell here until Excel
// has opened and saved the file.
Console.WriteLine($"Grand total  : {WorkbookEditor.ReadCell(workbook, "Summary", "B1")}");

Console.WriteLine("\nDone.");
