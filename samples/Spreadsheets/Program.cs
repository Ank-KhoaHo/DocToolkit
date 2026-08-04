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

Console.WriteLine("\nDone.");
