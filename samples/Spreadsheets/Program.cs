using DocToolkit;

Console.WriteLine("Spreadsheets");
Console.WriteLine("============");

// --- Create, read one cell, edit one cell -------------------------------------------------

#region create
byte[] xlsx = WorkbookEditor.Create("Sales", new object?[][]
{
    new object?[] { "Region", "Total" },
    new object?[] { "North",  1200 },
    new object?[] { "South",  950 },
});

string before = WorkbookEditor.ReadCell(xlsx, "Sales", "B2");
byte[] updated = WorkbookEditor.SetCell(xlsx, "Sales", "B2", 1500);
string after = WorkbookEditor.ReadCell(updated, "Sales", "B2");
#endregion

Console.WriteLine($"\nB2 before {before}, after SetCell {after}");

// --- Reading a workbook you were handed ---------------------------------------------------
// The point of these two: you do not need to know the workbook's shape in advance.

#region read
IReadOnlyList<string> sheets = WorkbookEditor.SheetNames(updated);
IReadOnlyList<IReadOnlyList<string>> grid = WorkbookEditor.ReadSheet(updated, sheets[0]);
#endregion

Console.WriteLine($"Sheets       : {string.Join(", ", sheets)}");
Console.WriteLine($"Shape        : {grid.Count} rows x {grid[0].Count} columns");

foreach (var row in grid)
    Console.WriteLine($"  | {string.Join(" | ", row)}");

// --- A workbook with more than one sheet, and a formula across them -----------------------
// Content comes from data rather than a template, so there is no source file to edit. A cell
// holding an XlsxFormula is written as a formula rather than as text.

#region multi-sheet
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
#endregion

Console.WriteLine($"\nSheets       : {string.Join(", ", WorkbookEditor.SheetNames(workbook))}");
Console.WriteLine($"Appended row : {string.Join(" | ", WorkbookEditor.ReadSheet(workbook, "Q1")[3])}");

// Read back through this package: ClosedXML evaluates the formula, because the file carries no
// cached value. A reader that only reads cached values would see an empty cell here until Excel
// has opened and saved the file.
Console.WriteLine($"Grand total  : {WorkbookEditor.ReadCell(workbook, "Summary", "B1")}");

// --- Making a generated sheet look like a report ------------------------------------------
// Format applies presentation to a sheet that already exists, so it composes with Create,
// AppendRows and SetCell rather than being an argument to any of them. XlsxFormat is immutable
// and every With* returns a new one, same as PageSetup.
//
// The boundary is a CLOSED vocabulary rather than a small one: six rule conditions, five
// validation kinds and four highlights, each enumerable and measured. Anything that cannot be
// expressed as a closed set - arbitrary fonts, borders, colour scales - is ClosedXML's job.

#region format
byte[] presented = WorkbookEditor.Format(workbook, "Q1", XlsxFormat.Report
    .WithNumberFormat("B", "#,##0.00")

    // Auto-fit sizes a column to what is in it today; an explicit width survives longer values.
    .WithColumnWidth("A", 14)

    // Report already freezes the header row. Naming a position freezes a column too, so the
    // region labels stay visible when a wide sheet scrolls sideways.
    .WithFreezeAt(row: 1, column: 1)
    .WithAutoFilter()

    // XlsxHighlight names an INTENT, never a colour - a colour picker cannot be enumerated,
    // and the moment one exists the closed vocabulary is gone.
    .WithRule(XlsxRule.GreaterThan("B2:B4", 1000, XlsxHighlight.Green))

    // The half of a generated workbook that survives a human editing it: Excel refuses a region
    // outside this list rather than accepting a typo that breaks tomorrow's formula.
    .WithValidation(XlsxValidation.OneOf("A2:A4", "EMEA", "APAC", "AMER")));
#endregion

Console.WriteLine($"\nFormatted    : {presented.Length:N0} bytes (was {workbook.Length:N0})");
Console.WriteLine($"B2 reads back: {WorkbookEditor.ReadCell(presented, "Q1", "B2")}");

// --- Handing a sheet to something that is not Excel ---------------------------------------
// One sheet at a time, by name - a workbook is not one table and neither format has a way to
// say "and now a different sheet".

#region export
string csv = XlsxToCsvConverter.Convert(presented, "Q1");
string html = XlsxToHtmlConverter.Convert(presented, "Q1");
#endregion

Console.WriteLine($"\nAs CSV       : {csv.Replace("\r\n", " / ").Replace("\n", " / ")}");
Console.WriteLine($"As HTML      : {html.Length:N0} chars, starts \"{html[..Math.Min(40, html.Length)]}\"");

// --- Pivoting existing sheet data -----------------------------------------------------------
// AddPivotTable reads a source range and writes a pivot definition elsewhere in the same sheet.
// Nothing that WRITES a workbook computes the aggregation - Excel does that the first time it
// opens the file - so the pivot's own cells read back empty until then.

#region pivot
byte[] sales = WorkbookEditor.Create("Sales", new object?[][]
{
    new object?[] { "Region", "Amount" },
    new object?[] { "North",  1200 },
    new object?[] { "South",  950 },
    new object?[] { "North",  300 },
    new object?[] { "South",  600 },
});

byte[] withPivot = WorkbookEditor.AddPivotTable(
    sales, "Sales", "A1:B5", "D1", "RegionSummary",
    rowFields: new[] { "Region" },
    dataFields: new[] { new PivotDataField("Amount", PivotFunction.Sum) });
#endregion

Console.WriteLine($"\nWith pivot   : {withPivot.Length:N0} bytes (was {sales.Length:N0})");
Console.WriteLine($"Pivot cell D1 right after creation: \"{WorkbookEditor.ReadCell(withPivot, "Sales", "D1")}\"");
Console.WriteLine("^ empty on purpose - open the file in Excel to see it recalculate.");

// --- A chart from the same data --------------------------------------------------------------
// WorkbookEditor.AddChart and PresentationEditor.AddChart (see the Presentations sample) share
// one ChartType/ChartData model, so the same data shape works for both formats.

#region chart
var chartData = new ChartData(
    new[] { "North", "South" },
    new[] { new ChartSeries("Total", new double[] { 1500, 1550 }) });

byte[] withChart = WorkbookEditor.AddChart(
    sales, "Sales", "D8", ChartType.ColumnClustered, chartData, title: "Regional Totals");
#endregion

// The exact byte count varies between runs and between platforms - AddChart writes chart XML
// parts through DocumentFormat.OpenXml directly, which assigns each saved part a fresh random
// relationship id (the same non-determinism the Presentations sample's own chart line has).
Console.WriteLine($"\nWith chart   : {withChart.Length:N0} bytes (does not touch the sheet's cell data)");
Console.WriteLine("^ unlike the pivot above, this DOES reach XlsxToPdfConverter's output - see the guide.");

Console.WriteLine("\nDone.");
