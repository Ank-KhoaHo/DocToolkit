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

// --- Trusting the computed value outside this package --------------------------------------
// ReadCell above already returns the right number because it recomputes in memory on every
// read - which does nothing for a reader that is NOT this package and NOT Excel, and only
// trusts whatever value the FILE already carries. InspectFormulas reports what the underlying
// engine actually understands; EvaluateFormulas writes a computed value into the file itself,
// for exactly that reader.

#region formula-evaluate
XlsxFormulaInspection inspection = WorkbookEditor.InspectFormulas(workbook);
byte[] withCachedValues = WorkbookEditor.EvaluateFormulas(workbook);
#endregion

Console.WriteLine($"\nFormulas found : {inspection.TotalFormulas} total, {inspection.SupportedFormulas} supported");
Console.WriteLine($"All supported  : {inspection.AllSupported}");
Console.WriteLine($"With cached values: {withCachedValues.Length:N0} bytes (was {workbook.Length:N0})");

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

// --- A real Excel table, rather than a range that looks like one ---------------------------
// A table (a ListObject) is what Excel's own "Format as Table" produces: a named object with
// banded rows and its own filter, which structured references like Revenue[Revenue] can point at.

#region table
// Its own Format call, not another line on the one above: a table brings its OWN autofilter, so
// WithTable and WithAutoFilter over overlapping ranges throw rather than silently picking one.
byte[] tabled = WorkbookEditor.Format(workbook, "Q1", XlsxFormat.None
    .WithTable(XlsxTable.Named("A1:B4", "Revenue", XlsxTableStyle.Medium)));
#endregion

// AppendRows has no awareness of a table on the sheet, and ClosedXML does not absorb an adjacent
// write into a table's range - measured. Recreate the table if you need it to cover new rows.
Console.WriteLine($"\nWith a table : {tabled.Length:N0} bytes (was {workbook.Length:N0})");

// --- The furniture that makes a sheet printable and annotated ------------------------------
// Page setup here is a WORKSHEET concern and a separate type from the DOCX PageSetup: the two
// formats do not agree on what a page is, so they do not share a type.

byte[] titled = WorkbookEditor.SetCell(tabled, "Q1", "D1", "Quarterly revenue");

#region worksheet-furniture
byte[] furnished = WorkbookEditor.Format(titled, "Q1", XlsxFormat.None
    // A banner cell across two columns.
    .WithMergedCells("D1:E1")

    // The URL must be absolute - a relative one is refused rather than written and silently
    // broken for whoever opens the file.
    .WithHyperlink(XlsxHyperlink.To("D2", "https://example.com/methodology"))

    // A note attached to the cell, not a value in it.
    .WithComment(XlsxComment.On("B2", "Restated after the FX correction."))

    // Landscape, a print area, and the header row repeated at the top of every printed page.
    .WithPageSetup(XlsxPageSetup.Of(XlsxPageOrientation.Landscape, printArea: "A1:E4", repeatRowCount: 1)));
#endregion

Console.WriteLine($"With furniture: {furnished.Length:N0} bytes");

// --- A named range and an embedded image ----------------------------------------------------
// Both are workbook EDITS rather than presentation, so they are their own WorkbookEditor methods
// rather than XlsxFormat members.

// Inline base64, so this sample carries no binary asset and borrows no test fixture. A real
// caller passes File.ReadAllBytes("logo.png"). PNG and JPEG only, decided by magic bytes rather
// than by a filename - a file named .png holding JPEG bytes is read as the JPEG it is.
const string LogoBase64 =
    "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAAAUElEQVR42u3PQQkAAAgEsOtlFCsZ2gi+hcEKLNXzWgQEBA" +
    "QEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQErsACwGghD5ay/wAAAAAASUVORK5CYII=";

#region defined-name-and-image
// The sheet name is always quoted in the reference this writes. Measured: an unquoted sheet name
// containing a space does not error - the defined name is simply gone when the file reopens.
byte[] named = WorkbookEditor.AddDefinedName(furnished, "RevenueCells", "Q1", "B2:B4");

// Sizes are PIXELS here, matching AddChart - not the points the DOCX/PPTX drawing model uses.
// Give neither and the image's own intrinsic size is used; give one and the other scales.
byte[] withLogo = WorkbookEditor.AddImage(
    named, "Q1", "G1", Convert.FromBase64String(LogoBase64), widthPixels: 64, heightPixels: 64);
#endregion

Console.WriteLine($"Named + image : {withLogo.Length:N0} bytes");

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

// --- Metadata ---------------------------------------------------------------------------------
// What a file manager shows in its properties panel, and what a search indexer reads - the same
// concept as PdfMetadata (see the PdfUtilities sample), on XLSX's own property bag.

#region metadata
byte[] withMetadata = WorkbookEditor.WithMetadata(workbook, new DocumentMetadata
{
    Title = "Q1 Regional Revenue",
    Creator = "Contoso Finance",
});

DocumentMetadata readBack = WorkbookEditor.ReadMetadata(withMetadata);
#endregion

Console.WriteLine($"\nTitle          : {readBack.Title}");
Console.WriteLine($"Creator        : {readBack.Creator}");
Console.WriteLine($"Subject        : {readBack.Subject?.ToString() ?? "(null - never set)"}");

// null means ABSENT, not blank - and a null property leaves what the document already had
// alone, so retitling below cannot silently erase the creator.
byte[] retitled = WorkbookEditor.WithMetadata(withMetadata, new DocumentMetadata { Title = "Superseded" });

Console.WriteLine($"After retitling: title \"{WorkbookEditor.ReadMetadata(retitled).Title}\", "
    + $"creator still \"{WorkbookEditor.ReadMetadata(retitled).Creator}\"");

Console.WriteLine("\nDone.");
