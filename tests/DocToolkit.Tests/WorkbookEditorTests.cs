using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using DocToolkit;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace DocToolkit.Tests;

public class WorkbookEditorTests
{
    private static byte[] SampleWorkbook() => WorkbookEditor.Create("Sales", new[]
    {
        new object?[] { "Region", "Total" },
        new object?[] { "North", 1200 },
        new object?[] { "South", 950 },
    });

    private static byte[] OneRow(params object?[] values)
        => WorkbookEditor.Create("S", new[] { values });

    private static IXLCell Cell(XLWorkbook workbook, string reference)
        => workbook.Worksheet("S").Cell(reference);

    private static XLWorkbook Open(byte[] xlsx) => new(new MemoryStream(xlsx));

    /// <summary>Data type and text of the first row's cells - enough to spot "landed as text".</summary>
    private static (XLDataType Type, string Text)[] FirstRow(byte[] xlsx, int width)
    {
        using var workbook = Open(xlsx);
        return Enumerable.Range(1, width)
            .Select(c => workbook.Worksheet("S").Cell(1, c))
            .Select(c => (c.DataType, c.GetString()))
            .ToArray();
    }

    [Fact]
    public void Create_ProducesAReadableWorkbook()
    {
        var xlsx = SampleWorkbook();

        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, xlsx.Take(4).ToArray());
        Assert.Equal("Region", WorkbookEditor.ReadCell(xlsx, "Sales", "A1"));
        Assert.Equal("North", WorkbookEditor.ReadCell(xlsx, "Sales", "A2"));
        Assert.Equal("1200", WorkbookEditor.ReadCell(xlsx, "Sales", "B2"));
    }

    [Fact]
    public void SetCell_EditsAnExistingWorkbook()
    {
        var edited = WorkbookEditor.SetCell(SampleWorkbook(), "Sales", "B2", 1500);

        Assert.Equal("1500", WorkbookEditor.ReadCell(edited, "Sales", "B2"));
        Assert.Equal("South", WorkbookEditor.ReadCell(edited, "Sales", "A3"));
    }

    [Fact]
    public void ReadCell_ThrowsForAMissingSheet()
    {
        Assert.Throws<DocumentConversionException>(
            () => WorkbookEditor.ReadCell(SampleWorkbook(), "NoSuchSheet", "A1"));
    }

    // ---------------------------------------------------------------------------------------
    // Argument validation (I-7): these used to surface as NullReferenceException.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void ReadCell_RejectsNullSheetName()
        => Assert.Throws<ArgumentNullException>(
            () => WorkbookEditor.ReadCell(SampleWorkbook(), null!, "A1"));

    [Fact]
    public void ReadCell_RejectsNullCellReference()
        => Assert.Throws<ArgumentNullException>(
            () => WorkbookEditor.ReadCell(SampleWorkbook(), "Sales", null!));

    [Fact]
    public void SetCell_RejectsNullSheetName()
        => Assert.Throws<ArgumentNullException>(
            () => WorkbookEditor.SetCell(SampleWorkbook(), null!, "A1", 1));

    [Fact]
    public void SetCell_RejectsBlankCellReference()
        => Assert.Throws<ArgumentException>(
            () => WorkbookEditor.SetCell(SampleWorkbook(), "Sales", "   ", 1));

    [Fact]
    public void Create_RejectsANullRow()
        => Assert.Throws<ArgumentException>(
            () => WorkbookEditor.Create("S", new IEnumerable<object?>[] { null! }));

    [Fact]
    public void Create_WithASheetNameExcelRejects_FailsFastWithArgumentException()
    {
        // Previously these reached ClosedXML and came back as DocumentConversionException wrapping
        // someone else's message. Failing fast names the parameter and says what is wrong.
        var rows = new[] { new object?[] { 1 } };

        var tooLong = Assert.Throws<ArgumentException>(
            () => WorkbookEditor.Create(new string('a', 32), rows));
        Assert.Equal("sheetName", tooLong.ParamName);
        Assert.Contains("32 characters", tooLong.Message, StringComparison.Ordinal);

        var badChar = Assert.Throws<ArgumentException>(
            () => WorkbookEditor.Create("Q1/Q2", rows));
        Assert.Equal("sheetName", badChar.ParamName);
        Assert.Contains("'/'", badChar.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Create_WithASheetNameOfExactlyThirtyOneCharacters_IsAccepted()
    {
        // The boundary is inclusive; 31 is legal, 32 is not.
        var xlsx = WorkbookEditor.Create(new string('a', 31), new[] { new object?[] { 1 } });
        Assert.Equal(new string('a', 31), Assert.Single(WorkbookEditor.SheetNames(xlsx)));
    }

    // ---------------------------------------------------------------------------------------
    // Create(IEnumerable<XlsxSheet>) and its overloads
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void Create_WithSeveralSheets_KeepsSequenceOrderAndContent()
    {
        var xlsx = WorkbookEditor.Create(new[]
        {
            XlsxSheet.Named("Summary", new[] { new object?[] { "Region", "Total" } }),
            XlsxSheet.Named("Detail",  new[] { new object?[] { "EMEA", 1200 } }),
            XlsxSheet.Named("Notes",   new[] { new object?[] { "n/a" } }),
        });

        // Order is asserted through SheetNames, which reads tab order - not through part order,
        // which can disagree.
        Assert.Equal(new[] { "Summary", "Detail", "Notes" }, WorkbookEditor.SheetNames(xlsx));
        Assert.Equal("Region", WorkbookEditor.ReadCell(xlsx, "Summary", "A1"));
        Assert.Equal("1200", WorkbookEditor.ReadCell(xlsx, "Detail", "B1"));
    }

    [Fact]
    public void Create_WithSheets_WritesFormulasFoundInRows()
    {
        var xlsx = WorkbookEditor.Create(new[]
        {
            XlsxSheet.Named("Summary", new[]
            {
                new object?[] { 1200, 1400, XlsxFormula.From("=A1+B1") },
            }),
        });

        Assert.Equal("2600", WorkbookEditor.ReadCell(xlsx, "Summary", "C1"));
    }

    [Fact]
    public void Create_WithSheets_RejectsEmptyNullAndDuplicates()
    {
        // Named argument, not a cast: Create is overloaded, and naming the parameter pins which
        // overload this asserts against without the redundant upcast CodeQL flags (cs/useless-upcast).
        Assert.Throws<ArgumentNullException>(() => WorkbookEditor.Create(sheets: null!));

        // A workbook with no worksheets is not a valid .xlsx - measured, ClosedXML throws
        // "Workbooks need at least one worksheet." Rejecting eagerly names the parameter.
        var empty = Assert.Throws<ArgumentException>(() => WorkbookEditor.Create(Array.Empty<XlsxSheet>()));
        Assert.Equal("sheets", empty.ParamName);

        var nullElement = Assert.Throws<ArgumentException>(
            () => WorkbookEditor.Create(new XlsxSheet[] { null! }));
        Assert.Equal("sheets", nullElement.ParamName);
        Assert.Contains("Sheet 1 was null.", nullElement.Message, StringComparison.Ordinal);

        var rows = new[] { new object?[] { 1 } };
        var duplicate = Assert.Throws<ArgumentException>(() => WorkbookEditor.Create(new[]
        {
            XlsxSheet.Named("Data", rows),
            XlsxSheet.Named("data", rows),   // Excel treats sheet names case-insensitively
        }));
        Assert.Equal("sheets", duplicate.ParamName);
        Assert.Contains("more than once", duplicate.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task CreateAsync_WithSheets_MatchesTheByteArrayOverload()
    {
        var sheets = new[]
        {
            XlsxSheet.Named("Summary", new[] { new object?[] { "Region", 1 } }),
            XlsxSheet.Named("Detail",  new[] { new object?[] { "EMEA", 2 } }),
        };

        var expected = WorkbookEditor.Create(sheets);

        using var destination = new MemoryStream();
        await WorkbookEditor.CreateAsync(sheets, destination);

        var streamed = destination.ToArray();

        // B16: literals first. SheetNames == SheetNames holds on two empty lists and
        // ReadCell == ReadCell holds on two empty strings, so neither line could tell a workbook
        // that was written from one that was not.
        Assert.Equal(new[] { "Summary", "Detail" }, WorkbookEditor.SheetNames(streamed));
        Assert.Equal("EMEA", WorkbookEditor.ReadCell(streamed, "Detail", "A1"));

        Assert.Equal(WorkbookEditor.SheetNames(expected), WorkbookEditor.SheetNames(streamed));
        Assert.Equal(
            WorkbookEditor.ReadCell(expected, "Detail", "A1"),
            WorkbookEditor.ReadCell(streamed, "Detail", "A1"));
    }

    [Fact]
    public async Task CreateToFileAsync_WithSheets_WritesAReadableWorkbook()
    {
        using var temp = new TempFile();

        await WorkbookEditor.CreateToFileAsync(
            new[] { XlsxSheet.Named("Summary", new[] { new object?[] { "hi" } }) },
            temp.Path);

        Assert.Equal("hi", WorkbookEditor.ReadCell(await File.ReadAllBytesAsync(temp.Path), "Summary", "A1"));
    }

    // ---------------------------------------------------------------------------------------
    // Error handling (I-6): raw FileFormatException/ArgumentException used to escape.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void ReadCell_WrapsCorruptInputInDocumentConversionException()
        => Assert.Throws<DocumentConversionException>(
            () => WorkbookEditor.ReadCell(new byte[] { 1, 2, 3, 4, 5 }, "Sales", "A1"));

    [Fact]
    public void SetCell_WrapsAnInvalidCellReferenceInDocumentConversionException()
        => Assert.Throws<DocumentConversionException>(
            () => WorkbookEditor.SetCell(SampleWorkbook(), "Sales", "not-a-reference", 1));

    // ---------------------------------------------------------------------------------------
    // Typed and culture gaps in SetCellValue.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void Create_StoresUnsignedAndSignedByteIntegersAsNumbers()
    {
        // These fell through to ToString() and landed as text, so SUM() skipped them.
        var xlsx = OneRow((uint)11, (ulong)12, (ushort)13, (sbyte)14);

        Assert.Equal(
            new[]
            {
                (XLDataType.Number, "11"),
                (XLDataType.Number, "12"),
                (XLDataType.Number, "13"),
                (XLDataType.Number, "14"),
            },
            FirstRow(xlsx, 4));
    }

    [Fact]
    public void Create_StoresDateOnlyTimeOnlyAndTimeSpanAsTemporalValuesNotText()
    {
        var xlsx = OneRow(new DateOnly(2026, 7, 29), new TimeOnly(13, 45), TimeSpan.FromMinutes(90));

        using var workbook = Open(xlsx);
        Assert.Equal(XLDataType.DateTime, Cell(workbook, "A1").DataType);
        Assert.Equal(new DateTime(2026, 7, 29), Cell(workbook, "A1").GetDateTime());
        Assert.Equal(XLDataType.TimeSpan, Cell(workbook, "B1").DataType);
        Assert.Equal(new TimeSpan(13, 45, 0), Cell(workbook, "B1").GetTimeSpan());
        Assert.Equal(XLDataType.TimeSpan, Cell(workbook, "C1").DataType);
        Assert.Equal(TimeSpan.FromMinutes(90), Cell(workbook, "C1").GetTimeSpan());
    }

    [Fact]
    public void Create_WritesTheSameValuesWhateverTheCurrentCulture()
    {
        // Built from InvariantCulture rather than looked up, so the test behaves identically on a
        // machine running in globalization-invariant mode.
        var german = (CultureInfo)CultureInfo.InvariantCulture.Clone();
        german.DateTimeFormat.ShortDatePattern = "dd.MM.yyyy";
        german.DateTimeFormat.LongTimePattern = "HH:mm:ss";
        german.NumberFormat.NumberDecimalSeparator = ",";
        german.NumberFormat.NumberGroupSeparator = ".";

        object?[] Row() => new object?[]
        {
            new DateOnly(2026, 7, 29),
            TimeSpan.FromMinutes(90),
            new DateTimeOffset(2026, 7, 29, 10, 30, 0, TimeSpan.Zero),
            1234.5m,
        };

        var underInvariant = WithCulture(CultureInfo.InvariantCulture, () => OneRow(Row()));
        var underGerman = WithCulture(german, () => OneRow(Row()));

        // Pin FirstRow itself before comparing two of its results (B16). The comparison below is
        // FirstRow(a) against FirstRow(b) - the helper against itself - which holds however broken
        // the helper is. If it returned four empty strings, or an empty array, the culture
        // guarantee would be unverified and nothing would say so. This is the same shape that let
        // A26 hide behind Assert.Contains for eight releases.
        var invariantRow = FirstRow(underInvariant, 4);
        Assert.Equal(4, invariantRow.Length);
        Assert.All(invariantRow, cell => Assert.False(string.IsNullOrWhiteSpace(cell.Text),
            "FirstRow returned a blank cell, so the comparison below would prove nothing"));

        // Both read back under the ambient test culture, so any difference is baked into the file.
        Assert.Equal(invariantRow, FirstRow(underGerman, 4));
    }

    private static T WithCulture<T>(CultureInfo culture, Func<T> body)
    {
        var previous = CultureInfo.CurrentCulture;
        CultureInfo.CurrentCulture = culture;
        try { return body(); }
        finally { CultureInfo.CurrentCulture = previous; }
    }

    // ---------------------------------------------------------------------------------------
    // SheetNames
    // ---------------------------------------------------------------------------------------

    /// <summary>
    /// A workbook with three sheets, the middle one hidden, built directly with ClosedXML because
    /// WorkbookEditor.Create only makes single-sheet workbooks.
    ///
    /// Added in the order Third, Hidden, First — deliberately not their tab order — and then
    /// repositioned to tab order First, Hidden, Third via the settable <c>Position</c>, so the
    /// fixture states its intent explicitly instead of getting tab order "for free" from insertion
    /// order.
    ///
    /// This alone does not make the ordering test below able to catch a missing
    /// <c>OrderBy(sheet => sheet.Position)</c> in <c>SheetNamesCore</c>, and no fixture reachable
    /// through the public byte[]/Stream API can: verified against ClosedXML 0.105.1's own source,
    /// <c>XLWorkbook_Save.cs</c> always writes &lt;sheet&gt; elements sorted by Position
    /// (`WorksheetsInternal.Cast&lt;XLWorksheet&gt;().OrderBy(w => w.Position)`), and
    /// <c>XLWorkbook_Load.cs</c> assigns Position as a plain running counter over that same
    /// document order and adds sheets to its internal dictionary in that order. So after any
    /// save/load round trip — which is the only way <c>SheetNamesCore</c> ever sees a workbook —
    /// raw <c>Worksheets</c> enumeration order and Position order are identical by construction;
    /// there is no way to encode a workbook where they'd differ. Confirmed with a throwaway probe:
    /// removing the <c>OrderBy</c> left every test in this class passing. The <c>OrderBy</c> stays
    /// for defense-in-depth and as documentation of the guarantee, even though this test suite
    /// cannot presently exercise its absence.
    /// </summary>
    private static byte[] ThreeSheetWorkbook()
    {
        using var workbook = new XLWorkbook();
        var third = workbook.Worksheets.Add("Third");
        third.Cell("A1").Value = "c";
        var hidden = workbook.Worksheets.Add("Hidden");
        hidden.Cell("A1").Value = "b";
        hidden.Hide();
        var first = workbook.Worksheets.Add("First");
        first.Cell("A1").Value = "a";

        first.Position = 1;
        hidden.Position = 2;
        third.Position = 3;

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        return ms.ToArray();
    }

    [Fact]
    public void SheetNames_ReturnsEverySheetInTabOrder_IncludingHiddenOnes()
    {
        Assert.Equal(
            new[] { "First", "Hidden", "Third" },
            WorkbookEditor.SheetNames(ThreeSheetWorkbook()));
    }

    [Fact]
    public void SheetNames_RejectsMissingContent()
    {
        Assert.Throws<ArgumentNullException>(() => WorkbookEditor.SheetNames(null!));
        Assert.Throws<ArgumentException>(() => WorkbookEditor.SheetNames(Array.Empty<byte>()));
    }

    [Fact]
    public void SheetNames_WrapsAFileThatIsNotAWorkbook()
    {
        Assert.Throws<DocumentConversionException>(
            () => WorkbookEditor.SheetNames(new byte[] { 1, 2, 3, 4 }));
    }

    [Fact]
    public async Task SheetNamesAsync_AgreesWithTheByteArrayOverload()
    {
        var xlsx = ThreeSheetWorkbook();
        using var source = new MemoryStream(xlsx);

        Assert.Equal(WorkbookEditor.SheetNames(xlsx), await WorkbookEditor.SheetNamesAsync(source));
    }

    // ---------------------------------------------------------------------------------------
    // ReadSheet
    // ---------------------------------------------------------------------------------------

    /// <summary>A workbook whose data deliberately does not start at A1.</summary>
    private static byte[] OffsetWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Data");
        sheet.Cell("C3").Value = "Region";
        sheet.Cell("D3").Value = "Total";
        sheet.Cell("C4").Value = "North";
        sheet.Cell("D4").Value = 1200;

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        return ms.ToArray();
    }

    [Fact]
    public void ReadSheet_ReturnsTheRowsItWasCreatedWith()
    {
        var rows = WorkbookEditor.ReadSheet(SampleWorkbook(), "Sales");

        Assert.Equal(3, rows.Count);
        Assert.Equal(new[] { "Region", "Total" }, rows[0]);
        Assert.Equal(new[] { "North", "1200" }, rows[1]);
        Assert.Equal(new[] { "South", "950" }, rows[2]);
    }

    /// <summary>
    /// The result is anchored at A1, not at the first used cell, so rows[r][c] means what a caller
    /// reading a spreadsheet would expect. Implementing this by iterating the used range instead
    /// puts "Region" at rows[0][0] and looks entirely plausible.
    /// </summary>
    [Fact]
    public void ReadSheet_AnchorsTheResultAtA1_WhenTheDataStartsElsewhere()
    {
        var rows = WorkbookEditor.ReadSheet(OffsetWorkbook(), "Data");

        Assert.Equal(4, rows.Count);
        Assert.Equal("Region", rows[2][2]);
        Assert.Equal("Total", rows[2][3]);
        Assert.Equal("North", rows[3][2]);
        Assert.Equal("1200", rows[3][3]);

        Assert.Equal(new[] { "", "", "", "" }, rows[0]);
        Assert.Equal(new[] { "", "", "", "" }, rows[1]);
    }

    [Fact]
    public void ReadSheet_PadsEveryRowToTheSameWidth()
    {
        var xlsx = WorkbookEditor.Create("S", new[]
        {
            new object?[] { "a", "b", "c" },
            new object?[] { "d" },
        });

        var rows = WorkbookEditor.ReadSheet(xlsx, "S");

        Assert.All(rows, row => Assert.Equal(3, row.Count));
        Assert.Equal(new[] { "d", "", "" }, rows[1]);
    }

    /// <summary>
    /// Blank rows inside the range are kept. Dropping them would shift every later index and
    /// silently break the positional guarantee the whole design rests on.
    /// </summary>
    [Fact]
    public void ReadSheet_KeepsBlankRowsInsideTheRange()
    {
        var xlsx = WorkbookEditor.Create("S", new[]
        {
            new object?[] { "a" },
            new object?[] { null },
            new object?[] { "c" },
        });

        var rows = WorkbookEditor.ReadSheet(xlsx, "S");

        Assert.Equal(3, rows.Count);
        Assert.Equal(new[] { "" }, rows[1]);
        Assert.Equal(new[] { "c" }, rows[2]);
    }

    [Fact]
    public void ReadSheet_ReturnsAnEmptyListForAnEmptySheet()
    {
        using var workbook = new XLWorkbook();
        workbook.Worksheets.Add("Blank");
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);

        Assert.Empty(WorkbookEditor.ReadSheet(ms.ToArray(), "Blank"));
    }

    /// <summary>
    /// "Used" must mean "has a value", never "has formatting" — otherwise one bolded empty cell
    /// pads every row out to it for no reason a caller could see.
    /// </summary>
    [Fact]
    public void ReadSheet_IgnoresCellsThatAreOnlyFormatted()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("S");
        sheet.Cell("A1").Value = "a";
        sheet.Cell("Z1").Style.Font.Bold = true;      // formatting only, no value
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);

        var rows = WorkbookEditor.ReadSheet(ms.ToArray(), "S");

        Assert.Single(rows);
        Assert.Equal(new[] { "a" }, rows[0]);
    }

    /// <summary>
    /// Builds a workbook where C1's cached formula value and a fresh evaluation of the same
    /// formula deliberately disagree, by editing the saved .xlsx package's XML directly: A1 is
    /// changed from 2 to 10 <i>after</i> save, while C1 keeps the formula "A1+B1" and a cached
    /// value of 5 - the answer for the original A1=2, B1=3, not the tampered A1=10. Reading it
    /// back can only produce "5" by using the cache; evaluating "A1+B1" fresh would read "13".
    ///
    /// This has to be done by hand because ClosedXML's own writer does not always emit a cached
    /// value for a formula cell in the first place - confirmed by dumping the saved XML for
    /// <c>sheet.Cell("C1").FormulaA1 = "A1+B1"</c>: it writes
    /// <c>&lt;x:c r="C1"&gt;&lt;x:f&gt;A1+B1&lt;/x:f&gt;&lt;/x:c&gt;</c>, no <c>&lt;x:v&gt;</c> at
    /// all. There is no cached value to leave stale through the public ClosedXML API alone, so the
    /// .xlsx (a zip) is opened directly and its <c>xl/worksheets/sheet1.xml</c> entry is rewritten.
    /// </summary>
    private static byte[] FormulaWorkbookWithStaleCachedValue()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("S");
        sheet.Cell("A1").Value = 2;
        sheet.Cell("B1").Value = 3;
        sheet.Cell("C1").FormulaA1 = "A1+B1";
        using var built = new MemoryStream();
        workbook.SaveAs(built);

        using var package = new MemoryStream();
        built.Position = 0;
        built.CopyTo(package);
        package.Position = 0;

        using (var zip = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml")
                ?? throw new InvalidOperationException(
                    "xl/worksheets/sheet1.xml was not found in the fixture; ClosedXML's package layout changed.");

            string xml;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                xml = reader.ReadToEnd();

            var withStaleA1 = xml.Replace(
                "<x:c r=\"A1\" s=\"0\"><x:v>2</x:v></x:c>",
                "<x:c r=\"A1\" s=\"0\"><x:v>10</x:v></x:c>");
            var withCachedC1 = withStaleA1.Replace(
                "<x:c r=\"C1\" s=\"0\"><x:f>A1+B1</x:f></x:c>",
                "<x:c r=\"C1\" s=\"0\"><x:f>A1+B1</x:f><x:v>5</x:v></x:c>");

            // Each replacement must be checked independently - checking only the combined result
            // would stay silent if just the A1 replacement stopped matching (ClosedXML changing
            // only the value-cell shape, e.g. t="n" or a different style index) while the C1
            // replacement still landed. That leaves A1 at its original 2, so the cached 5 on C1
            // would equal a fresh evaluation of A1+B1, and the test would pass having proven
            // nothing about honouring the cache - silently reinstating the exact defect this
            // fixture exists to catch.
            if (withStaleA1 == xml)
                throw new InvalidOperationException(
                    "The A1 replacement did not match; ClosedXML's value-cell output shape changed. " +
                    "Without it the cached and evaluated values agree and this fixture proves nothing.");
            if (withCachedC1 == withStaleA1)
                throw new InvalidOperationException(
                    "The C1 replacement did not match; ClosedXML's formula-cell output shape changed.");

            entry.Delete();
            var replacement = zip.CreateEntry("xl/worksheets/sheet1.xml");
            using var writer = new StreamWriter(replacement.Open(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            writer.Write(withCachedC1);
        }

        return package.ToArray();
    }

    /// <summary>
    /// Nothing in this library evaluates formulas; a formula cell reads back its cached value.
    /// The fixture's cached value (5) and a fresh evaluation of its formula (13) are made to
    /// disagree, so only cache-honouring behaviour can pass - see
    /// <see cref="FormulaWorkbookWithStaleCachedValue"/> for why a fixture built solely through
    /// ClosedXML's public API cannot tell the two apart.
    /// </summary>
    [Fact]
    public void ReadSheet_ReturnsTheCachedValueOfAFormulaCell()
    {
        Assert.Equal("5", WorkbookEditor.ReadSheet(FormulaWorkbookWithStaleCachedValue(), "S")[0][2]);
    }

    /// <summary>
    /// XFD1048576 is not a made-up extreme — it is Excel's own maximum address (16,384 columns,
    /// 1,048,576 rows). One stray value there, with nothing filled in between, still describes a
    /// 1,048,576 x 16,384 = 17,179,869,184-cell rectangle, because ReadSheetCore's extent comes
    /// from the single farthest-out used cell, not from how much of the sheet actually holds
    /// data. Materialising that many string-array slots would exhaust memory long before this
    /// method returns, so ReadSheet must refuse before it allocates anything — and the fixture
    /// itself stays a few KB on disk, since ClosedXML only ever writes the cells that were set.
    /// </summary>
    [Fact]
    public void ReadSheet_ThrowsRatherThanAllocateForAFarFlungStrayValue()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sales");
        sheet.Cell("A1").Value = "a";
        sheet.Cell("XFD1048576").Value = "x";
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);

        var ex = Assert.Throws<DocumentConversionException>(
            () => WorkbookEditor.ReadSheet(ms.ToArray(), "Sales"));

        // The message must name the actual extent and the cap - something a caller can act on -
        // not a bare "failed to read" that gives no hint memory, rather than the file, was the
        // problem.
        Assert.Contains("Sales", ex.Message);
        Assert.Contains("1048576", ex.Message);
        Assert.Contains("16384", ex.Message);
        Assert.Contains("17179869184", ex.Message);
        Assert.Contains("2000000", ex.Message);
    }

    /// <summary>Ordinary-sized data, comfortably under the 2,000,000-cell cap, must still work.</summary>
    [Fact]
    public void ReadSheet_ReadsASheetComfortablyUnderTheCellCap()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Big");
        for (var r = 1; r <= 500; r++)
            for (var c = 1; c <= 20; c++)
                sheet.Cell(r, c).Value = $"{r}-{c}";
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);

        var rows = WorkbookEditor.ReadSheet(ms.ToArray(), "Big");

        Assert.Equal(500, rows.Count);
        Assert.All(rows, row => Assert.Equal(20, row.Count));
        Assert.Equal("1-1", rows[0][0]);
        Assert.Equal("500-20", rows[499][19]);
    }

    [Fact]
    public void ReadSheet_ThrowsWhenTheSheetDoesNotExist()
    {
        Assert.Throws<DocumentConversionException>(
            () => WorkbookEditor.ReadSheet(SampleWorkbook(), "Nope"));
    }

    [Fact]
    public void ReadSheet_RejectsMissingArguments()
    {
        Assert.Throws<ArgumentNullException>(() => WorkbookEditor.ReadSheet(null!, "Sales"));
        Assert.Throws<ArgumentNullException>(() => WorkbookEditor.ReadSheet(SampleWorkbook(), null!));
        Assert.Throws<ArgumentException>(() => WorkbookEditor.ReadSheet(Array.Empty<byte>(), "Sales"));
        Assert.Throws<ArgumentException>(() => WorkbookEditor.ReadSheet(SampleWorkbook(), " "));
    }

    [Fact]
    public async Task ReadSheetAsync_AgreesWithTheByteArrayOverload()
    {
        var xlsx = OffsetWorkbook();
        using var source = new MemoryStream(xlsx);

        var expected = WorkbookEditor.ReadSheet(xlsx, "Data");
        var actual = await WorkbookEditor.ReadSheetAsync(source, "Data");

        Assert.Equal(expected, actual);
    }

    // ---------------------------------------------------------------------------------------
    // File-path overloads
    // ---------------------------------------------------------------------------------------

    [Fact]
    public async Task CreateToFileAsync_WritesAWorkbookTheOtherOverloadsCanRead()
    {
        using var output = new TempFile();

        await WorkbookEditor.CreateToFileAsync("Sales", new[]
        {
            new object?[] { "Region", "Total" },
            new object?[] { "North", 1200 },
        }, output.Path);

        Assert.Equal("1200", await WorkbookEditor.ReadCellAsync(output.Path, "Sales", "B2"));
        Assert.Equal(new[] { "Sales" }, await WorkbookEditor.SheetNamesAsync(output.Path));
    }

    [Fact]
    public async Task SetCellAsync_FromFileToFile_MatchesTheByteArrayOverload()
    {
        var xlsx = SampleWorkbook();

        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, xlsx);

        await WorkbookEditor.SetCellAsync(input.Path, output.Path, "Sales", "B2", 1500);

        Assert.Equal(
            WorkbookEditor.ReadCell(WorkbookEditor.SetCell(xlsx, "Sales", "B2", 1500), "Sales", "B2"),
            await WorkbookEditor.ReadCellAsync(output.Path, "Sales", "B2"));
    }

    [Fact]
    public async Task ReadSheetAsync_FromFile_MatchesTheByteArrayOverload()
    {
        var xlsx = SampleWorkbook();

        using var input = new TempFile();
        await File.WriteAllBytesAsync(input.Path, xlsx);

        Assert.Equal(
            WorkbookEditor.ReadSheet(xlsx, "Sales"),
            await WorkbookEditor.ReadSheetAsync(input.Path, "Sales"));
    }

    /// <summary>
    /// Exercised directly rather than only transitively through
    /// <see cref="CreateToFileAsync_WritesAWorkbookTheOtherOverloadsCanRead"/>, so a bug here cannot
    /// be mis-attributed to <c>CreateToFileAsync</c>, and two compensating bugs cannot mask each
    /// other.
    /// </summary>
    [Fact]
    public async Task SheetNamesAsync_FromFile_MatchesTheByteArrayOverload()
    {
        var xlsx = ThreeSheetWorkbook();

        using var input = new TempFile();
        await File.WriteAllBytesAsync(input.Path, xlsx);

        Assert.Equal(
            WorkbookEditor.SheetNames(xlsx),
            await WorkbookEditor.SheetNamesAsync(input.Path));
    }

    [Fact]
    public void SetCell_WithAnXlsxFormula_WritesAFormulaThatReadsBackComputed()
    {
        var xlsx = WorkbookEditor.Create("Sheet1", new[]
        {
            new object?[] { 1 },
            new object?[] { 2 },
        });

        var edited = WorkbookEditor.SetCell(xlsx, "Sheet1", "B1", XlsxFormula.From("=A1+A2"));

        // ClosedXML evaluates on demand, so this package's own reader returns the computed value.
        // Measured 2026-08-08; see docs/superpowers/specs/2026-08-08-xlsx-multisheet-design.md.
        Assert.Equal("3", WorkbookEditor.ReadCell(edited, "Sheet1", "B1"));
    }

    [Theory]
    [InlineData("=1/0", "#DIV/0!")]
    [InlineData("=TOTALLYNOTAFUNCTION(A1)", "#NAME?")]
    [InlineData("=Sheet2!A1", "#REF!")]
    public void SetCell_WithABrokenFormula_ReadsBackAsTheExcelErrorString(string formula, string expected)
    {
        // The error contract: a broken formula surfaces exactly as Excel shows it, and nothing
        // throws. Pinned because it is what a consumer actually hits.
        var xlsx = WorkbookEditor.Create("Sheet1", new[] { new object?[] { 5 } });

        var edited = WorkbookEditor.SetCell(xlsx, "Sheet1", "B1", XlsxFormula.From(formula));

        Assert.Equal(expected, WorkbookEditor.ReadCell(edited, "Sheet1", "B1"));
    }

    [Fact]
    public void SetCell_WithAFormula_WritesNoCachedValue()
    {
        // The documented limit, asserted against the raw XML so a ClosedXML upgrade that starts
        // caching values cannot change the promise silently.
        var xlsx = WorkbookEditor.Create("Sheet1", new[] { new object?[] { 1 }, new object?[] { 2 } });
        var edited = WorkbookEditor.SetCell(xlsx, "Sheet1", "B1", XlsxFormula.From("=A1+A2"));

        using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(edited));
        var entry = zip.Entries.Single(e => e.FullName.EndsWith("sheet1.xml", StringComparison.Ordinal));
        using var reader = new StreamReader(entry.Open());
        var xml = reader.ReadToEnd();

        var cell = xml[xml.IndexOf("r=\"B1\"", StringComparison.Ordinal)..];
        cell = cell[..cell.IndexOf("</x:c>", StringComparison.Ordinal)];

        Assert.Contains(":f>A1+A2<", cell, StringComparison.Ordinal);
        Assert.DoesNotContain(":v>", cell, StringComparison.Ordinal);
    }

    [Fact]
    public void XlsxFormula_StripsALeadingEquals_SoBothSpellingsAgree()
    {
        Assert.Equal("SUM(A1:A2)", XlsxFormula.From("=SUM(A1:A2)").Formula);
        Assert.Equal("SUM(A1:A2)", XlsxFormula.From("SUM(A1:A2)").Formula);
    }

    [Fact]
    public void XlsxFormula_RejectsNullAndEmpty()
    {
        Assert.Throws<ArgumentNullException>(() => XlsxFormula.From(null!));
        Assert.Equal("formula", Assert.Throws<ArgumentException>(() => XlsxFormula.From("   ")).ParamName);
        Assert.Equal("formula", Assert.Throws<ArgumentException>(() => XlsxFormula.From("=")).ParamName);
    }

    [Fact]
    public void XlsxSheet_Named_DeepCopiesRowsSoLaterMutationDoesNotAffectTheSheet()
    {
        var row1 = new object?[] { "EMEA", 1200 };
        var rows = new List<IEnumerable<object?>>
        {
            new object?[] { "Region", "Total" },
            row1,
        };

        var sheet = XlsxSheet.Named("Summary", rows);

        // Mutate the caller's data after construction - the outer sequence (replace an element
        // wholesale) and a row's own array (mutate an element inside it). If either mutation were
        // visible in sheet.Rows, this proves materialisation was shallow rather than deep; a test
        // that never mutates after construction cannot tell eager-but-shallow apart from deep.
        rows[1] = new object?[] { "REPLACED", -1 };
        row1[0] = "MUTATED";
        row1[1] = -999;

        Assert.Equal("Summary", sheet.Name);
        Assert.Equal(2, sheet.Rows.Count);
        Assert.Equal(new object?[] { "EMEA", 1200 }, sheet.Rows[1]);
    }

    [Fact]
    public void XlsxSheet_Named_RejectsBadInput()
    {
        var rows = new[] { new object?[] { 1 } };

        Assert.Throws<ArgumentNullException>(() => XlsxSheet.Named(null!, rows));
        Assert.Throws<ArgumentNullException>(() => XlsxSheet.Named("Sheet1", null!));

        // Reuses the same rules as WorkbookEditor.Create - one helper, so they cannot drift.
        Assert.Equal("name", Assert.Throws<ArgumentException>(
            () => XlsxSheet.Named(new string('a', 32), rows)).ParamName);
        Assert.Equal("name", Assert.Throws<ArgumentException>(
            () => XlsxSheet.Named("Q1/Q2", rows)).ParamName);

        var nullRow = Assert.Throws<ArgumentException>(
            () => XlsxSheet.Named("Sheet1", new IEnumerable<object?>[] { null! }));
        Assert.Equal("rows", nullRow.ParamName);
        Assert.Contains("Row 1 was null.", nullRow.Message, StringComparison.Ordinal);
    }

    // ---------------------------------------------------------------------------------------
    // AppendRows and its overloads
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void AppendRows_AddsAfterTheLastUsedRow()
    {
        var xlsx = WorkbookEditor.Create("Log", new[]
        {
            new object?[] { "when", "what" },
            new object?[] { "09:00", "started" },
        });

        var appended = WorkbookEditor.AppendRows(xlsx, "Log", new[]
        {
            new object?[] { "09:05", "continued" },
            new object?[] { "09:10", "finished" },
        });

        var sheet = WorkbookEditor.ReadSheet(appended, "Log");
        Assert.Equal(4, sheet.Count);
        Assert.Equal(new[] { "09:05", "continued" }, sheet[2]);
        Assert.Equal(new[] { "09:10", "finished" }, sheet[3]);
    }

    [Fact]
    public void AppendRows_LeavesOtherSheetsAlone()
    {
        // "The appended rows are right" would pass while a second sheet silently vanished, so the
        // other sheet is asserted directly.
        var xlsx = WorkbookEditor.Create(new[]
        {
            XlsxSheet.Named("Log",   new[] { new object?[] { "start" } }),
            XlsxSheet.Named("Notes", new[] { new object?[] { "keep me" } }),
        });

        var appended = WorkbookEditor.AppendRows(xlsx, "Log", new[] { new object?[] { "more" } });

        Assert.Equal(new[] { "Log", "Notes" }, WorkbookEditor.SheetNames(appended));
        Assert.Equal("keep me", WorkbookEditor.ReadCell(appended, "Notes", "A1"));
    }

    [Fact]
    public void AppendRows_WithFormulas_WritesThemAsFormulas()
    {
        var xlsx = WorkbookEditor.Create("Data", new[] { new object?[] { 10, 20 } });

        var appended = WorkbookEditor.AppendRows(xlsx, "Data", new[]
        {
            new object?[] { 30, 40, XlsxFormula.From("=A2+B2") },
        });

        // The formula references the appended row's own cells (30 + 40 = 70), a value that
        // matches no literal in the row. A formula written as the literal 30 - or any other
        // literal already present - would read back as that literal, not "70", so this can only
        // pass if C2 is a real, evaluated formula.
        Assert.Equal("30", WorkbookEditor.ReadCell(appended, "Data", "A2"));
        Assert.Equal("70", WorkbookEditor.ReadCell(appended, "Data", "C2"));
    }

    [Fact]
    public void AppendRows_WithNoRows_IsANoOpThatStillReturnsAValidWorkbook()
    {
        var xlsx = WorkbookEditor.Create("Log", new[] { new object?[] { "only" } });

        var appended = WorkbookEditor.AppendRows(xlsx, "Log", Array.Empty<IEnumerable<object?>>());

        Assert.Single(WorkbookEditor.ReadSheet(appended, "Log"));
        Assert.Equal("only", WorkbookEditor.ReadCell(appended, "Log", "A1"));
    }

    [Fact]
    public void AppendRows_RejectsBadInput()
    {
        var xlsx = WorkbookEditor.Create("Log", new[] { new object?[] { 1 } });
        var rows = new[] { new object?[] { 2 } };

        Assert.Throws<ArgumentNullException>(() => WorkbookEditor.AppendRows(null!, "Log", rows));
        Assert.Throws<ArgumentNullException>(() => WorkbookEditor.AppendRows(xlsx, "Log", null!));
        Assert.Equal("sheetName", Assert.Throws<ArgumentException>(
            () => WorkbookEditor.AppendRows(xlsx, "  ", rows)).ParamName);

        // A missing sheet is a DocumentConversionException, matching the existing Sheet() helper
        // that ReadSheet, ReadCell and SetCell all use. One convention for one mistake.
        var missing = Assert.Throws<DocumentConversionException>(
            () => WorkbookEditor.AppendRows(xlsx, "Nope", rows));
        Assert.Contains("'Nope'", missing.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AppendRows_RejectsAnEmptyWorkbook()
    {
        // Same guard, and now the same wording, as every other byte[] overload's empty-xlsx check
        // (ValidateWorkbook) - this used to say "Workbook was empty.", diverging from the rest.
        var rows = new[] { new object?[] { 1 } };

        var ex = Assert.Throws<ArgumentException>(
            () => WorkbookEditor.AppendRows(Array.Empty<byte>(), "Log", rows));

        Assert.Equal("xlsx", ex.ParamName);
        Assert.Contains("Workbook content was empty.", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AppendRows_ToASheetWithNoUsedRows_StartsAtRowOne()
    {
        // The ?? 0 branch: LastRowUsed() is null for a sheet with no used rows at all, so the
        // append must start at row 1, not row 0 or throw.
        var xlsx = WorkbookEditor.Create("Log", Array.Empty<IEnumerable<object?>>());

        var appended = WorkbookEditor.AppendRows(xlsx, "Log", new[] { new object?[] { "first" } });

        var sheet = WorkbookEditor.ReadSheet(appended, "Log");
        Assert.Single(sheet);
        Assert.Equal(new[] { "first" }, sheet[0]);
    }

    [Fact]
    public async Task AppendRowsAsync_MatchesTheByteArrayOverload()
    {
        var xlsx = WorkbookEditor.Create("Log", new[] { new object?[] { "a" } });
        var rows = new[] { new object?[] { "b" } };

        var expected = WorkbookEditor.AppendRows(xlsx, "Log", rows);

        using var source = new MemoryStream(xlsx);
        using var destination = new MemoryStream();
        await WorkbookEditor.AppendRowsAsync(source, "Log", rows, destination);

        // B16: the literal grid. Without it this test never checked that "b" arrived - two empty
        // grids compare equal, so an AppendRows that appended nothing passed.
        Assert.Equal(
            new[] { new[] { "a" }, new[] { "b" } },
            WorkbookEditor.ReadSheet(destination.ToArray(), "Log"));

        Assert.Equal(
            WorkbookEditor.ReadSheet(expected, "Log"),
            WorkbookEditor.ReadSheet(destination.ToArray(), "Log"));
    }

    [Fact]
    public async Task AppendRowsAsync_FromFileToFile_MatchesTheByteArrayOverload()
    {
        var xlsx = WorkbookEditor.Create("Log", new[] { new object?[] { "a" } });
        var rows = new[] { new object?[] { "b" } };

        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, xlsx);

        await WorkbookEditor.AppendRowsAsync(input.Path, output.Path, "Log", rows);

        // B16: see AppendRowsAsync_MatchesTheByteArrayOverload above - the parity line alone
        // holds on two empty grids.
        Assert.Equal(
            new[] { new[] { "a" }, new[] { "b" } },
            WorkbookEditor.ReadSheet(await File.ReadAllBytesAsync(output.Path), "Log"));

        Assert.Equal(
            WorkbookEditor.ReadSheet(WorkbookEditor.AppendRows(xlsx, "Log", rows), "Log"),
            WorkbookEditor.ReadSheet(await File.ReadAllBytesAsync(output.Path), "Log"));
    }

    [Fact]
    public void AddChart_AddsAChart_ThatSurvivesTheXlsxRoundTrip()
    {
        var xlsx = WorkbookEditor.Create(
            "Sheet1",
            new object[][]
            {
                new object[] { "Region", "Total" },
                new object[] { "North", 1200 },
                new object[] { "South", 980 },
            });
        var data = new ChartData(
            new[] { "North", "South" },
            new[] { new ChartSeries("Total", new double[] { 1200, 980 }) });

        var result = WorkbookEditor.AddChart(
            xlsx, "Sheet1", "B6", ChartType.ColumnClustered, data, title: "Regional Totals");

        using var source = new MemoryStream(result, writable: false);
        using var doc = OfficeIMO.Excel.ExcelDocument.Load(source);
        var sheet = doc.Sheets.First(s => s.Name == "Sheet1");
        Assert.Single(sheet.Charts);
        Assert.Equal("Regional Totals", sheet.Charts.Single().Title);
    }

    [Fact]
    public void AddChart_TheChartSurvivesTheXlsxToPdfRender()
    {
        var xlsx = WorkbookEditor.Create(
            "Sheet1",
            new object[][]
            {
                new object[] { "Region", "Total" },
                new object[] { "North", 1200 },
                new object[] { "South", 980 },
            });
        var data = new ChartData(
            new[] { "North", "South" },
            new[] { new ChartSeries("Total", new double[] { 1200, 980 }) });

        var result = WorkbookEditor.AddChart(
            xlsx, "Sheet1", "B6", ChartType.ColumnClustered, data, title: "Regional Totals");

        var text = PdfProbe.ExtractText(XlsxToPdfConverter.Convert(result));
        Assert.Contains("Regional Totals", text);
        Assert.Contains("North", text);
        Assert.Contains("South", text);
        // "North"/"South" above are satisfied by the fixture's own cell data whether or not the
        // chart itself rendered (both appear as literal cell values). "1.2k" is the value-axis
        // tick label OfficeIMO renders for the 1200 series point - it exists only because the
        // chart rendered, so this is the assertion that actually discriminates a dropped chart
        // from a present one. Measured directly in this exact scenario (see the design doc's own
        // XLSX round-trip evidence: "...Regional Totals 1.2k 900 600 300 0 North South Total").
        Assert.Contains("1.2k", text);
    }

    [Fact]
    public void AddChart_UnknownSheetName_Throws()
    {
        var xlsx = WorkbookEditor.Create("Sheet1", new object[][] { new object[] { "A" } });
        var data = new ChartData(new[] { "A" }, new[] { new ChartSeries("S", new double[] { 1 }) });

        var ex = Assert.Throws<DocumentConversionException>(
            () => WorkbookEditor.AddChart(xlsx, "NoSuchSheet", "A1", ChartType.Line, data));
        Assert.Contains("was not found", ex.Message);
    }

    [Fact]
    public void AddChart_InvalidCellRef_Throws()
    {
        var xlsx = WorkbookEditor.Create("Sheet1", new object[][] { new object[] { "A" } });
        var data = new ChartData(new[] { "A" }, new[] { new ChartSeries("S", new double[] { 1 }) });

        var ex = Assert.Throws<DocumentConversionException>(
            () => WorkbookEditor.AddChart(xlsx, "Sheet1", "not a cell", ChartType.Line, data));
        Assert.Contains("is not a valid A1-style cell reference", ex.Message);
    }

    [Fact]
    public void AddChart_AcceptsAbsoluteCellReferences()
    {
        var xlsx = WorkbookEditor.Create(
            "Sheet1", new object[][] { new object[] { "A" }, new object[] { 1 } });
        var data = new ChartData(new[] { "A" }, new[] { new ChartSeries("S", new double[] { 1 }) });

        // "$B$2" passes XLHelper.IsValidA1Address but GetColumnNumberFromAddress throws on the
        // literal "$" - measured directly. SetCell/ReadCell already accept "$"-prefixed refs (they
        // hand the string straight to ClosedXML), so AddChart should too.
        var result = WorkbookEditor.AddChart(xlsx, "Sheet1", "$B$2", ChartType.Line, data);

        using var source = new MemoryStream(result, writable: false);
        using var doc = OfficeIMO.Excel.ExcelDocument.Load(source);
        Assert.Single(doc.Sheets.First(s => s.Name == "Sheet1").Charts);
    }

    [Fact]
    public void AddChart_NullData_Throws()
    {
        var xlsx = WorkbookEditor.Create("Sheet1", new object[][] { new object[] { "A" } });

        Assert.Throws<ArgumentNullException>(
            () => WorkbookEditor.AddChart(xlsx, "Sheet1", "A1", ChartType.Line, null!));
    }

    [Fact]
    public async Task AddChartAsync_FromStream_MatchesTheByteArrayOverload()
    {
        // Not a byte-for-byte comparison, unlike this file's other *Async parity tests - measured
        // directly (two successive WorkbookEditor.AddChart calls on identical input) that
        // OfficeIMO.Excel's chart writer mints a fresh relationship id (e.g. "Rb7d502c93c1f4fc8")
        // for the hidden chart-data sheet and a fresh pair of c:axId values on every call, so
        // xl/workbook.xml, both *.rels parts and xl/drawings/charts/chart1.xml differ between any
        // two calls even with nothing else changed. That is OfficeIMO's own non-determinism, not a
        // defect in AddChartCore - same family as this repo's documented "never assert an exact
        // PDF size or a byte-for-byte comparison" lesson for OfficeIMO-rendered output. So this
        // asserts the two overloads did the same EDIT rather than produced identical bytes.
        var xlsx = WorkbookEditor.Create("Sheet1", new object[][] { new object[] { "A" }, new object[] { 1 } });
        var data = new ChartData(new[] { "A" }, new[] { new ChartSeries("S", new double[] { 1 }) });
        var expected = WorkbookEditor.AddChart(
            xlsx, "Sheet1", "C1", ChartType.Line, data, title: "Regional Totals");

        using var source = new MemoryStream(xlsx, writable: false);
        using var destination = new MemoryStream();
        await WorkbookEditor.AddChartAsync(
            source, "Sheet1", "C1", ChartType.Line, data, destination, title: "Regional Totals");
        var actual = destination.ToArray();

        Assert.Equal(
            WorkbookEditor.ReadSheet(expected, "Sheet1"),
            WorkbookEditor.ReadSheet(actual, "Sheet1"));

        using var expectedDoc = OfficeIMO.Excel.ExcelDocument.Load(new MemoryStream(expected, writable: false));
        using var actualDoc = OfficeIMO.Excel.ExcelDocument.Load(new MemoryStream(actual, writable: false));
        var expectedChart = Assert.Single(expectedDoc.Sheets.First(s => s.Name == "Sheet1").Charts);
        var actualChart = Assert.Single(actualDoc.Sheets.First(s => s.Name == "Sheet1").Charts);
        Assert.Equal(expectedChart.Title, actualChart.Title);
        Assert.Equal(expectedChart.ChartType, actualChart.ChartType);

        // ExcelChart.DataRange is populated only for a chart built in-memory - measured directly
        // (add a chart, read DataRange straight back: real SeriesCount/CategoryCount; save, then
        // Load the identical bytes and read the SAME chart's DataRange again: null on both sides
        // here). So series/category counts cannot be pinned through the OfficeIMO model once the
        // chart has round-tripped through Save/Load, and this reads the raw chart XML instead -
        // the one place per-series and per-category content actually survives that round trip.
        //
        // Masking exactly the two token families the comment above names as OfficeIMO's own
        // non-determinism (a relationship id, and the axId/crossAx pair the axes cross-reference)
        // leaves everything else in chart1.xml - the chart-type element, every series' formula and
        // cached values, the category range - to compare byte-for-byte. Verified this discriminates
        // rather than passing vacuously: swapping ChartType (line -> bar) or adding a second series
        // changes the masked text; two calls on identical input do not.
        var expectedChartXml = MaskNonDeterministicChartIds(
            ReadZipEntryText(expected, "xl/drawings/charts/chart1.xml"));
        var actualChartXml = MaskNonDeterministicChartIds(
            ReadZipEntryText(actual, "xl/drawings/charts/chart1.xml"));
        Assert.Equal(expectedChartXml, actualChartXml);
    }

    [Fact]
    public async Task AddChartAsync_FromFileToFile_MatchesTheByteArrayOverload()
    {
        var xlsx = WorkbookEditor.Create("Sheet1", new object[][] { new object[] { "A" }, new object[] { 1 } });
        var data = new ChartData(new[] { "A" }, new[] { new ChartSeries("S", new double[] { 1 }) });

        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, xlsx);

        await WorkbookEditor.AddChartAsync(
            input.Path, output.Path, "Sheet1", "C1", ChartType.Line, data, title: "Regional Totals");

        using var expectedDoc = OfficeIMO.Excel.ExcelDocument.Load(
            new MemoryStream(
                WorkbookEditor.AddChart(xlsx, "Sheet1", "C1", ChartType.Line, data, title: "Regional Totals"),
                writable: false));
        using var actualDoc = OfficeIMO.Excel.ExcelDocument.Load(
            new MemoryStream(await File.ReadAllBytesAsync(output.Path), writable: false));
        var expectedChart = Assert.Single(expectedDoc.Sheets.First(s => s.Name == "Sheet1").Charts);
        var actualChart = Assert.Single(actualDoc.Sheets.First(s => s.Name == "Sheet1").Charts);
        Assert.Equal(expectedChart.Title, actualChart.Title);
        Assert.Equal(expectedChart.ChartType, actualChart.ChartType);
    }

    public static IEnumerable<object[]> AllChartTypes() =>
        Enum.GetValues<ChartType>().Select(t => new object[] { t });

    [Theory]
    [MemberData(nameof(AllChartTypes))]
    public void AddChart_EveryChartTypeProducesALoadableChart(ChartType type)
    {
        // Only ColumnClustered/Line were ever probed against this ticket's categories+values
        // ChartData model before this theory. Scatter and Bubble are genuinely different-shaped
        // chart types (X/Y pairs) fed the same model here - this proves all 17 values, including
        // those two, actually succeed rather than throw or silently produce zero charts.
        var xlsx = WorkbookEditor.Create(
            "Sheet1", new object[][] { new object[] { "North" }, new object[] { "South" } });
        var data = new ChartData(
            new[] { "North", "South" },
            new[] { new ChartSeries("Total", new double[] { 1200, 980 }) });

        var result = WorkbookEditor.AddChart(xlsx, "Sheet1", "C1", type, data);

        using var source = new MemoryStream(result, writable: false);
        using var doc = OfficeIMO.Excel.ExcelDocument.Load(source);
        var chart = Assert.Single(doc.Sheets.First(s => s.Name == "Sheet1").Charts);
        Assert.NotNull(chart);
    }

    [Fact]
    public void AddPivotTable_CreatesAPivotTable_ThatSurvivesTheXlsxRoundTrip()
    {
        var xlsx = WorkbookEditor.Create(
            "Data",
            new object[][]
            {
                new object[] { "Region", "Amount" },
                new object[] { "North", 100 },
                new object[] { "South", 200 },
            });

        var result = WorkbookEditor.AddPivotTable(
            xlsx, "Data", "A1:B3", "D1", "MyPivot",
            new[] { "Region" },
            new[] { new PivotDataField("Amount", PivotFunction.Sum) });

        using var source = new MemoryStream(result, writable: false);
        using var doc = OfficeIMO.Excel.ExcelDocument.Load(source);
        Assert.Single(doc.GetPivotTables());
    }

    [Fact]
    public void AddPivotTable_TheGridIsEmptyUntilExcelRecalculates()
    {
        // Load-bearing: pins this design's central finding. A pivot table's result grid is
        // computed only by Excel, on open - nothing that WRITES the file (OfficeIMO here)
        // populates it. If a future OfficeIMO version starts writing computed values, or a
        // regression populates garbage, this test is what catches either.
        var xlsx = WorkbookEditor.Create(
            "Data",
            new object[][]
            {
                new object[] { "Region", "Amount" },
                new object[] { "North", 100 },
                new object[] { "South", 200 },
            });

        var result = WorkbookEditor.AddPivotTable(
            xlsx, "Data", "A1:B3", "D1", "MyPivot",
            new[] { "Region" },
            new[] { new PivotDataField("Amount", PivotFunction.Sum) });

        // Establish the premise first: a pivot table genuinely exists here, so the empty reads
        // below are the interesting finding rather than evidence AddPivotTable did nothing.
        using (var source = new MemoryStream(result, writable: false))
        using (var doc = OfficeIMO.Excel.ExcelDocument.Load(source))
            Assert.Single(doc.GetPivotTables());

        foreach (var cell in new[] { "D1", "D2", "D3", "E1", "E2", "E3" })
            Assert.Equal(string.Empty, WorkbookEditor.ReadCell(result, "Data", cell));
    }

    [Fact]
    public void AddPivotTable_SurvivesAnUnrelatedClosedXmlEdit()
    {
        // ClosedXML re-serializes the pivot table's XML on ANY subsequent edit through this
        // file's other (ClosedXML-based) methods - measured directly. The structure survives;
        // the bytes do not. This pins that the field configuration specifically survives.
        var xlsx = WorkbookEditor.Create(
            "Data",
            new object[][]
            {
                new object[] { "Region", "Amount" },
                new object[] { "North", 100 },
                new object[] { "South", 200 },
            });

        var withPivot = WorkbookEditor.AddPivotTable(
            xlsx, "Data", "A1:B3", "D1", "MyPivot",
            new[] { "Region" },
            new[] { new PivotDataField("Amount", PivotFunction.Sum) });

        var afterEdit = WorkbookEditor.SetCell(withPivot, "Data", "A10", "unrelated edit");

        using var source = new MemoryStream(afterEdit, writable: false);
        using var doc = OfficeIMO.Excel.ExcelDocument.Load(source);
        Assert.Single(doc.GetPivotTables());

        var pivotXml = ReadZipEntryText(afterEdit, "xl/pivotTables/pivotTable.xml");
        Assert.Contains("location ref=\"D1:E2\"", pivotXml);
        Assert.Contains("<field x=\"0\" />", pivotXml);
        Assert.Contains("dataField name=\"Sum of Amount\" fld=\"1\"", pivotXml);
    }

    [Fact]
    public void AddPivotTable_WithColumnFieldsAndMultipleDataFields_Succeeds()
    {
        var xlsx = WorkbookEditor.Create(
            "Data",
            new object[][]
            {
                new object[] { "Region", "Product", "Amount", "Qty" },
                new object[] { "North", "Widget", 100, 2 },
                new object[] { "South", "Gadget", 200, 4 },
            });

        var withPivot = WorkbookEditor.AddPivotTable(
            xlsx, "Data", "A1:D3", "F1", "RichPivot",
            new[] { "Region" },
            new[]
            {
                new PivotDataField("Amount", PivotFunction.Sum),
                new PivotDataField("Qty", PivotFunction.Average),
            },
            columnFields: new[] { "Product" });

        var afterEdit = WorkbookEditor.SetCell(withPivot, "Data", "A10", "x");

        using var source = new MemoryStream(afterEdit, writable: false);
        using var doc = OfficeIMO.Excel.ExcelDocument.Load(source);
        Assert.Single(doc.GetPivotTables());

        var pivotXml = ReadZipEntryText(afterEdit, "xl/pivotTables/pivotTable.xml");
        // Sum is OOXML's own default and its explicit subtotal="sum" attribute is dropped by
        // ClosedXML's re-serialization (measured) - only the non-default aggregation's
        // attribute survives explicitly, which is what proves the choice really round-tripped
        // rather than every field silently defaulting to Sum.
        Assert.Contains("dataField name=\"Average of Qty\"", pivotXml);
        Assert.Contains("subtotal=\"average\"", pivotXml);
        // Without this, the test's own name promises columnFields coverage that nothing above
        // actually checks - a dropped or misrouted columnFields argument would still pass every
        // other assertion here. "Product" is column index 1 (0-based), confirmed by probe.
        Assert.Contains("<colFields count=\"1\">", pivotXml);
        Assert.Contains("<field x=\"1\" />", pivotXml);
    }

    [Fact]
    public void AddPivotTable_UnknownSheetName_Throws()
    {
        var xlsx = WorkbookEditor.Create("Data", new object[][] { new object[] { "A" } });

        var ex = Assert.Throws<DocumentConversionException>(() => WorkbookEditor.AddPivotTable(
            xlsx, "NoSuchSheet", "A1:A1", "D1", "P",
            new[] { "A" }, new[] { new PivotDataField("A", PivotFunction.Sum) }));
        Assert.Contains("was not found", ex.Message);
    }

    [Fact]
    public void AddPivotTable_InvalidDestinationCell_Throws()
    {
        var xlsx = WorkbookEditor.Create("Data", new object[][] { new object[] { "A" } });

        var ex = Assert.Throws<DocumentConversionException>(() => WorkbookEditor.AddPivotTable(
            xlsx, "Data", "A1:A1", "not a cell", "P",
            new[] { "A" }, new[] { new PivotDataField("A", PivotFunction.Sum) }));
        Assert.Contains("is not a valid A1-style cell reference", ex.Message);
    }

    [Fact]
    public void AddPivotTable_EmptyRowFields_Throws()
    {
        var xlsx = WorkbookEditor.Create("Data", new object[][] { new object[] { "A" } });

        var ex = Assert.Throws<ArgumentException>(() => WorkbookEditor.AddPivotTable(
            xlsx, "Data", "A1:A1", "D1", "P",
            Array.Empty<string>(), new[] { new PivotDataField("A", PivotFunction.Sum) }));
        // Without this, a helper called with the wrong paramName (e.g. blaming dataFields for a
        // rowFields problem) would still pass, since both cases throw the same exception type.
        Assert.Equal("rowFields", ex.ParamName);
    }

    [Fact]
    public void AddPivotTable_EmptyDataFields_Throws()
    {
        var xlsx = WorkbookEditor.Create("Data", new object[][] { new object[] { "A" } });

        var ex = Assert.Throws<ArgumentException>(() => WorkbookEditor.AddPivotTable(
            xlsx, "Data", "A1:A1", "D1", "P",
            new[] { "A" }, Array.Empty<PivotDataField>()));
        Assert.Equal("dataFields", ex.ParamName);
    }

    [Fact]
    public void AddPivotTable_NullRowFields_Throws()
    {
        var xlsx = WorkbookEditor.Create("Data", new object[][] { new object[] { "A" } });

        var ex = Assert.Throws<ArgumentNullException>(() => WorkbookEditor.AddPivotTable(
            xlsx, "Data", "A1:A1", "D1", "P",
            null!, new[] { new PivotDataField("A", PivotFunction.Sum) }));
        Assert.Equal("rowFields", ex.ParamName);
    }

    [Fact]
    public void AddPivotTable_NullDataFields_Throws()
    {
        var xlsx = WorkbookEditor.Create("Data", new object[][] { new object[] { "A" } });

        var ex = Assert.Throws<ArgumentNullException>(() => WorkbookEditor.AddPivotTable(
            xlsx, "Data", "A1:A1", "D1", "P",
            new[] { "A" }, null!));
        Assert.Equal("dataFields", ex.ParamName);
    }

    [Fact]
    public void AddPivotTable_WithPageFields_Succeeds()
    {
        var xlsx = WorkbookEditor.Create(
            "Data",
            new object[][]
            {
                new object[] { "Region", "Product", "Amount" },
                new object[] { "North", "Widget", 100 },
                new object[] { "South", "Gadget", 200 },
            });

        var withPivot = WorkbookEditor.AddPivotTable(
            xlsx, "Data", "A1:C3", "E1", "P",
            new[] { "Region" },
            new[] { new PivotDataField("Amount", PivotFunction.Sum) },
            pageFields: new[] { "Product" });

        // Round-tripping through an unrelated ClosedXML edit, matching every other pivot test in
        // this file - the raw OfficeIMO output namespace-prefixes every element ("<x:pageFields>"),
        // which ClosedXML's own re-serialization strips, so reading straight off AddPivotTable's
        // output would assert against a different (and less representative, since every caller of
        // this library also feeds the result through further edits) XML shape than the rest of
        // this file checks.
        var afterEdit = WorkbookEditor.SetCell(withPivot, "Data", "A10", "unrelated edit");

        using var source = new MemoryStream(afterEdit, writable: false);
        using var doc = OfficeIMO.Excel.ExcelDocument.Load(source);
        Assert.Single(doc.GetPivotTables());

        // Without this, a dropped or misrouted pageFields argument would still pass every other
        // assertion here - matches the columnFields test's own discipline above. "Product" is
        // field index 1 (0-based), confirmed by probe.
        var pivotXml = ReadZipEntryText(afterEdit, "xl/pivotTables/pivotTable.xml");
        Assert.Contains("<pageFields count=\"1\">", pivotXml);
        Assert.Contains("<pageField fld=\"1\" />", pivotXml);
    }

    [Fact]
    public void AddPivotTable_AcceptsAbsoluteCellReferences()
    {
        // OfficeIMO's own AddPivotTable rejects a "$"-prefixed destinationCell outright
        // ("Invalid destination cell '$D$1'") - measured directly. AddPivotTableCore strips "$"
        // before passing the string through, matching AddChart's identical fix for the same
        // reason. Nothing else in this file exercised that stripping until this test - a
        // regression that passed the ORIGINAL destinationCell through instead of the cleaned one
        // would otherwise pass the whole suite silently.
        var xlsx = WorkbookEditor.Create(
            "Data",
            new object[][]
            {
                new object[] { "Region", "Amount" },
                new object[] { "North", 100 },
                new object[] { "South", 200 },
            });

        var result = WorkbookEditor.AddPivotTable(
            xlsx, "Data", "A1:B3", "$D$1", "MyPivot",
            new[] { "Region" },
            new[] { new PivotDataField("Amount", PivotFunction.Sum) });

        using var source = new MemoryStream(result, writable: false);
        using var doc = OfficeIMO.Excel.ExcelDocument.Load(source);
        Assert.Single(doc.GetPivotTables());
    }

    [Fact]
    public async Task AddPivotTableAsync_FromStream_MatchesTheByteArrayOverload()
    {
        var xlsx = WorkbookEditor.Create(
            "Data",
            new object[][]
            {
                new object[] { "Region", "Amount" },
                new object[] { "North", 100 },
                new object[] { "South", 200 },
            });
        var rowFields = new[] { "Region" };
        var dataFields = new[] { new PivotDataField("Amount", PivotFunction.Sum) };

        var expected = WorkbookEditor.AddPivotTable(xlsx, "Data", "A1:B3", "D1", "P", rowFields, dataFields);

        using var source = new MemoryStream(xlsx, writable: false);
        using var destination = new MemoryStream();
        await WorkbookEditor.AddPivotTableAsync(
            source, "Data", "A1:B3", "D1", "P", rowFields, dataFields, destination);
        var actual = destination.ToArray();

        // NOT Assert.Equal(expected, actual) on the whole workbook - confirmed non-deterministic
        // by probe (two independent AddPivotTable calls on identical input produced 8510 vs 8509
        // total bytes). But xl/pivotTables/pivotTable.xml itself IS byte-identical between two
        // independent calls (measured directly - unlike A95's chart XML, which needed masking for
        // a relationship id and axId/crossAx pair), so this compares that part directly rather
        // than only checking a single substring is present on each side.
        using var expectedStream = new MemoryStream(expected, writable: false);
        using var actualStream = new MemoryStream(actual, writable: false);
        using var expectedDoc = OfficeIMO.Excel.ExcelDocument.Load(expectedStream);
        using var actualDoc = OfficeIMO.Excel.ExcelDocument.Load(actualStream);
        Assert.Single(expectedDoc.GetPivotTables());
        Assert.Single(actualDoc.GetPivotTables());

        var expectedXml = ReadZipEntryText(expected, "xl/pivotTables/pivotTable.xml");
        var actualXml = ReadZipEntryText(actual, "xl/pivotTables/pivotTable.xml");
        Assert.Equal(expectedXml, actualXml);
    }

    [Fact]
    public async Task AddPivotTableAsync_FromFileToFile_MatchesTheByteArrayOverload()
    {
        var xlsx = WorkbookEditor.Create(
            "Data",
            new object[][]
            {
                new object[] { "Region", "Amount" },
                new object[] { "North", 100 },
                new object[] { "South", 200 },
            });
        var rowFields = new[] { "Region" };
        var dataFields = new[] { new PivotDataField("Amount", PivotFunction.Sum) };
        var expected = WorkbookEditor.AddPivotTable(xlsx, "Data", "A1:B3", "D1", "P", rowFields, dataFields);

        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, xlsx);

        await WorkbookEditor.AddPivotTableAsync(
            input.Path, output.Path, "Data", "A1:B3", "D1", "P", rowFields, dataFields);
        var actual = await File.ReadAllBytesAsync(output.Path);

        using var actualStream = new MemoryStream(actual, writable: false);
        using var doc = OfficeIMO.Excel.ExcelDocument.Load(actualStream);
        Assert.Single(doc.GetPivotTables());

        // See AddPivotTableAsync_FromStream_MatchesTheByteArrayOverload's comment: pivotTable.xml
        // is byte-identical between independent calls, so this compares it directly.
        Assert.Equal(
            ReadZipEntryText(expected, "xl/pivotTables/pivotTable.xml"),
            ReadZipEntryText(actual, "xl/pivotTables/pivotTable.xml"));
    }

    [Fact]
    public async Task AddPivotTableAsync_NullDataFields_Throws()
    {
        var xlsx = WorkbookEditor.Create("Data", new object[][] { new object[] { "A" } });

        using var source = new MemoryStream(xlsx, writable: false);
        using var destination = new MemoryStream();
        var ex = await Assert.ThrowsAsync<ArgumentNullException>(() => WorkbookEditor.AddPivotTableAsync(
            source, "Data", "A1:A1", "D1", "P",
            new[] { "A" }, null!, destination));
        Assert.Equal("dataFields", ex.ParamName);
    }

    [Fact]
    public async Task AddPivotTableAsync_EmptyDataFields_Throws()
    {
        var xlsx = WorkbookEditor.Create("Data", new object[][] { new object[] { "A" } });

        using var source = new MemoryStream(xlsx, writable: false);
        using var destination = new MemoryStream();
        var ex = await Assert.ThrowsAsync<ArgumentException>(() => WorkbookEditor.AddPivotTableAsync(
            source, "Data", "A1:A1", "D1", "P",
            new[] { "A" }, Array.Empty<PivotDataField>(), destination));
        Assert.Equal("dataFields", ex.ParamName);
    }

    [Fact]
    public async Task AddPivotTableAsync_WithColumnAndPageFields_Succeeds()
    {
        var xlsx = WorkbookEditor.Create(
            "Data",
            new object[][]
            {
                new object[] { "Region", "Product", "Channel", "Amount" },
                new object[] { "North", "Widget", "Retail", 100 },
                new object[] { "South", "Gadget", "Online", 200 },
            });

        using var source = new MemoryStream(xlsx, writable: false);
        using var destination = new MemoryStream();
        await WorkbookEditor.AddPivotTableAsync(
            source, "Data", "A1:D3", "F1", "P",
            new[] { "Region" },
            new[] { new PivotDataField("Amount", PivotFunction.Sum) },
            destination,
            columnFields: new[] { "Product" },
            pageFields: new[] { "Channel" });

        // Round-tripping through an unrelated ClosedXML edit - see
        // AddPivotTable_WithPageFields_Succeeds's comment: the raw output namespace-prefixes every
        // element, which only the re-serialized form (what the rest of this file checks) drops.
        var afterEdit = WorkbookEditor.SetCell(destination.ToArray(), "Data", "A10", "unrelated edit");

        using var afterEditStream = new MemoryStream(afterEdit, writable: false);
        using var doc = OfficeIMO.Excel.ExcelDocument.Load(afterEditStream);
        Assert.Single(doc.GetPivotTables());

        // Without this, a dropped or misrouted columnFields/pageFields argument on the Stream
        // overload specifically would still pass every other assertion here - the byte[] overload
        // has its own coverage of each field individually, but never together through this
        // overload. "Product" is field index 1, "Channel" is field index 2 (both 0-based).
        var pivotXml = ReadZipEntryText(afterEdit, "xl/pivotTables/pivotTable.xml");
        Assert.Contains("<colFields count=\"1\">", pivotXml);
        Assert.Contains("<field x=\"1\" />", pivotXml);
        Assert.Contains("<pageFields count=\"1\">", pivotXml);
        Assert.Contains("<pageField fld=\"2\" />", pivotXml);
    }

    [Fact]
    public void PivotFunction_MirrorsOfficeChartDataFunctionByName()
    {
        // ToPivotDataFunction's doc comment claims this 1:1 mirror "so nothing here can drift
        // from what the writer beneath supports" - that claim needs something re-checking it on
        // every run, or a future OfficeIMO version adding a twelfth ExcelPivotDataFunction value
        // drifts silently. Matches ChartType's own parity discipline from A95.
        Assert.Equal(
            Enum.GetNames<OfficeIMO.Excel.ExcelPivotDataFunction>().OrderBy(n => n, StringComparer.Ordinal),
            Enum.GetNames<PivotFunction>().OrderBy(n => n, StringComparer.Ordinal));
    }

    public static IEnumerable<object[]> AllPivotFunctions() =>
        Enum.GetValues<PivotFunction>().Select(f => new object[] { f });

    [Theory]
    [MemberData(nameof(AllPivotFunctions))]
    public void AddPivotTable_EveryAggregationFunctionProducesALoadablePivotTable(PivotFunction function)
    {
        // Only Sum and Average were exercised by the other pivot tests in this file before this
        // theory - a swapped mapping arm (e.g. Variance <-> VarianceP) would have shipped with
        // nothing failing. Probed directly against the real OfficeIMO.Excel 3.2.6 API first
        // (unlike A95's Scatter/Bubble, all 11 ExcelPivotDataFunction values succeed with this
        // ticket's categories+values shape - no exclusion needed here).
        var xlsx = WorkbookEditor.Create(
            "Data",
            new object[][]
            {
                new object[] { "Region", "Amount" },
                new object[] { "North", 100 },
                new object[] { "South", 200 },
            });

        var result = WorkbookEditor.AddPivotTable(
            xlsx, "Data", "A1:B3", "D1", "P",
            new[] { "Region" }, new[] { new PivotDataField("Amount", function) });

        using var source = new MemoryStream(result, writable: false);
        using var doc = OfficeIMO.Excel.ExcelDocument.Load(source);
        Assert.Single(doc.GetPivotTables());
    }

    private static string ReadZipEntryText(byte[] xlsx, string entryName)
    {
        using var zip = new ZipArchive(new MemoryStream(xlsx, writable: false));
        var entry = zip.GetEntry(entryName)
            ?? throw new InvalidOperationException($"\"{entryName}\" was not found in the XLSX package.");
        using var reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }

    private static string MaskNonDeterministicChartIds(string chartXml)
    {
        var masked = Regex.Replace(chartXml, "r:id=\"[^\"]*\"", "r:id=\"MASKED\"");
        masked = Regex.Replace(masked, "<c:axId val=\"[0-9]+\"\\s*/>", "<c:axId val=\"MASKED\"/>");
        masked = Regex.Replace(masked, "<c:crossAx val=\"[0-9]+\"\\s*/>", "<c:crossAx val=\"MASKED\"/>");
        return masked;
    }

    /// <summary>
    /// A worksheet-level embedded package, anchored the way Excel anchors one: a
    /// &lt;drawing r:id="..."/&gt; element in the worksheet XML pointing at a
    /// <see cref="DrawingsPart"/>, plus a separate <see cref="EmbeddedPackagePart"/> carrying the
    /// payload, added directly to the <see cref="WorksheetPart"/> rather than nested under the
    /// <see cref="DrawingsPart"/>. OfficeIMO 3.2.6 has no XLSX-side embedding API at all (confirmed
    /// by reflection), so this is built directly through DocumentFormat.OpenXml.Packaging's own
    /// typed part API - the SDK-level equivalent of "authored the way the library under test
    /// authors one", since there is no higher-level typed API to prefer.
    ///
    /// Note: DrawingsPart does NOT support EmbeddedPackagePart as a direct child via this SDK's
    /// typed relationship API (attempting it is a compile error, CS0311 - only WorksheetPart,
    /// ChartPart and a few others do). The embedded package is therefore its own
    /// WorksheetPart-level relationship, independent of the DrawingsPart's own relationship - both
    /// are real, valid worksheet-part-level relationships either way.
    /// </summary>
    private static byte[] WorkbookWithEmbeddedPackage()
    {
        var xlsx = WorkbookEditor.Create("Data", new object[][]
        {
            new object[] { "A", "B" },
            new object[] { 1, 2 },
        });

        using var ms = new MemoryStream();
        ms.Write(xlsx, 0, xlsx.Length);
        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, true))
        {
            var wsPart = doc.WorkbookPart?.WorksheetParts.FirstOrDefault();
            if (wsPart == null)
                throw new InvalidOperationException("WorkbookPart or worksheets not found");

            var drawingsPart = wsPart!.AddNewPart<DrawingsPart>();
            var relId = wsPart!.GetIdOfPart(drawingsPart);

            var embedPart = wsPart!.AddEmbeddedPackagePart("application/octet-stream");
            var payload = Encoding.UTF8.GetBytes("PRETEND EMBEDDED PACKAGE PAYLOAD");
            using (var partStream = embedPart.GetStream(FileMode.Create, FileAccess.Write))
                partStream.Write(payload, 0, payload.Length);

            using (var drawingStream = drawingsPart.GetStream(FileMode.Create, FileAccess.Write))
            {
                var drawingXml = "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" " +
                                  "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>";
                var drawingBytes = Encoding.UTF8.GetBytes(drawingXml);
                drawingStream.Write(drawingBytes, 0, drawingBytes.Length);
            }

            wsPart!.Worksheet!.Append(new Drawing { Id = relId });
            wsPart!.Worksheet!.Save();
            doc.Save();
        }
        return ms.ToArray();
    }

    /// <summary>
    /// True if <paramref name="xlsx"/>'s first worksheet still carries the &lt;drawing&gt;
    /// element, its <see cref="DrawingsPart"/>, and that part's <see cref="EmbeddedPackagePart"/> -
    /// i.e. whether <see cref="WorkbookWithEmbeddedPackage"/>'s fixture is still fully intact and
    /// reachable, not merely whether some bytes happen to survive somewhere in the package.
    /// </summary>
    private static bool EmbeddedPackageIsFullyIntact(byte[] xlsx)
    {
        using var ms = new MemoryStream(xlsx, writable: false);
        using var doc = SpreadsheetDocument.Open(ms, false);
        var wsPart = doc.WorkbookPart?.WorksheetParts.FirstOrDefault();
        if (wsPart == null)
            return false;
        var worksheet = wsPart.Worksheet;
        if (worksheet == null)
            return false;
        var hasDrawingElement = worksheet.Elements<Drawing>().Any();
        var drawingsParts = wsPart.GetPartsOfType<DrawingsPart>().ToList();
        return hasDrawingElement
            && drawingsParts.Count == 1
            && wsPart.EmbeddedPackageParts.Any();
    }

    [Fact]
    public void AddChart_PreservesAnEmbeddedPackage()
    {
        var xlsx = WorkbookWithEmbeddedPackage();
        var data = new ChartData(
            new[] { "North", "South" },
            new[] { new ChartSeries("Total", new double[] { 1200, 980 }) });

        var result = WorkbookEditor.AddChart(xlsx, "Data", "D1", ChartType.ColumnClustered, data);

        Assert.True(EmbeddedPackageIsFullyIntact(result));
    }

    [Fact]
    public void AddPivotTable_PreservesAnEmbeddedPackage()
    {
        var xlsx = WorkbookWithEmbeddedPackage();

        var result = WorkbookEditor.AddPivotTable(
            xlsx, "Data", "A1:B3", "D1", "P",
            new[] { "A" }, new[] { new PivotDataField("B", PivotFunction.Sum) });

        Assert.True(EmbeddedPackageIsFullyIntact(result));
    }

    [Fact]
    public void ProtectThenUnprotect_PreservesAnEmbeddedPackage()
    {
        var xlsx = WorkbookWithEmbeddedPackage();

        var protectedXlsx = WorkbookEditor.Protect(xlsx, "pw123");
        var unprotected = WorkbookEditor.Unprotect(protectedXlsx, "pw123");

        Assert.True(EmbeddedPackageIsFullyIntact(unprotected));
    }

    /// <summary>
    /// PINS A KNOWN, MEASURED LIMITATION - see src/DocToolkit/README.md's "Known limitations"
    /// table, the row about ClosedXML-backed WorkbookEditor operations and embedded objects. This
    /// is not an oversight: SetCell (like Create, AppendRows, and every other ClosedXML-backed
    /// method) opens the package through ClosedXML's XLWorkbook, which reconstructs the package
    /// from its own object model on save and does not know about a &lt;drawing&gt; reference it
    /// did not create - so it silently drops the reference and its DrawingsPart, while the
    /// EmbeddedPackagePart itself survives as dead, unreachable weight in the file.
    ///
    /// If a future ClosedXML/OfficeIMO version fixes this, THIS TEST FAILS - which is the point:
    /// the fix has to be a deliberate change to this assertion (and the README row), not a
    /// silently-noticed behaviour change either way.
    /// </summary>
    [Fact]
    public void SetCell_DropsTheEmbeddedPackagesDrawingAnchor_KnownLimitation()
    {
        var xlsx = WorkbookWithEmbeddedPackage();
        Assert.True(EmbeddedPackageIsFullyIntact(xlsx), "premise: the fixture starts intact");

        var afterSetCell = WorkbookEditor.SetCell(xlsx, "Data", "C1", "new value");

        Assert.False(EmbeddedPackageIsFullyIntact(afterSetCell));

        // Precisely what survives and what does not - the finding is not merely "something is
        // gone", it is "the anchor is gone but the orphaned bytes remain".
        using var ms = new MemoryStream(afterSetCell, writable: false);
        using var doc = SpreadsheetDocument.Open(ms, false);
        var wsPart = doc.WorkbookPart?.WorksheetParts.FirstOrDefault();
        Assert.NotNull(wsPart);
        Assert.Empty(wsPart!.Worksheet!.Elements<Drawing>());
        Assert.Empty(wsPart.GetPartsOfType<DrawingsPart>());
        Assert.Single(wsPart.EmbeddedPackageParts);
    }

    /// <summary>
    /// The discriminating control for <see cref="SetCell_DropsTheEmbeddedPackagesDrawingAnchor_KnownLimitation"/>:
    /// proves the loss above is specific to drawing content ClosedXML did not create, not a
    /// general "any drawing is lost on SetCell" defect. A picture ClosedXML inserted itself,
    /// through its own AddPicture API, survives the identical SetCell round-trip.
    /// </summary>
    [Fact]
    public void SetCell_PreservesAPictureClosedXmlInsertedItself()
    {
        byte[] withPicture;
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Data");
            worksheet.Cell(1, 1).Value = "A";
            var png = Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
            using var imageStream = new MemoryStream(png);
            worksheet.AddPicture(imageStream, XLPictureFormat.Png, "TestPic")
                .MoveTo(worksheet.Cell(2, 1));
            using var ms = new MemoryStream();
            workbook.SaveAs(ms);
            withPicture = ms.ToArray();
        }

        var afterSetCell = WorkbookEditor.SetCell(withPicture, "Data", "C1", "x");

        using var reopened = new XLWorkbook(new MemoryStream(afterSetCell, writable: false));
        var reopenedWorksheet = reopened.Worksheet("Data");
        Assert.NotNull(reopenedWorksheet);
        Assert.Single(reopenedWorksheet.Pictures);
    }

    [Fact]
    public void InspectSignatures_ReportsAnUnsignedWorkbookCleanly()
    {
        var xlsx = WorkbookEditor.Create("Data", new object[][] { new object[] { "A" } });

        var info = WorkbookEditor.InspectSignatures(xlsx);

        Assert.False(info.HasSignatures);
        Assert.Equal(0, info.SignatureCount);
    }

    [Fact]
    public async Task InspectSignaturesAsync_MatchesTheByteArrayOverload()
    {
        var xlsx = WorkbookEditor.Create("Data", new object[][] { new object[] { "A" } });

        var expected = WorkbookEditor.InspectSignatures(xlsx);
        using var source = new MemoryStream(xlsx, writable: false);
        var actual = await WorkbookEditor.InspectSignaturesAsync(source);

        Assert.Equal(expected.HasSignatures, actual.HasSignatures);
        Assert.Equal(expected.SignatureCount, actual.SignatureCount);
    }

    [Fact]
    public void ValidateSignatures_ReportsAnUnsignedWorkbookCleanly()
    {
        var xlsx = WorkbookEditor.Create("Data", new object[][] { new object[] { "A" } });

        var report = WorkbookEditor.ValidateSignatures(xlsx);

        Assert.False(report.HasSignatures);
        Assert.False(report.IsCryptographicallyValid);
        Assert.Empty(report.Signatures);
    }

    [Fact]
    public async Task ValidateSignaturesAsync_MatchesTheByteArrayOverload()
    {
        var xlsx = WorkbookEditor.Create("Data", new object[][] { new object[] { "A" } });

        var expected = WorkbookEditor.ValidateSignatures(xlsx);
        using var source = new MemoryStream(xlsx, writable: false);
        var actual = await WorkbookEditor.ValidateSignaturesAsync(source);

        Assert.Equal(expected.HasSignatures, actual.HasSignatures);
        Assert.Equal(expected.IsCryptographicallyValid, actual.IsCryptographicallyValid);
    }

    [Fact]
    public void InspectSignatures_NullXlsx_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => WorkbookEditor.InspectSignatures(null!));
    }

    [Fact]
    public void InspectSignatures_EmptyXlsx_Throws()
    {
        var ex = Assert.Throws<ArgumentException>(() => WorkbookEditor.InspectSignatures(Array.Empty<byte>()));
        Assert.Equal("xlsx", ex.ParamName);
    }

    // =========================================================================================
    // Metadata
    // =========================================================================================

    [Fact]
    public void MetadataSurvivesARoundTrip()
    {
        var xlsx = SampleWorkbook();

        var stamped = WorkbookEditor.WithMetadata(xlsx, new DocumentMetadata
        {
            Title = "Quarterly report",
            Creator = "Contoso Ltd",
            Subject = "Revenue",
            Keywords = "revenue, quarterly",
        });

        var read = WorkbookEditor.ReadMetadata(stamped);

        Assert.Equal("Quarterly report", read.Title);
        Assert.Equal("Contoso Ltd", read.Creator);
        Assert.Equal("Revenue", read.Subject);
        Assert.Equal("revenue, quarterly", read.Keywords);
    }

    [Fact]
    public void MetadataNotSetReadsBackAsNullRatherThanEmpty()
    {
        var read = WorkbookEditor.ReadMetadata(SampleWorkbook());

        Assert.Null(read.Title);
        Assert.Null(read.Creator);
        Assert.Null(read.Subject);
        Assert.Null(read.Keywords);
    }

    [Fact]
    public void WithMetadata_ANullPropertyLeavesTheExistingValueInPlace()
    {
        var withTitle = WorkbookEditor.WithMetadata(SampleWorkbook(), new DocumentMetadata { Title = "Original title" });

        var stamped = WorkbookEditor.WithMetadata(withTitle, new DocumentMetadata { Creator = "Later author" });

        var read = WorkbookEditor.ReadMetadata(stamped);
        Assert.Equal("Original title", read.Title);
        Assert.Equal("Later author", read.Creator);
    }

    [Fact]
    public void WithMetadata_AnEmptyStringClearsAnExistingValue()
    {
        var withTitle = WorkbookEditor.WithMetadata(SampleWorkbook(), new DocumentMetadata { Title = "Original title" });

        var cleared = WorkbookEditor.WithMetadata(withTitle, new DocumentMetadata { Title = "" });

        Assert.Equal("", WorkbookEditor.ReadMetadata(cleared).Title);
    }

    [Fact]
    public void WithMetadata_DoesNotDisturbTheSheetData()
    {
        var xlsx = SampleWorkbook();

        var stamped = WorkbookEditor.WithMetadata(xlsx, new DocumentMetadata { Title = "T" });

        using var workbook = Open(stamped);
        Assert.Equal("North", workbook.Worksheet("Sales").Cell("A2").GetString());
    }

    [Fact]
    public void ReadMetadata_NullXlsx_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => WorkbookEditor.ReadMetadata(null!));
    }

    [Fact]
    public void ReadMetadata_EmptyXlsx_Throws()
    {
        var ex = Assert.Throws<ArgumentException>(() => WorkbookEditor.ReadMetadata(Array.Empty<byte>()));
        Assert.Equal("xlsx", ex.ParamName);
    }

    [Fact]
    public void WithMetadata_NullXlsx_Throws()
    {
        Assert.Throws<ArgumentNullException>(
            () => WorkbookEditor.WithMetadata(null!, new DocumentMetadata()));
    }

    [Fact]
    public void WithMetadata_NullMetadata_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => WorkbookEditor.WithMetadata(SampleWorkbook(), null!));
    }

    [Fact]
    public void WithMetadata_EmptyXlsx_Throws()
    {
        var ex = Assert.Throws<ArgumentException>(
            () => WorkbookEditor.WithMetadata(Array.Empty<byte>(), new DocumentMetadata()));
        Assert.Equal("xlsx", ex.ParamName);
    }

    // =========================================================================================
    // Formula evaluation and inspection
    // =========================================================================================

    private static byte[] WorkbookWithFormula(string formula) => WorkbookEditor.Create("S", new object?[][]
    {
        new object?[] { 2, 40, XlsxFormula.From(formula) },
    });

    [Fact]
    public void InspectFormulas_ReportsASupportedFormula()
    {
        var xlsx = WorkbookWithFormula("=SUM(A1:B1)");

        var inspection = WorkbookEditor.InspectFormulas(xlsx);

        Assert.Equal(1, inspection.TotalFormulas);
        Assert.Equal(1, inspection.SupportedFormulas);
        Assert.Equal(0, inspection.UnsupportedFormulas);
        Assert.True(inspection.AllSupported);
        var cell = Assert.Single(inspection.Formulas);
        Assert.Equal("S", cell.SheetName);
        Assert.Equal("C1", cell.CellReference);
        Assert.Equal("SUM(A1:B1)", cell.Formula);
        Assert.True(cell.IsSupported);
        Assert.Null(cell.UnsupportedReason);
    }

    [Fact]
    public void InspectFormulas_ReportsAnUnsupportedFormulaWithAReasonRatherThanAWrongValue()
    {
        // The control this whole capability exists for: a caller must be able to tell "this is a
        // real 42" from "this looks like a number but nothing computed it" before trusting either.
        var xlsx = WorkbookWithFormula("=NOTAREALFUNCTION(A1:B1)");

        var inspection = WorkbookEditor.InspectFormulas(xlsx);

        Assert.Equal(0, inspection.SupportedFormulas);
        Assert.Equal(1, inspection.UnsupportedFormulas);
        Assert.False(inspection.AllSupported);
        var cell = Assert.Single(inspection.Formulas);
        Assert.False(cell.IsSupported);
        Assert.NotNull(cell.UnsupportedReason);
        Assert.Contains("NOTAREALFUNCTION", cell.UnsupportedReason, StringComparison.Ordinal);
    }

    [Fact]
    public void InspectFormulas_OnAWorkbookWithNoFormulas_ReportsZero()
    {
        var inspection = WorkbookEditor.InspectFormulas(SampleWorkbook());

        Assert.Equal(0, inspection.TotalFormulas);
        Assert.Empty(inspection.Formulas);
        Assert.True(inspection.AllSupported);
    }

    [Fact]
    public void InspectFormulas_NullXlsx_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => WorkbookEditor.InspectFormulas(null!));
    }

    [Fact]
    public void InspectFormulas_EmptyXlsx_Throws()
    {
        var ex = Assert.Throws<ArgumentException>(() => WorkbookEditor.InspectFormulas(Array.Empty<byte>()));
        Assert.Equal("xlsx", ex.ParamName);
    }

    [Fact]
    public void EvaluateFormulas_WritesTheComputedValueIntoTheFileItself()
    {
        // The distinction EvaluateFormulas exists for: ReadCell already returns "42" via
        // ClosedXML's own in-memory evaluation on every read, whatever the file contains - so
        // this asserts against the RAW XML, not ReadCell, or the test could pass even if nothing
        // was ever written back.
        var xlsx = WorkbookWithFormula("=SUM(A1:B1)");
        var before = RawSheetXml(xlsx, "sheet1.xml");
        Assert.DoesNotContain(">42<", before, StringComparison.Ordinal);

        var evaluated = WorkbookEditor.EvaluateFormulas(xlsx);

        // Namespace-prefixed elements (<x:v>, not <v>) - this is OfficeIMO's own writer, a
        // different library from the ClosedXML writer WorkbookWithFormula went through, and the
        // two are not obliged to share a convention. ">42<" matches either.
        var after = RawSheetXml(evaluated, "sheet1.xml");
        Assert.Contains(">42<", after, StringComparison.Ordinal);
        // The formula itself must survive - this replaces a missing cached value, not the formula.
        Assert.Contains("SUM(A1:B1)", after, StringComparison.Ordinal);
    }

    private static string RawSheetXml(byte[] xlsx, string entryNameContains)
    {
        using var ms = new MemoryStream(xlsx, writable: false);
        using var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read);
        var entry = zip.Entries.First(e => e.FullName.Contains(entryNameContains, StringComparison.Ordinal));
        using var stream = entry.Open();
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    [Fact]
    public void EvaluateFormulas_DoesNotDisturbOrdinaryCellValues()
    {
        var xlsx = WorkbookWithFormula("=SUM(A1:B1)");

        var evaluated = WorkbookEditor.EvaluateFormulas(xlsx);

        Assert.Equal("2", WorkbookEditor.ReadCell(evaluated, "S", "A1"));
        Assert.Equal("40", WorkbookEditor.ReadCell(evaluated, "S", "B1"));
        Assert.Equal("42", WorkbookEditor.ReadCell(evaluated, "S", "C1"));
    }

    [Fact]
    public void EvaluateFormulas_NullXlsx_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => WorkbookEditor.EvaluateFormulas(null!));
    }

    [Fact]
    public void EvaluateFormulas_EmptyXlsx_Throws()
    {
        var ex = Assert.Throws<ArgumentException>(() => WorkbookEditor.EvaluateFormulas(Array.Empty<byte>()));
        Assert.Equal("xlsx", ex.ParamName);
    }
}
