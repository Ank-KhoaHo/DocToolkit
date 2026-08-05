using System.Globalization;
using System.IO.Compression;
using System.Text;
using ClosedXML.Excel;
using DocToolkit;
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

        // Both read back under the ambient test culture, so any difference is baked into the file.
        Assert.Equal(FirstRow(underInvariant, 4), FirstRow(underGerman, 4));
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

        await WorkbookEditor.CreateToFileAsync(output.Path, "Sales", new[]
        {
            new object?[] { "Region", "Total" },
            new object?[] { "North", 1200 },
        });

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
}
