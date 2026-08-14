using ClosedXML.Excel;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// Sheet formatting (A24).
///
/// Every assertion reads the SAVED WORKBOOK back with ClosedXML rather than inspecting the
/// <see cref="XlsxFormat"/> that went in — a test that asserts the options object holds what it was
/// given proves the options object, not the formatting.
/// </summary>
public class XlsxFormatTests
{
    private static byte[] Workbook() => WorkbookEditor.Create("Data", new[]
    {
        new object?[] { "Region", "Total" },
        new object?[] { "North", 1234.5 },
        new object?[] { "South", 980 },
    });

    private static T Read<T>(byte[] xlsx, Func<IXLWorksheet, T> read)
    {
        using var ms = new MemoryStream(xlsx);
        using var workbook = new XLWorkbook(ms);
        return read(workbook.Worksheet("Data"));
    }

    // =====================================================================================
    // Each setting, on its own
    // =====================================================================================

    [Fact]
    public void BoldHeaderRow_BoldsTheFirstRowAndNothingElse()
    {
        var formatted = WorkbookEditor.Format(Workbook(), "Data", XlsxFormat.None.WithBoldHeaderRow());

        Assert.True(Read(formatted, s => s.Cell("A1").Style.Font.Bold));
        Assert.True(Read(formatted, s => s.Cell("B1").Style.Font.Bold));

        // The data rows must NOT be bold - "everything is bold" would satisfy the two above.
        Assert.False(Read(formatted, s => s.Cell("A2").Style.Font.Bold));
        Assert.False(Read(formatted, s => s.Cell("B3").Style.Font.Bold));
    }

    [Fact]
    public void FrozenHeaderRow_SplitsBelowRowOne()
    {
        var formatted = WorkbookEditor.Format(Workbook(), "Data", XlsxFormat.None.WithFrozenHeaderRow());

        Assert.Equal(1, Read(formatted, s => s.SheetView.SplitRow));

        // ...and is genuinely off by default, or the assertion above says nothing.
        Assert.Equal(0, Read(Workbook(), s => s.SheetView.SplitRow));
    }

    [Fact]
    public void AutoFitColumns_WidensAColumnToItsContents()
    {
        var wide = WorkbookEditor.Create("Data", new[]
        {
            new object?[] { "A rather long header that needs room" },
        });

        var before = Read(wide, s => s.Column(1).Width);
        var after = Read(
            WorkbookEditor.Format(wide, "Data", XlsxFormat.None.WithAutoFitColumns()),
            s => s.Column(1).Width);

        Assert.True(after > before, $"expected the column to widen, went from {before} to {after}");
    }

    [Fact]
    public void NumberFormat_AppliesToTheNamedColumnOnly()
    {
        var formatted = WorkbookEditor.Format(
            Workbook(), "Data", XlsxFormat.None.WithNumberFormat("B", "#,##0.00"));

        Assert.Equal("#,##0.00", Read(formatted, s => s.Column("B").Style.NumberFormat.Format));
        Assert.NotEqual("#,##0.00", Read(formatted, s => s.Column("A").Style.NumberFormat.Format));
    }

    [Fact]
    public void NumberFormat_ColumnLetterIsCaseInsensitive()
    {
        var formatted = WorkbookEditor.Format(
            Workbook(), "Data", XlsxFormat.None.WithNumberFormat("b", "#,##0.00"));

        Assert.Equal("#,##0.00", Read(formatted, s => s.Column("B").Style.NumberFormat.Format));
    }

    // =====================================================================================
    // The preset, and immutability
    // =====================================================================================

    /// <summary>
    /// <see cref="XlsxFormat.Report"/> is the answer to "make this readable", so it is pinned to
    /// the three settings it promises rather than to whatever it happens to contain.
    /// </summary>
    [Fact]
    public void Report_AppliesBoldFrozenAndAutoFit()
    {
        // A long header so auto-fit's effect is a WIDENING. Against the short values in
        // Workbook() it narrows instead - a column of "Region"/"North" fits in less than the 8.43
        // default - which is correct behaviour and made an "is wider" assertion here fail.
        var wide = WorkbookEditor.Create("Data", new[]
        {
            new object?[] { "A rather long header that needs room", "Total" },
            new object?[] { "North", 1234.5 },
        });

        var formatted = WorkbookEditor.Format(wide, "Data", XlsxFormat.Report);

        Assert.True(Read(formatted, s => s.Cell("A1").Style.Font.Bold));
        Assert.Equal(1, Read(formatted, s => s.SheetView.SplitRow));
        Assert.True(Read(formatted, s => s.Column(1).Width) > Read(wide, s => s.Column(1).Width));

        // ...and does NOT quietly set a number format, which nothing could have asked it to guess.
        Assert.Empty(XlsxFormat.Report.ColumnNumberFormats);
    }

    /// <summary>
    /// The <c>With…</c> methods return new instances. Without this, <see cref="XlsxFormat.None"/>
    /// and <see cref="XlsxFormat.Report"/> — both static, both shared — would accumulate every
    /// caller's settings for the lifetime of the process.
    /// </summary>
    [Fact]
    public void WithMethodsDoNotMutateTheInstanceTheyAreCalledOn()
    {
        var bold = XlsxFormat.None.WithBoldHeaderRow();
        var boldAndFrozen = bold.WithFrozenHeaderRow();
        var withFormat = bold.WithNumberFormat("B", "0.00");

        Assert.False(XlsxFormat.None.BoldHeaderRow);
        Assert.Empty(XlsxFormat.None.ColumnNumberFormats);

        Assert.True(bold.BoldHeaderRow);
        Assert.False(bold.FreezeHeaderRow);
        Assert.Empty(bold.ColumnNumberFormats);

        Assert.True(boldAndFrozen.FreezeHeaderRow);
        Assert.Single(withFormat.ColumnNumberFormats);
    }

    /// <summary>
    /// The number-format map cannot be mutated by casting the facade back to its runtime type.
    ///
    /// <b>This was a real defect, found reviewing the source against the repo's own rules on
    /// 2026-08-14.</b> The map was handed out as a plain <c>Dictionary</c> behind an
    /// <c>IReadOnlyDictionary</c>, so one cast and one write poisoned <see cref="XlsxFormat.None"/>
    /// or <see cref="XlsxFormat.Report"/> — which are <b>static</b> — for the rest of the process,
    /// for every caller. Measured: the cast succeeded and the injected entry was visible.
    ///
    /// <see cref="WithMethodsDoNotMutateTheInstanceTheyAreCalledOn"/> guards the other route to the
    /// same damage. Both are needed: that one covers the supported API, this one covers stepping
    /// around it.
    /// </summary>
    [Fact]
    public void TheNumberFormatMapCannotBeMutatedThroughItsRuntimeType()
    {
        Assert.IsNotType<Dictionary<string, string>>(XlsxFormat.Report.ColumnNumberFormats);
        Assert.IsNotType<Dictionary<string, string>>(XlsxFormat.None.ColumnNumberFormats);
        Assert.IsNotType<Dictionary<string, string>>(
            XlsxFormat.None.WithNumberFormat("B", "0.00").ColumnNumberFormats);

        // The shared statics are still pristine, which is the damage being prevented.
        Assert.Empty(XlsxFormat.Report.ColumnNumberFormats);
        Assert.Empty(XlsxFormat.None.ColumnNumberFormats);
    }

    // =====================================================================================
    // Formatting composes, and does not disturb content
    // =====================================================================================

    [Fact]
    public void FormattingLeavesTheValuesAlone()
    {
        var formatted = WorkbookEditor.Format(Workbook(), "Data", XlsxFormat.Report);

        Assert.Equal(
            WorkbookEditor.ReadSheet(Workbook(), "Data"),
            WorkbookEditor.ReadSheet(formatted, "Data"));

        // The literal, so the equality above cannot pass on two empty grids.
        Assert.Equal("North", WorkbookEditor.ReadCell(formatted, "Data", "A2"));
    }

    [Fact]
    public void FormattingComposesWithAppendRows()
    {
        var formatted = WorkbookEditor.Format(Workbook(), "Data", XlsxFormat.Report);
        var appended = WorkbookEditor.AppendRows(formatted, "Data",
            new[] { new object?[] { "East", 42 } });

        Assert.Equal("East", WorkbookEditor.ReadCell(appended, "Data", "A4"));
        Assert.True(Read(appended, s => s.Cell("A1").Style.Font.Bold), "formatting was lost");
    }

    // =====================================================================================
    // Guards
    // =====================================================================================

    [Fact]
    public async Task RejectsBadArguments()
    {
        Assert.Throws<ArgumentNullException>(() => WorkbookEditor.Format(null!, "Data", XlsxFormat.Report));
        Assert.Throws<ArgumentNullException>(() => WorkbookEditor.Format(Workbook(), "Data", null!));
        Assert.Throws<ArgumentException>(() => WorkbookEditor.Format(Workbook(), " ", XlsxFormat.Report));

        var missing = Assert.Throws<DocumentConversionException>(
            () => WorkbookEditor.Format(Workbook(), "Nope", XlsxFormat.Report));
        Assert.Contains("Nope", missing.Message, StringComparison.Ordinal);

        using var destination = new MemoryStream();
        using var source = new MemoryStream(Workbook(), writable: false);
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => WorkbookEditor.FormatAsync(source, "Data", null!, destination));
    }

    /// <summary>A typo fails where it was written, not later at apply time.</summary>
    [Fact]
    public void WithNumberFormat_RejectsAColumnThatIsNotLetters()
    {
        Assert.Throws<ArgumentNullException>(() => XlsxFormat.None.WithNumberFormat(null!, "0.00"));
        Assert.Throws<ArgumentNullException>(() => XlsxFormat.None.WithNumberFormat("B", null!));

        foreach (var bad in new[] { "", "1", "B2", "!" })
            Assert.Throws<ArgumentException>(() => XlsxFormat.None.WithNumberFormat(bad, "0.00"));

        Assert.Throws<ArgumentException>(() => XlsxFormat.None.WithNumberFormat("B", " "));
    }

    [Fact]
    public async Task StreamOverloadMatchesTheByteArrayForm()
    {
        var xlsx = Workbook();
        var expected = WorkbookEditor.Format(xlsx, "Data", XlsxFormat.Report);

        using var source = new MemoryStream(xlsx, writable: false);
        using var destination = new MemoryStream();
        await WorkbookEditor.FormatAsync(source, "Data", XlsxFormat.Report, destination);

        var streamed = destination.ToArray();
        Assert.True(Read(streamed, s => s.Cell("A1").Style.Font.Bold));
        Assert.Equal(1, Read(streamed, s => s.SheetView.SplitRow));
        Assert.Equal(
            WorkbookEditor.ReadSheet(expected, "Data"),
            WorkbookEditor.ReadSheet(streamed, "Data"));
        Assert.True(source.CanRead, "FormatAsync disposed a source it does not own");
    }
}
