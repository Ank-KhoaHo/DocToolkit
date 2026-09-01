using System;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// A106's reachable half: sparklines through ClosedXML. The other half of that ticket — slicers —
/// is <b>not</b> implemented and is not implementable from this dependency graph; see the ticket
/// for the measurement, and <c>SlicersAreNotReachable</c> below for the guard that keeps the
/// measurement honest.
/// </summary>
public class XlsxSparklineTests
{
    private static byte[] Sheet() => WorkbookEditor.Create("Data",
    [
        ["Jan", "Feb", "Mar"],
        [1, 5, 3],
        [4, 2, 6],
    ]);

    private static IXLWorksheet Read(byte[] xlsx, out XLWorkbook workbook)
    {
        workbook = new XLWorkbook(new MemoryStream(xlsx, writable: false));
        return workbook.Worksheet("Data");
    }

    [Fact]
    public void ASparklineReachesTheSavedWorkbook()
    {
        byte[] xlsx = WorkbookEditor.Format(Sheet(), "Data",
            XlsxFormat.None.WithSparkline(XlsxSparkline.At("D2", "A2:C2")));

        var sheet = Read(xlsx, out var workbook);
        using (workbook)
        {
            var group = Assert.Single(sheet.SparklineGroups);
            var sparkline = Assert.Single(group.ToList());

            // The location AND the source range, not just that a group exists: a group pointing at
            // the wrong cells is exactly what a count-only assertion cannot see.
            Assert.Equal("D2", sparkline.Location.Address.ToString());
            Assert.Contains("A2:C2", sparkline.SourceData.RangeAddress.ToString(), StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData(XlsxSparklineKind.Line, XLSparklineType.Line)]
    [InlineData(XlsxSparklineKind.Column, XLSparklineType.Column)]
    [InlineData(XlsxSparklineKind.Stacked, XLSparklineType.Stacked)]
    public void EveryKindMapsToItsClosedXMLType(XlsxSparklineKind kind, XLSparklineType expected)
    {
        byte[] xlsx = WorkbookEditor.Format(Sheet(), "Data",
            XlsxFormat.None.WithSparkline(XlsxSparkline.At("D2", "A2:C2", kind)));

        var sheet = Read(xlsx, out var workbook);
        using (workbook)
        {
            Assert.Equal(expected, sheet.SparklineGroups.Single().Type);
        }
    }

    [Fact]
    public void TheKindVocabularyIsExactlyClosedXMLs()
    {
        // The enum is not a curated subset - it is the whole set, and this fails if ClosedXML
        // grows a fourth so the gap is noticed rather than silently under-exposed.
        Assert.Equal(
            Enum.GetNames<XLSparklineType>().Length,
            Enum.GetNames<XlsxSparklineKind>().Length);
    }

    [Fact]
    public void TwoSparklinesCoexistOnOneSheet()
    {
        byte[] xlsx = WorkbookEditor.Format(Sheet(), "Data", XlsxFormat.None
            .WithSparkline(XlsxSparkline.At("D2", "A2:C2"))
            .WithSparkline(XlsxSparkline.At("D3", "A3:C3", XlsxSparklineKind.Column)));

        var sheet = Read(xlsx, out var workbook);
        using (workbook)
        {
            Assert.Equal(2, sheet.SparklineGroups.Count());
        }
    }

    [Fact]
    public void ASparklineSurvivesASecondSave()
    {
        // One that survives the first round trip and vanishes on the next passes a naive test and
        // fails a real workflow.
        byte[] xlsx = WorkbookEditor.Format(Sheet(), "Data",
            XlsxFormat.None.WithSparkline(XlsxSparkline.At("D2", "A2:C2")));

        byte[] resaved = WorkbookEditor.WithMetadata(xlsx, new DocumentMetadata { Title = "T" });

        var sheet = Read(resaved, out var workbook);
        using (workbook)
        {
            Assert.Single(sheet.SparklineGroups);
        }
    }

    [Fact]
    public void XlsxSparkline_RefusesABlankOrSheetQualifiedReference()
    {
        Assert.Equal("cell", Assert.Throws<ArgumentException>(
            () => XlsxSparkline.At("  ", "A1:C1")).ParamName);
        Assert.Equal("cell", Assert.Throws<ArgumentException>(
            () => XlsxSparkline.At("Other!D1", "A1:C1")).ParamName);
        Assert.Equal("sourceRange", Assert.Throws<ArgumentException>(
            () => XlsxSparkline.At("D1", "  ")).ParamName);
        Assert.Equal("sourceRange", Assert.Throws<ArgumentException>(
            () => XlsxSparkline.At("D1", "Other!A1:C1")).ParamName);

        Assert.Throws<ArgumentNullException>(() => XlsxSparkline.At(null!, "A1:C1"));
        Assert.Throws<ArgumentNullException>(() => XlsxSparkline.At("D1", null!));
        Assert.Throws<ArgumentNullException>(() => XlsxFormat.None.WithSparkline(null!));

        // The control: a plain reference is accepted.
        var ok = XlsxSparkline.At("D1", "A1:C1");
        Assert.Equal("D1", ok.Cell);
        Assert.Equal("A1:C1", ok.SourceRange);
        Assert.Equal(XlsxSparklineKind.Line, ok.Kind);
    }

    [Fact]
    public void ApplyingASparklineDoesNotDisturbTheCellValues()
    {
        byte[] xlsx = WorkbookEditor.Format(Sheet(), "Data",
            XlsxFormat.None.WithSparkline(XlsxSparkline.At("D2", "A2:C2")));

        Assert.Equal(WorkbookEditor.ReadSheet(Sheet(), "Data"), WorkbookEditor.ReadSheet(xlsx, "Data"));
    }

    [Fact]
    public void SlicersAreNotReachable_WhichIsWhyNoneIsExposed()
    {
        // A106 recorded a measured NO beside its measured YES, and this keeps that honest: if
        // ClosedXML ever grows a slicer entry point, this fails and the ticket gets revisited
        // rather than the absence quietly staying true because nobody re-checked.
        var slicerMembers = typeof(IXLWorksheet).GetMembers()
            .Where(m => m.Name.Contains("Slicer", StringComparison.OrdinalIgnoreCase))
            .Select(m => m.Name)
            .Distinct()
            .ToList();

        Assert.Empty(slicerMembers);
    }
}
