using System.Linq;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class WorkbookEditorServiceTests
{
    [Fact]
    public void Create_ReadCell_SetCell_RoundTripCorrectly()
    {
        var sut = new WorkbookEditorService();

        var xlsx = sut.Create("Sales", new object?[][]
        {
            new object?[] { "Region", "Total" },
            new object?[] { "North", 1200 },
        });

        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, xlsx.Take(4).ToArray());
        Assert.Equal("Region", sut.ReadCell(xlsx, "Sales", "A1"));
        Assert.Equal("1200", sut.ReadCell(xlsx, "Sales", "B2"));

        var updated = sut.SetCell(xlsx, "Sales", "B2", 1500);
        Assert.Equal("1500", sut.ReadCell(updated, "Sales", "B2"));
    }

    [Fact]
    public void Create_RejectsABlankSheetName()
    {
        var sut = new WorkbookEditorService();

        Assert.Throws<ArgumentException>(() => sut.Create(" ", new object?[][] { }));
    }

    [Fact]
    public async Task CreateAsync_ReadCellAsync_SetCellAsync_MatchTheStaticMethods()
    {
        var sut = new WorkbookEditorService();
        var rows = new object?[][]
        {
            new object?[] { "Region", "Total" },
            new object?[] { "North", 1200 },
        };

        using var created = new MemoryStream();
        await sut.CreateAsync("Sales", rows, created);
        var xlsx = created.ToArray();

        // Parity is asserted on readable content rather than on bytes: ClosedXML stamps every
        // package it builds with fresh metadata, so two Create calls on identical input never
        // produce identical bytes - not even two calls to the same static method.
        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, xlsx.Take(4).ToArray());
        Assert.Equal("Region", WorkbookEditor.ReadCell(xlsx, "Sales", "A1"));

        using var wrapperSource = new MemoryStream(xlsx);
        using var staticSource = new MemoryStream(xlsx);

        var cell = await sut.ReadCellAsync(wrapperSource, "Sales", "B2");
        Assert.Equal(await WorkbookEditor.ReadCellAsync(staticSource, "Sales", "B2"), cell);
        Assert.Equal("1200", cell);

        // This half asserts on readable content too, for the same reason as above.
        //
        // An earlier version claimed that editing an existing package IS deterministic, unlike
        // building one from scratch, and held the wrapper to byte-exact parity here. That is
        // false: ClosedXML rewrites the whole package on save, re-stamping ZIP entry timestamps,
        // which have two-second granularity. Two SetCellAsync calls that straddle a tick differ
        // by one byte, so the assertion passed only when both landed in the same tick - a flake
        // that went unnoticed for months and then failed CI on an unrelated docs-only change.
        //
        // Measured before changing this: the difference appears at byte 5222, the exact offset
        // CI reported. The same probe found PPTX and DOCX edits and DOCX->PDF conversion all
        // byte-deterministic, which is why the byte-equality assertions in the other service
        // tests are sound and were left alone. Only ClosedXML rebuilds the package this way.
        using var setWrapperSource = new MemoryStream(xlsx);
        using var updated = new MemoryStream();
        await sut.SetCellAsync(setWrapperSource, "Sales", "B2", 1500, updated);

        using var setStaticSource = new MemoryStream(xlsx);
        using var expectedUpdated = new MemoryStream();
        await WorkbookEditor.SetCellAsync(setStaticSource, "Sales", "B2", 1500, expectedUpdated);

        var fromWrapper = updated.ToArray();
        var fromStatic = expectedUpdated.ToArray();

        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, fromWrapper.Take(4).ToArray());
        Assert.Equal(fromStatic.Length, fromWrapper.Length);
        foreach (var cellRef in new[] { "A1", "B1", "A2", "B2" })
        {
            Assert.Equal(
                WorkbookEditor.ReadCell(fromStatic, "Sales", cellRef),
                WorkbookEditor.ReadCell(fromWrapper, "Sales", cellRef));
        }

        using var verifySource = new MemoryStream(fromWrapper);
        Assert.Equal("1500", await sut.ReadCellAsync(verifySource, "Sales", "B2"));
    }

    [Fact]
    public async Task ReadCellAsync_HonorsCancellation()
    {
        var sut = new WorkbookEditorService();
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        using var source = new MemoryStream();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => sut.ReadCellAsync(source, "Sales", "A1", cts.Token));
    }

    [Fact]
    public void SheetNames_MatchesTheStaticMethod()
    {
        var sut = new WorkbookEditorService();
        var xlsx = sut.Create("Sales", new object?[][]
        {
            new object?[] { "Region", "Total" },
        });

        Assert.Equal(WorkbookEditor.SheetNames(xlsx), sut.SheetNames(xlsx));
        Assert.Equal(new[] { "Sales" }, sut.SheetNames(xlsx));
    }

    [Fact]
    public async Task SheetNamesAsync_MatchesTheStaticMethod()
    {
        var sut = new WorkbookEditorService();
        var xlsx = sut.Create("Sales", new object?[][]
        {
            new object?[] { "Region", "Total" },
        });

        using var wrapperSource = new MemoryStream(xlsx);
        using var staticSource = new MemoryStream(xlsx);

        var fromWrapper = await sut.SheetNamesAsync(wrapperSource);
        var fromStatic = await WorkbookEditor.SheetNamesAsync(staticSource);

        Assert.Equal(fromStatic, fromWrapper);
        Assert.Equal(new[] { "Sales" }, fromWrapper);
    }

    [Fact]
    public void ReadSheet_MatchesTheStaticMethod()
    {
        var sut = new WorkbookEditorService();
        var xlsx = sut.Create("Sales", new object?[][]
        {
            new object?[] { "Region", "Total" },
            new object?[] { "North", 1200 },
        });

        Assert.Equal(WorkbookEditor.ReadSheet(xlsx, "Sales"), sut.ReadSheet(xlsx, "Sales"));
        Assert.Equal(
            new[]
            {
                new[] { "Region", "Total" },
                new[] { "North", "1200" },
            },
            sut.ReadSheet(xlsx, "Sales"));
    }

    [Fact]
    public async Task ReadSheetAsync_MatchesTheStaticMethod()
    {
        var sut = new WorkbookEditorService();
        var xlsx = sut.Create("Sales", new object?[][]
        {
            new object?[] { "Region", "Total" },
            new object?[] { "North", 1200 },
        });

        using var wrapperSource = new MemoryStream(xlsx);
        using var staticSource = new MemoryStream(xlsx);

        var fromWrapper = await sut.ReadSheetAsync(wrapperSource, "Sales");
        var fromStatic = await WorkbookEditor.ReadSheetAsync(staticSource, "Sales");

        Assert.Equal(fromStatic, fromWrapper);
        Assert.Equal(
            new[]
            {
                new[] { "Region", "Total" },
                new[] { "North", "1200" },
            },
            fromWrapper);
    }

    [Fact]
    public async Task ReadSheetAsync_HonorsCancellation()
    {
        var sut = new WorkbookEditorService();
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        using var source = new MemoryStream();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => sut.ReadSheetAsync(source, "Sales", cts.Token));
    }

    [Fact]
    public void Create_WithSheets_MatchesTheStaticMethod()
    {
        var sheets = new[]
        {
            XlsxSheet.Named("Sales", new[]
            {
                new object?[] { "Region", "Total" },
                new object?[] { "EMEA", 1200 },
            }),
            XlsxSheet.Named("Summary", new[]
            {
                new object?[] { "Grand total", XlsxFormula.From("SUM(Sales!B2:B2)") },
            }),
        };
        var sut = new WorkbookEditorService();

        var expected = WorkbookEditor.Create(sheets);
        var actual = sut.Create(sheets);

        // Semantic agreement, not byte equality: a freshly built workbook carries zip entry
        // timestamps, so two Create calls a second apart legitimately differ byte-for-byte.
        Assert.Equal(WorkbookEditor.SheetNames(expected), WorkbookEditor.SheetNames(actual));
        Assert.Equal(
            WorkbookEditor.ReadCell(expected, "Summary", "B1"),
            WorkbookEditor.ReadCell(actual, "Summary", "B1"));
        Assert.Equal("1200", WorkbookEditor.ReadCell(actual, "Sales", "B2"));
    }

    [Fact]
    public void AppendRows_MatchesTheStaticMethod()
    {
        var xlsx = WorkbookEditor.Create("Log", new[] { new object?[] { "start" } });
        var rows = new[] { new object?[] { "a" }, new object?[] { "b" } };
        var sut = new WorkbookEditorService();

        var expected = WorkbookEditor.AppendRows(xlsx, "Log", rows);
        var actual = sut.AppendRows(xlsx, "Log", rows);

        // B16: the literal grid, not just parity. ReadSheet compared against itself passes on two
        // empty grids, so it could not tell an append that worked from one that produced nothing.
        Assert.Equal(
            new[] { new[] { "start" }, new[] { "a" }, new[] { "b" } },
            WorkbookEditor.ReadSheet(actual, "Log"));

        Assert.Equal(
            WorkbookEditor.ReadSheet(expected, "Log"),
            WorkbookEditor.ReadSheet(actual, "Log"));
    }

    [Fact]
    public async Task CreateAsync_WithSheets_WritesTheSameWorkbook()
    {
        var sheets = new[] { XlsxSheet.Named("S", new[] { new object?[] { "only" } }) };
        var sut = new WorkbookEditorService();

        using var destination = new MemoryStream();
        await sut.CreateAsync(sheets, destination);

        Assert.Equal("only", WorkbookEditor.ReadCell(destination.ToArray(), "S", "A1"));
    }

    [Fact]
    public async Task AppendRowsAsync_MatchesTheByteArrayOverload()
    {
        var xlsx = WorkbookEditor.Create("Log", new[] { new object?[] { "start" } });
        var rows = new[] { new object?[] { "next" } };
        var sut = new WorkbookEditorService();

        var expected = WorkbookEditor.AppendRows(xlsx, "Log", rows);

        using var source = new MemoryStream(xlsx);
        using var destination = new MemoryStream();
        await sut.AppendRowsAsync(source, "Log", rows, destination);

        // B16: literal first — see AppendRows_MatchesTheStaticMethod above.
        Assert.Equal(
            new[] { new[] { "start" }, new[] { "next" } },
            WorkbookEditor.ReadSheet(destination.ToArray(), "Log"));

        Assert.Equal(
            WorkbookEditor.ReadSheet(expected, "Log"),
            WorkbookEditor.ReadSheet(destination.ToArray(), "Log"));
    }
}
