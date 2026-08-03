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

        var cell = await sut.ReadCellAsync(new MemoryStream(xlsx), "Sales", "B2");
        Assert.Equal(await WorkbookEditor.ReadCellAsync(new MemoryStream(xlsx), "Sales", "B2"), cell);
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
        using var updated = new MemoryStream();
        await sut.SetCellAsync(new MemoryStream(xlsx), "Sales", "B2", 1500, updated);

        using var expectedUpdated = new MemoryStream();
        await WorkbookEditor.SetCellAsync(new MemoryStream(xlsx), "Sales", "B2", 1500, expectedUpdated);

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

        Assert.Equal("1500", await sut.ReadCellAsync(new MemoryStream(fromWrapper), "Sales", "B2"));
    }

    [Fact]
    public async Task ReadCellAsync_HonorsCancellation()
    {
        var sut = new WorkbookEditorService();
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => sut.ReadCellAsync(new MemoryStream(), "Sales", "A1", cts.Token));
    }
}
