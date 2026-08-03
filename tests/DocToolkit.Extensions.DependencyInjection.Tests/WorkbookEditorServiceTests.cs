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
        Assert.Equal(WorkbookEditor.ReadCell(xlsx, "Sales", "A1"), "Region");

        var cell = await sut.ReadCellAsync(new MemoryStream(xlsx), "Sales", "B2");
        Assert.Equal(await WorkbookEditor.ReadCellAsync(new MemoryStream(xlsx), "Sales", "B2"), cell);
        Assert.Equal("1200", cell);

        using var updated = new MemoryStream();
        await sut.SetCellAsync(new MemoryStream(xlsx), "Sales", "B2", 1500, updated);

        var expectedAfterSet = WorkbookEditor.ReadCell(
            WorkbookEditor.SetCell(xlsx, "Sales", "B2", 1500), "Sales", "B2");
        Assert.Equal(
            expectedAfterSet,
            await sut.ReadCellAsync(new MemoryStream(updated.ToArray()), "Sales", "B2"));
        Assert.Equal("1500", expectedAfterSet);
    }
}
