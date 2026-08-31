using DocToolkit;
using Xunit;

namespace DocToolkit.Tests;

public class ChartDataTests
{
    [Fact]
    public void ChartSeries_Constructs_WithNameAndValues()
    {
        var series = new ChartSeries("Revenue", new double[] { 1, 2, 3 });

        Assert.Equal("Revenue", series.Name);
        Assert.Equal(new double[] { 1, 2, 3 }, series.Values);
    }

    [Fact]
    public void ChartSeries_NullName_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => new ChartSeries(null!, new double[] { 1 }));
    }

    [Fact]
    public void ChartSeries_BlankName_Throws()
    {
        Assert.Throws<ArgumentException>(() => new ChartSeries("   ", new double[] { 1 }));
    }

    [Fact]
    public void ChartSeries_NullValues_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => new ChartSeries("Revenue", null!));
    }

    [Fact]
    public void ChartSeries_EmptyValues_Throws()
    {
        Assert.Throws<ArgumentException>(() => new ChartSeries("Revenue", Array.Empty<double>()));
    }

    [Fact]
    public void ChartData_Constructs_WithCategoriesAndSeries()
    {
        var data = new ChartData(
            new[] { "North", "South" },
            new[] { new ChartSeries("Total", new double[] { 1200, 980 }) });

        Assert.Equal(new[] { "North", "South" }, data.Categories);
        Assert.Single(data.Series);
        Assert.Equal("Total", data.Series[0].Name);
    }

    [Fact]
    public void ChartData_NullCategories_Throws()
    {
        Assert.Throws<ArgumentNullException>(
            () => new ChartData(null!, new[] { new ChartSeries("Total", new double[] { 1 }) }));
    }

    [Fact]
    public void ChartData_NullSeries_Throws()
    {
        Assert.Throws<ArgumentNullException>(
            () => new ChartData(new[] { "North" }, null!));
    }

    [Fact]
    public void ChartData_EmptyCategories_Throws()
    {
        Assert.Throws<ArgumentException>(
            () => new ChartData(
                Array.Empty<string>(), new[] { new ChartSeries("Total", new double[] { 1 }) }));
    }

    [Fact]
    public void ChartData_EmptySeries_Throws()
    {
        Assert.Throws<ArgumentException>(
            () => new ChartData(new[] { "North" }, Array.Empty<ChartSeries>()));
    }

    [Fact]
    public void ChartData_MismatchedSeriesLength_Throws()
    {
        var ex = Assert.Throws<ArgumentException>(() => new ChartData(
            new[] { "North", "South" },
            new[] { new ChartSeries("Total", new double[] { 1200 }) }));

        Assert.Contains("Total", ex.Message);
    }
}
