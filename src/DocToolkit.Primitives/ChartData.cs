namespace DocToolkit;

/// <summary>
/// The categories and series a chart plots — one category axis shared by every series, the same
/// shape <c>OfficeIMO.Drawing.OfficeChartData</c> uses, which is what lets AddChart methods
/// on WorkbookEditor and PresentationEditor share one data model.
/// </summary>
public sealed class ChartData
{
    /// <param name="categories">The category axis labels — one per data point in every series.</param>
    /// <param name="series">One or more named value series, plotted against <paramref name="categories"/>.</param>
    /// <exception cref="ArgumentNullException"><paramref name="categories"/> or <paramref name="series"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="categories"/> is empty, <paramref name="series"/> is empty, or a series'
    /// value count does not match <paramref name="categories"/>' count.
    /// </exception>
    public ChartData(IEnumerable<string> categories, IEnumerable<ChartSeries> series)
    {
        ArgumentNullException.ThrowIfNull(categories);
        ArgumentNullException.ThrowIfNull(series);
        var categoryList = categories.ToList();
        var seriesList = series.ToList();
        if (categoryList.Count == 0)
            throw new ArgumentException("Categories were empty.", nameof(categories));
        if (seriesList.Count == 0)
            throw new ArgumentException("Series were empty.", nameof(series));
        var mismatched = seriesList.FirstOrDefault(s => s.Values.Count != categoryList.Count);
        if (mismatched is not null)
        {
            throw new ArgumentException(
                $"Series \"{mismatched.Name}\" has {mismatched.Values.Count} value(s) but there "
                + $"are {categoryList.Count} categor(y/ies).",
                nameof(series));
        }

        Categories = categoryList;
        Series = seriesList;
    }

    /// <summary>The category axis labels — one per data point in every series.</summary>
    public IReadOnlyList<string> Categories { get; }

    /// <summary>One or more named value series, plotted against <see cref="Categories"/>.</summary>
    public IReadOnlyList<ChartSeries> Series { get; }
}
