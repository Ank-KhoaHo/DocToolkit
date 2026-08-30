namespace DocToolkit;

/// <summary>One named series of numeric values in a <see cref="ChartData"/>.</summary>
public sealed class ChartSeries
{
    /// <param name="name">The series' label, shown in the chart's legend.</param>
    /// <param name="values">
    /// The series' values, in the same order as <see cref="ChartData.Categories"/>.
    /// </param>
    /// <exception cref="ArgumentNullException"><paramref name="name"/> or <paramref name="values"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="name"/> is blank, or <paramref name="values"/> is empty.</exception>
    public ChartSeries(string name, IEnumerable<double> values)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ArgumentNullException.ThrowIfNull(values);
        var list = values.ToList();
        if (list.Count == 0) throw new ArgumentException("Series values were empty.", nameof(values));

        Name = name;
        Values = list;
    }

    /// <summary>The series' label, shown in the chart's legend.</summary>
    public string Name { get; }

    /// <summary>The series' values, in the same order as <see cref="ChartData.Categories"/>.</summary>
    public IReadOnlyList<double> Values { get; }
}
