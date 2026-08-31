namespace DocToolkit;

/// <summary>
/// The chart shapes AddChart methods on WorkbookEditor and PresentationEditor can create
/// — one closed vocabulary shared by both, mirroring OfficeIMO's own <c>OfficeChartKind</c>,
/// minus two values that do not fit the shared categories-and-value-series
/// <see cref="ChartData"/> shape, so nothing here can drift from what the renderer beneath
/// actually draws.
///
/// <c>OfficeChartKind</c> also has <c>Scatter</c> and <c>Bubble</c>, deliberately excluded here:
/// measured directly against both the Excel and PowerPoint chart APIs, both reject a
/// <see cref="ChartData"/> built from this ticket's model — Excel's Scatter path requires
/// numeric X values where <see cref="ChartData.Categories"/> is a string label, PowerPoint's
/// Scatter path rejects the same shape for the identical reason, Excel's shared chart API
/// refuses Bubble outright ("not supported by the shared Excel chart API"), and PowerPoint's
/// Bubble path requires a bubble size per point this model has no field for. A future chart
/// feature could add a companion X/Y-and-size data type for these two; forcing them into this
/// one would have shipped two enum values that always throw.
/// </summary>
public enum ChartType
{
    /// <summary>Clustered column chart.</summary>
    ColumnClustered,
    /// <summary>Stacked column chart.</summary>
    ColumnStacked,
    /// <summary>100% stacked column chart.</summary>
    ColumnStacked100,
    /// <summary>Clustered bar chart.</summary>
    BarClustered,
    /// <summary>Stacked bar chart.</summary>
    BarStacked,
    /// <summary>100% stacked bar chart.</summary>
    BarStacked100,
    /// <summary>Line chart.</summary>
    Line,
    /// <summary>Stacked line chart.</summary>
    LineStacked,
    /// <summary>100% stacked line chart.</summary>
    LineStacked100,
    /// <summary>Area chart.</summary>
    Area,
    /// <summary>Stacked area chart.</summary>
    AreaStacked,
    /// <summary>100% stacked area chart.</summary>
    AreaStacked100,
    /// <summary>Radar chart.</summary>
    Radar,
    /// <summary>Pie chart.</summary>
    Pie,
    /// <summary>Doughnut chart.</summary>
    Doughnut,
}
