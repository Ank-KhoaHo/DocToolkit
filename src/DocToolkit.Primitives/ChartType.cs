namespace DocToolkit;

/// <summary>
/// The chart shapes AddChart methods on WorkbookEditor and PresentationEditor can create
/// — one closed vocabulary shared by both, mirroring OfficeIMO's own <c>OfficeChartKind</c>
/// 1:1 so nothing here can drift from what the renderer beneath actually draws.
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
    /// <summary>Scatter chart.</summary>
    Scatter,
    /// <summary>Radar chart.</summary>
    Radar,
    /// <summary>Pie chart.</summary>
    Pie,
    /// <summary>Doughnut chart.</summary>
    Doughnut,
    /// <summary>Bubble chart.</summary>
    Bubble,
}
