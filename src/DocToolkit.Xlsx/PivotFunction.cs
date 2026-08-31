namespace DocToolkit;

/// <summary>
/// The aggregation a <see cref="PivotDataField"/> applies — mirrors
/// <c>OfficeIMO.Excel.ExcelPivotDataFunction</c>'s 11 values 1:1, confirmed by
/// <c>Enum.GetNames</c>, so nothing here can drift from what the writer beneath supports.
/// </summary>
public enum PivotFunction
{
    /// <summary>Sum of the field's values.</summary>
    Sum,
    /// <summary>Average of the field's values.</summary>
    Average,
    /// <summary>Count of entries, including non-numeric ones.</summary>
    Count,
    /// <summary>Count of numeric entries only.</summary>
    CountNumbers,
    /// <summary>Largest value.</summary>
    Maximum,
    /// <summary>Smallest value.</summary>
    Minimum,
    /// <summary>Product of the field's values.</summary>
    Product,
    /// <summary>Sample standard deviation.</summary>
    StandardDeviation,
    /// <summary>Population standard deviation.</summary>
    StandardDeviationP,
    /// <summary>Sample variance.</summary>
    Variance,
    /// <summary>Population variance.</summary>
    VarianceP,
}
