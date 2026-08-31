namespace DocToolkit;

/// <summary>One aggregated value column in a pivot table's data area.</summary>
public sealed class PivotDataField
{
    /// <summary>Creates a data field aggregating one source column.</summary>
    /// <param name="fieldName">The source column's header text, e.g. <c>"Amount"</c>.</param>
    /// <param name="function">The aggregation applied to <paramref name="fieldName"/>'s values.</param>
    /// <exception cref="ArgumentNullException"><paramref name="fieldName"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="fieldName"/> is blank.</exception>
    public PivotDataField(string fieldName, PivotFunction function)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(fieldName);

        FieldName = fieldName;
        Function = function;
    }

    /// <summary>The source column's header text.</summary>
    public string FieldName { get; }

    /// <summary>The aggregation applied to <see cref="FieldName"/>'s values.</summary>
    public PivotFunction Function { get; }
}
