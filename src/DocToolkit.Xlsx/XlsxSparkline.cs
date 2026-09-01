namespace DocToolkit;

/// <summary>The shape of a sparkline.</summary>
/// <remarks>
/// Three values, because ClosedXML's own sparkline vocabulary is exactly three. This is not a
/// curated subset of something larger — it is the whole set, measured.
/// </remarks>
public enum XlsxSparklineKind
{
    /// <summary>A line through the values.</summary>
    Line = 0,

    /// <summary>One column per value.</summary>
    Column = 1,

    /// <summary>A win/loss bar: one stacked column per value.</summary>
    Stacked = 2,
}

/// <summary>
/// A sparkline: a small chart drawn inside one cell, summarising a range on the same sheet.
/// </summary>
/// <remarks>
/// Unlike <see cref="XlsxTable"/> or a conditional rule, a sparkline names <b>two</b> places — the
/// cell it is drawn in, and the range it reads. Both are on the sheet
/// <see cref="WorkbookEditor.Format(byte[], string, XlsxFormat)"/> is given, and neither may carry
/// a sheet qualifier.
/// </remarks>
public sealed class XlsxSparkline
{
    private XlsxSparkline(string cell, string sourceRange, XlsxSparklineKind kind)
    {
        Cell = cell;
        SourceRange = sourceRange;
        Kind = kind;
    }

    /// <summary>The cell the sparkline is drawn in, such as <c>D1</c>. A single cell, not a range.</summary>
    public string Cell { get; }

    /// <summary>The range the sparkline summarises, such as <c>A1:C1</c>.</summary>
    public string SourceRange { get; }

    /// <summary>The sparkline's shape.</summary>
    public XlsxSparklineKind Kind { get; }

    /// <summary>Draws a sparkline in <paramref name="cell"/> from <paramref name="sourceRange"/>.</summary>
    /// <param name="cell">The cell to draw in, such as <c>D1</c>. A single cell, not a range.</param>
    /// <param name="sourceRange">The range to summarise, such as <c>A1:C1</c>.</param>
    /// <param name="kind">The shape. Defaults to <see cref="XlsxSparklineKind.Line"/>.</param>
    /// <exception cref="ArgumentNullException">Either string argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="cell"/> or <paramref name="sourceRange"/> is blank or names a sheet.
    /// </exception>
    public static XlsxSparkline At(
        string cell, string sourceRange, XlsxSparklineKind kind = XlsxSparklineKind.Line)
    {
        RequireUnqualified(cell, nameof(cell));
        RequireUnqualified(sourceRange, nameof(sourceRange));

        return new XlsxSparkline(cell, sourceRange, kind);
    }

    private static void RequireUnqualified(string value, string paramName)
    {
        ArgumentNullException.ThrowIfNull(value, paramName);
        ArgumentException.ThrowIfNullOrWhiteSpace(value, paramName);

        if (value.Contains('!'))
        {
            throw new ArgumentException(
                $"\"{value}\" names a sheet, and the sheet qualifier is silently discarded rather "
                + "than honoured. Pass the reference alone; Format's own sheetName parameter "
                + "chooses the sheet.",
                paramName);
        }
    }
}
