namespace DocToolkit;

/// <summary>A built-in Excel table style tier.</summary>
/// <remarks>
/// <b>Four named tiers, not the 61 named themes the library beneath offers.</b> ClosedXML's
/// <c>XLTableTheme</c> is not an enum — a class with <c>GetAllThemes()</c>/<c>FromName(string)</c>
/// and 61 public static fields (<c>None</c> plus 21 Light, 28 Medium and 11 Dark tiers), and
/// <c>FromName</c> on an unrecognised name does not throw — it returns a theme whose own
/// <c>ToString()</c> is empty, which would silently produce a malformed table for a caller's typo
/// rather than failing where the mistake was made. This type holds to the same closed-vocabulary
/// discipline as <see cref="XlsxHighlight"/>, and for the identical reason stated on
/// <see cref="XlsxRule"/>: the library beneath offers more than this needs, and every extra value
/// is a member to test, document and support forever.
/// </remarks>
public enum XlsxTableStyle
{
    /// <summary>No banding, no header emphasis — a plain named range with structured references.</summary>
    None = 0,

    /// <summary>A light banded style, the lightest tier ClosedXML offers.</summary>
    Light = 1,

    /// <summary>
    /// A medium banded style — the tier ClosedXML itself applies to a freshly created table when no
    /// theme is set, measured directly rather than assumed.
    /// </summary>
    Medium = 2,

    /// <summary>A dark banded style, the strongest tier ClosedXML offers.</summary>
    Dark = 3,
}

/// <summary>
/// An Excel table (a <c>ListObject</c>, what a spreadsheet user means by "make this a table"):
/// a named, banded range with an autofilter and structured references a formula elsewhere in the
/// workbook can use.
/// </summary>
/// <remarks>
/// <b>This does not make <see cref="WorkbookEditor.AppendRows"/> keep the table current, and that
/// is measured rather than assumed.</b> <c>AppendRows</c> writes with a raw cell value at the
/// sheet's last used row — it has no awareness of any table on the sheet, and ClosedXML does not
/// retroactively absorb an adjacent cell write into a table's range on its own. A row appended
/// after this table sits adjacent to it, not inside it, until the table is recreated over the new
/// range. This is the identical shape of caveat <see cref="WorkbookEditor.AddPivotTable"/> already
/// carries for its own result grid — a real, measured limitation stated here rather than left to be
/// discovered.
/// </remarks>
public sealed class XlsxTable
{
    private XlsxTable(string range, string name, XlsxTableStyle style)
    {
        Range = range;
        Name = name;
        Style = style;
    }

    /// <summary>The cells this table covers, header row included, such as <c>A1:C10</c>.</summary>
    public string Range { get; }

    /// <summary>
    /// The table's name — what a formula elsewhere in the workbook uses for a structured reference,
    /// such as <c>Sales[Total]</c>.
    /// </summary>
    public string Name { get; }

    /// <summary>The built-in style tier applied to the table.</summary>
    public XlsxTableStyle Style { get; }

    /// <summary>Names a range as a table.</summary>
    /// <param name="range">The cells this table covers, header row included, such as <c>A1:C10</c>.</param>
    /// <param name="name">
    /// The table's name. Must be a valid Excel name: it cannot contain a space, and cannot look
    /// like a cell reference such as <c>A1</c>.
    /// </param>
    /// <param name="style">The built-in style tier to apply. Defaults to <see cref="XlsxTableStyle.Medium"/>.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="range"/> or <paramref name="name"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="range"/> or <paramref name="name"/> is blank, or <paramref name="range"/>
    /// names a sheet.
    /// </exception>
    public static XlsxTable Named(string range, string name, XlsxTableStyle style = XlsxTableStyle.Medium)
        => new(RequireRange(range), RequireName(name), style);

    /// <summary>
    /// Checks the range here, where the caller supplied it, so the exception names their argument.
    /// Mirrors <see cref="XlsxRule"/>'s own identical check.
    /// </summary>
    private static string RequireRange(string range)
    {
        ArgumentNullException.ThrowIfNull(range);
        ArgumentException.ThrowIfNullOrWhiteSpace(range);

        if (range.Contains('!'))
        {
            throw new ArgumentException(
                $"\"{range}\" names a sheet, and the sheet qualifier is silently discarded rather "
                + "than honoured. Pass the range alone; Format's own sheetName parameter chooses "
                + "the sheet.",
                nameof(range));
        }

        return range;
    }

    private static string RequireName(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        return name;
    }
}
