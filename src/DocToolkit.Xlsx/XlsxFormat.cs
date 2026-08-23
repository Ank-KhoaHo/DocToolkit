using System.Collections.ObjectModel;

namespace DocToolkit;

/// <summary>
/// The formatting <see cref="WorkbookEditor.Format(byte[], string, XlsxFormat)"/> applies to a
/// sheet: a bold header row, a frozen header row, auto-fitted columns, and a number format per
/// column.
/// </summary>
/// <remarks>
/// <b>This is deliberately a small, closed set, and the smallness is the design.</b> Cell
/// formatting is an open-ended surface — fonts, borders, fills, conditional rules, merged
/// ranges — and this package's premise is a narrow one it can actually guarantee. The four
/// settings here are the ones that turn a generated grid into a report somebody can read; the
/// underlying library supports far more, and adding it here means owning it forever.
///
/// If you need more than this, the honest answer is to use ClosedXML directly rather than for
/// DocToolkit to grow a second, worse styling API in front of it.
///
/// Immutable, with <c>With…</c> methods returning a new instance — the same shape as
/// <see cref="PageSetup"/>.
/// </remarks>
public sealed class XlsxFormat
{
    private XlsxFormat(
        bool boldHeaderRow,
        bool freezeHeaderRow,
        bool autoFitColumns,
        IReadOnlyDictionary<string, string> columnNumberFormats)
    {
        BoldHeaderRow = boldHeaderRow;
        FreezeHeaderRow = freezeHeaderRow;
        AutoFitColumns = autoFitColumns;

        // Wrapped, not just typed as IReadOnlyDictionary. A plain Dictionary handed out behind
        // that interface casts straight back to Dictionary, and this type's two instances are
        // STATIC - so one cast and one write would poison XlsxFormat.None or .Report for the
        // lifetime of the process, for every caller. Measured 2026-08-14 before this wrapper
        // existed: the cast succeeded and the injected entry was visible through Report.
        //
        // Same reasoning as DocToolkitOptions.RemoteImage being get-only: an immutability claim
        // that a caller can step around is not a claim, and the cost of holding it is one
        // allocation on a path nobody calls in a loop.
        // Copied as well as wrapped. The copy is what guarantees the case-insensitive comparer
        // survives every With... call, and it costs one allocation on a path nobody calls in a
        // loop - these maps hold a handful of column letters.
        ColumnNumberFormats = new ReadOnlyDictionary<string, string>(
            new Dictionary<string, string>(columnNumberFormats, StringComparer.OrdinalIgnoreCase));
    }

    /// <summary>Applies nothing. The starting point for building a format up.</summary>
    public static XlsxFormat None { get; } =
        new(false, false, false, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));

    /// <summary>
    /// The three settings that make a generated sheet readable: a bold header row, that row frozen
    /// so it stays visible while scrolling, and columns wide enough to show their contents.
    /// </summary>
    /// <remarks>
    /// A preset rather than three calls because it is the answer to the question people actually
    /// have — "make this look like a report" — and because leaving it out would mean every caller
    /// rediscovering the same three settings.
    /// </remarks>
    public static XlsxFormat Report { get; } =
        new(true, true, true, new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase));

    /// <summary>Whether the first row is bold.</summary>
    public bool BoldHeaderRow { get; }

    /// <summary>
    /// Whether the first row stays visible while the rest of the sheet scrolls.
    /// </summary>
    public bool FreezeHeaderRow { get; }

    /// <summary>Whether each column is widened to fit its contents.</summary>
    public bool AutoFitColumns { get; }

    /// <summary>
    /// Number formats by column letter — <c>"B"</c> to <c>"#,##0.00"</c>, for instance. Empty
    /// unless set.
    /// </summary>
    /// <remarks>
    /// Keyed by <b>column letter</b> rather than index, because that is how a spreadsheet's user
    /// refers to a column and how <see cref="WorkbookEditor.ReadCell(byte[], string, string)"/>
    /// already addresses cells. Case-insensitive: <c>"b"</c> and <c>"B"</c> are the same column.
    ///
    /// The strings are Excel's own number-format codes, passed through unaltered. DocToolkit does
    /// not validate or translate them — inventing a format language in front of a standard one
    /// would be a second thing to learn and a second thing to get wrong.
    /// </remarks>
    public IReadOnlyDictionary<string, string> ColumnNumberFormats { get; }

    /// <summary>Returns a copy with <see cref="BoldHeaderRow"/> set.</summary>
    public XlsxFormat WithBoldHeaderRow(bool bold = true)
        => new(bold, FreezeHeaderRow, AutoFitColumns, ColumnNumberFormats);

    /// <summary>Returns a copy with <see cref="FreezeHeaderRow"/> set.</summary>
    public XlsxFormat WithFrozenHeaderRow(bool frozen = true)
        => new(BoldHeaderRow, frozen, AutoFitColumns, ColumnNumberFormats);

    /// <summary>Returns a copy with <see cref="AutoFitColumns"/> set.</summary>
    public XlsxFormat WithAutoFitColumns(bool autoFit = true)
        => new(BoldHeaderRow, FreezeHeaderRow, autoFit, ColumnNumberFormats);

    /// <summary>
    /// Returns a copy that formats <paramref name="column"/> with <paramref name="numberFormat"/>.
    /// </summary>
    /// <param name="column">A column letter, such as <c>"B"</c>. Case-insensitive.</param>
    /// <param name="numberFormat">
    /// An Excel number-format code, such as <c>"#,##0.00"</c> or <c>"yyyy-mm-dd"</c>.
    /// </param>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="column"/> is not one or more letters, or <paramref name="numberFormat"/> is
    /// blank. Checked here rather than at apply time so a typo fails where it was written.
    /// </exception>
    public XlsxFormat WithNumberFormat(string column, string numberFormat)
    {
        ArgumentNullException.ThrowIfNull(column);
        ArgumentNullException.ThrowIfNull(numberFormat);

        if (column.Length == 0 || !column.All(char.IsAsciiLetter))
        {
            throw new ArgumentException(
                $"Column must be one or more letters, such as \"B\". Got \"{column}\".",
                nameof(column));
        }

        if (string.IsNullOrWhiteSpace(numberFormat))
            throw new ArgumentException("Number format was blank.", nameof(numberFormat));

        var formats = new Dictionary<string, string>(ColumnNumberFormats, StringComparer.OrdinalIgnoreCase)
        {
            [column] = numberFormat,
        };

        return new XlsxFormat(BoldHeaderRow, FreezeHeaderRow, AutoFitColumns, formats);
    }
}
