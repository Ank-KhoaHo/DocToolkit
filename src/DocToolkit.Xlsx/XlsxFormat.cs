using System.Collections.ObjectModel;

namespace DocToolkit;

/// <summary>Where a sheet is frozen: how many rows and columns stay visible while scrolling.</summary>
/// <remarks>
/// A <c>record struct</c> rather than a class, because a two-int position has value semantics and
/// cannot meaningfully be null once present. Absence lives on <see cref="XlsxFormat.FreezeAt"/> being
/// null instead.
/// </remarks>
/// <param name="Row">How many rows stay visible. Zero freezes no rows.</param>
/// <param name="Column">How many columns stay visible. Zero freezes no columns.</param>
public readonly record struct XlsxFreeze(int Row, int Column);

/// <summary>
/// The presentation <see cref="WorkbookEditor.Format(byte[], string, XlsxFormat)"/> applies to a
/// sheet: a bold header row, a freeze position, auto-fitted or explicit column widths, a number
/// format per column, an autofilter, conditional formats, data validations, tables, a print setup,
/// merged cells, hyperlinks and comments.
/// </summary>
/// <remarks>
/// <b>The boundary here is a CLOSED vocabulary, not a small one — and that is a change.</b> This type
/// used to say the smallness was the design, and excluded conditional rules by name. It was reversed
/// deliberately on 2026-08-26, because "small" stops being a boundary the moment anything is added,
/// while "closed" survives the question.
///
/// <list type="table">
/// <listheader><term>in</term><description>out</description></listheader>
/// <item><term>
/// a vocabulary this library can enumerate, measure and guarantee — six rule conditions, five
/// validation kinds, four highlights, four table style tiers, a freeze position, a column width, a
/// merged range, a hyperlink, a comment
/// </term><description>
/// an open one it would have to own forever — arbitrary fonts, borders, fills, colour scales, icon
/// sets
/// </description></item>
/// </list>
///
/// <see cref="XlsxHighlight"/> is the test case for that line: four named intents can be enumerated
/// and guaranteed, a colour picker cannot. <b>If what you need cannot be expressed as a closed set,
/// the original answer still stands — use ClosedXML directly rather than have this package grow a
/// second, worse styling API in front of it.</b>
///
/// Immutable, with <c>With…</c> methods returning a new instance — the same shape as
/// <see cref="PageSetup"/>.
/// </remarks>
public sealed class XlsxFormat
{
    private XlsxFormat(
        bool boldHeaderRow,
        bool autoFitColumns,
        IReadOnlyDictionary<string, string> columnNumberFormats,
        IReadOnlyDictionary<string, double> columnWidths,
        XlsxFreeze? freezeAt,
        bool autoFilter,
        IReadOnlyList<XlsxRule> rules,
        IReadOnlyList<XlsxValidation> validations,
        IReadOnlyList<XlsxTable> tables,
        XlsxPageSetup? pageSetup,
        IReadOnlyList<string> mergedRanges,
        IReadOnlyList<XlsxHyperlink> hyperlinks,
        IReadOnlyList<XlsxComment> comments)
    {
        BoldHeaderRow = boldHeaderRow;
        AutoFitColumns = autoFitColumns;
        FreezeAt = freezeAt;
        AutoFilter = autoFilter;
        PageSetup = pageSetup;

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
        //
        // EVERY collection below gets the same treatment, for the same reason. A List<T> handed
        // out as IReadOnlyList<T> casts back just as a Dictionary does.
        ColumnNumberFormats = new ReadOnlyDictionary<string, string>(
            new Dictionary<string, string>(columnNumberFormats, StringComparer.OrdinalIgnoreCase));
        ColumnWidths = new ReadOnlyDictionary<string, double>(
            new Dictionary<string, double>(columnWidths, StringComparer.OrdinalIgnoreCase));
        Rules = new ReadOnlyCollection<XlsxRule>([.. rules]);
        Validations = new ReadOnlyCollection<XlsxValidation>([.. validations]);
        Tables = new ReadOnlyCollection<XlsxTable>([.. tables]);
        MergedRanges = new ReadOnlyCollection<string>([.. mergedRanges]);
        Hyperlinks = new ReadOnlyCollection<XlsxHyperlink>([.. hyperlinks]);
        Comments = new ReadOnlyCollection<XlsxComment>([.. comments]);
    }

    /// <summary>Applies nothing. The starting point for building a format up.</summary>
    public static XlsxFormat None { get; } = new(
        false, false,
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
        new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase),
        null, false, [], [], [], null, [], [], []);

    /// <summary>
    /// The three settings that make a generated sheet readable: a bold header row, that row frozen
    /// so it stays visible while scrolling, and columns wide enough to show their contents.
    /// </summary>
    /// <remarks>
    /// A preset rather than three calls because it is the answer to the question people actually
    /// have — "make this look like a report" — and because leaving it out would mean every caller
    /// rediscovering the same three settings.
    /// </remarks>
    public static XlsxFormat Report { get; } = new(
        true, true,
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase),
        new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase),
        new XlsxFreeze(1, 0), false, [], [], [], null, [], [], []);

    /// <summary>Whether the first row is bold.</summary>
    public bool BoldHeaderRow { get; }

    /// <summary>
    /// Whether the first row stays visible while the rest of the sheet scrolls.
    /// </summary>
    /// <remarks>
    /// <b>Derived from <see cref="FreezeAt"/></b> rather than stored, so the two cannot disagree.
    /// Freezing anywhere else makes this false, which is the honest answer rather than a stale one.
    /// </remarks>
    public bool FreezeHeaderRow => FreezeAt is { Row: 1, Column: 0 };

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

    /// <summary>
    /// An explicit width per column letter. Empty unless set.
    /// </summary>
    /// <remarks>
    /// Applied <b>after</b> <see cref="AutoFitColumns"/>, so a named column takes this width while
    /// the rest stay auto-fitted. A specific instruction beats a blanket one.
    /// </remarks>
    public IReadOnlyDictionary<string, double> ColumnWidths { get; }

    /// <summary>Where the sheet is frozen, or <see langword="null"/> if it is not.</summary>
    public XlsxFreeze? FreezeAt { get; }

    /// <summary>Whether the sheet's used range carries an autofilter.</summary>
    /// <remarks>
    /// A sheet with no data has no used range, so it gets no filter rather than an error.
    /// </remarks>
    public bool AutoFilter { get; }

    /// <summary>The conditional formats to apply, in the order given. Empty unless set.</summary>
    public IReadOnlyList<XlsxRule> Rules { get; }

    /// <summary>The data validations to apply, in the order given. Empty unless set.</summary>
    /// <remarks>
    /// <b>Overlapping ranges are consolidated by the library beneath, and the LATER one wins.</b>
    /// Measured: a whole-number validation on <c>B2:B10</c> followed by a list on <c>B5:B15</c>
    /// leaves the first covering only <c>B2:B4</c>; the same range twice leaves only the second.
    /// <see cref="Rules"/> does <b>not</b> behave this way — six conditional formats over two ranges
    /// stay six — so do not reason from one to the other.
    /// </remarks>
    public IReadOnlyList<XlsxValidation> Validations { get; }

    /// <summary>The tables to create, in the order given. Empty unless set.</summary>
    public IReadOnlyList<XlsxTable> Tables { get; }

    /// <summary>The sheet's print setup, or <see langword="null"/> if none is set.</summary>
    public XlsxPageSetup? PageSetup { get; }

    /// <summary>The ranges to merge, in the order given. Empty unless set.</summary>
    public IReadOnlyList<string> MergedRanges { get; }

    /// <summary>The hyperlinks to add, in the order given. Empty unless set.</summary>
    public IReadOnlyList<XlsxHyperlink> Hyperlinks { get; }

    /// <summary>The comments to add, in the order given. Empty unless set.</summary>
    public IReadOnlyList<XlsxComment> Comments { get; }

    /// <summary>Returns a copy with <see cref="BoldHeaderRow"/> set.</summary>
    /// <param name="bold">Whether the first row is bold.</param>
    public XlsxFormat WithBoldHeaderRow(bool bold = true) => With(boldHeaderRow: bold);

    /// <summary>Returns a copy that freezes row 1, or that freezes nothing.</summary>
    /// <remarks>
    /// <b><paramref name="frozen"/> false clears whatever freeze is set</b>, not only a header-row
    /// one — the state is a single position rather than two independent switches, and a method that
    /// did something different depending on the current value would be worse than one that is blunt
    /// about it.
    /// </remarks>
    /// <param name="frozen">Whether to freeze row 1.</param>
    public XlsxFormat WithFrozenHeaderRow(bool frozen = true)
        => frozen ? With(freezeAt: new XlsxFreeze(1, 0)) : With(clearFreeze: true);

    /// <summary>Returns a copy with <see cref="AutoFitColumns"/> set.</summary>
    /// <param name="autoFit">Whether each column is widened to fit its contents.</param>
    public XlsxFormat WithAutoFitColumns(bool autoFit = true) => With(autoFitColumns: autoFit);

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
        RequireColumn(column);
        ArgumentNullException.ThrowIfNull(numberFormat);

        if (string.IsNullOrWhiteSpace(numberFormat))
            throw new ArgumentException("Number format was blank.", nameof(numberFormat));

        var formats = new Dictionary<string, string>(ColumnNumberFormats, StringComparer.OrdinalIgnoreCase)
        {
            [column] = numberFormat,
        };

        return With(columnNumberFormats: formats);
    }

    /// <summary>Returns a copy that gives <paramref name="column"/> an explicit width.</summary>
    /// <remarks>
    /// Applied after <see cref="AutoFitColumns"/>, so this wins for the column it names while the
    /// rest stay auto-fitted.
    /// </remarks>
    /// <param name="column">A column letter, such as <c>"A"</c>. Case-insensitive.</param>
    /// <param name="width">The width in Excel's character units. Must be positive.</param>
    /// <exception cref="ArgumentNullException"><paramref name="column"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="column"/> is not one or more letters.</exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="width"/> is not positive.</exception>
    public XlsxFormat WithColumnWidth(string column, double width)
    {
        RequireColumn(column);
        if (width <= 0)
            throw new ArgumentOutOfRangeException(nameof(width), width, "A column width must be positive.");

        var widths = new Dictionary<string, double>(ColumnWidths, StringComparer.OrdinalIgnoreCase)
        {
            [column] = width,
        };

        return With(columnWidths: widths);
    }

    /// <summary>Returns a copy frozen at a position.</summary>
    /// <remarks>
    /// <c>(0, 0)</c> is refused: it would freeze nothing while making <see cref="FreezeAt"/> report a
    /// value, which is two spellings of one state. Use <see cref="WithFrozenHeaderRow"/> with false
    /// to freeze nothing. Rows-only and columns-only are both legal.
    /// </remarks>
    /// <param name="row">How many rows stay visible.</param>
    /// <param name="column">How many columns stay visible.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// Either is negative, or both are zero.
    /// </exception>
    public XlsxFormat WithFreezeAt(int row, int column)
    {
        if (row < 0)
            throw new ArgumentOutOfRangeException(nameof(row), row, "A freeze position cannot be negative.");

        if (row == 0 && column == 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(row), row, "Freeze at least one row or column; use WithFrozenHeaderRow(false) to freeze nothing.");
        }

        if (column < 0)
            throw new ArgumentOutOfRangeException(nameof(column), column, "A freeze position cannot be negative.");

        return With(freezeAt: new XlsxFreeze(row, column));
    }

    /// <summary>Returns a copy that puts an autofilter on the sheet's used range.</summary>
    /// <param name="enabled">Whether to apply one.</param>
    public XlsxFormat WithAutoFilter(bool enabled = true) => With(autoFilter: enabled);

    /// <summary>Returns a copy carrying one more conditional format.</summary>
    /// <param name="rule">The rule to add.</param>
    /// <exception cref="ArgumentNullException"><paramref name="rule"/> is null.</exception>
    public XlsxFormat WithRule(XlsxRule rule)
    {
        ArgumentNullException.ThrowIfNull(rule);
        return With(rules: [.. Rules, rule]);
    }

    /// <summary>Returns a copy carrying one more data validation.</summary>
    /// <param name="validation">The validation to add.</param>
    /// <exception cref="ArgumentNullException"><paramref name="validation"/> is null.</exception>
    public XlsxFormat WithValidation(XlsxValidation validation)
    {
        ArgumentNullException.ThrowIfNull(validation);
        return With(validations: [.. Validations, validation]);
    }

    /// <summary>Returns a copy carrying one more table.</summary>
    /// <remarks>
    /// <b>Do not also call <see cref="WithAutoFilter"/> over a range that overlaps this table.</b>
    /// A ClosedXML table already carries its own autofilter, and applying the sheet-wide one on
    /// top of an overlapping table range throws <see cref="DocumentConversionException"/> at
    /// <see cref="WorkbookEditor.Format(byte[], string, XlsxFormat)"/> time — measured directly,
    /// not merely likely. Use one or the other over the same cells.
    /// </remarks>
    /// <param name="table">The table to add.</param>
    /// <exception cref="ArgumentNullException"><paramref name="table"/> is null.</exception>
    public XlsxFormat WithTable(XlsxTable table)
    {
        ArgumentNullException.ThrowIfNull(table);
        return With(tables: [.. Tables, table]);
    }

    /// <summary>Returns a copy carrying a print setup.</summary>
    /// <param name="pageSetup">The print setup to apply.</param>
    /// <exception cref="ArgumentNullException"><paramref name="pageSetup"/> is null.</exception>
    public XlsxFormat WithPageSetup(XlsxPageSetup pageSetup)
    {
        ArgumentNullException.ThrowIfNull(pageSetup);
        return With(pageSetup: pageSetup);
    }

    /// <summary>Returns a copy carrying one more merged range.</summary>
    /// <param name="range">The cells to merge, such as <c>A1:C1</c>.</param>
    /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="range"/> is blank or names a sheet.</exception>
    public XlsxFormat WithMergedCells(string range)
        => With(mergedRanges: [.. MergedRanges, RequireRange(range, nameof(range))]);

    /// <summary>Returns a copy carrying one more hyperlink.</summary>
    /// <param name="hyperlink">The hyperlink to add.</param>
    /// <exception cref="ArgumentNullException"><paramref name="hyperlink"/> is null.</exception>
    public XlsxFormat WithHyperlink(XlsxHyperlink hyperlink)
    {
        ArgumentNullException.ThrowIfNull(hyperlink);
        return With(hyperlinks: [.. Hyperlinks, hyperlink]);
    }

    /// <summary>Returns a copy carrying one more comment.</summary>
    /// <param name="comment">The comment to add.</param>
    /// <exception cref="ArgumentNullException"><paramref name="comment"/> is null.</exception>
    public XlsxFormat WithComment(XlsxComment comment)
    {
        ArgumentNullException.ThrowIfNull(comment);
        return With(comments: [.. Comments, comment]);
    }

    /// <summary>
    /// The one place a modified copy is made. Every <c>With…</c> method goes through it, so adding a
    /// field means changing one call site rather than nine.
    /// </summary>
    /// <remarks>
    /// <paramref name="clearFreeze"/> exists because <c>freezeAt: null</c> cannot otherwise be told
    /// apart from "not supplied" — the one place C#'s optional-parameter pattern breaks down for a
    /// nullable field.
    /// </remarks>
    private XlsxFormat With(
        bool? boldHeaderRow = null,
        bool? autoFitColumns = null,
        IReadOnlyDictionary<string, string>? columnNumberFormats = null,
        IReadOnlyDictionary<string, double>? columnWidths = null,
        XlsxFreeze? freezeAt = null,
        bool clearFreeze = false,
        bool? autoFilter = null,
        IReadOnlyList<XlsxRule>? rules = null,
        IReadOnlyList<XlsxValidation>? validations = null,
        IReadOnlyList<XlsxTable>? tables = null,
        XlsxPageSetup? pageSetup = null,
        IReadOnlyList<string>? mergedRanges = null,
        IReadOnlyList<XlsxHyperlink>? hyperlinks = null,
        IReadOnlyList<XlsxComment>? comments = null)
        => new(boldHeaderRow ?? BoldHeaderRow,
               autoFitColumns ?? AutoFitColumns,
               columnNumberFormats ?? ColumnNumberFormats,
               columnWidths ?? ColumnWidths,
               clearFreeze ? null : freezeAt ?? FreezeAt,
               autoFilter ?? AutoFilter,
               rules ?? Rules,
               validations ?? Validations,
               tables ?? Tables,
               pageSetup ?? PageSetup,
               mergedRanges ?? MergedRanges,
               hyperlinks ?? Hyperlinks,
               comments ?? Comments);

    /// <summary>
    /// A cheap SHAPE check, run here rather than at apply time so a typo fails where it was
    /// written. It is deliberately not a column validator: "ZZZZZZ" is letters and passes, while
    /// Excel stops at XFD. Whether a column exists is left to the library beneath, for the same
    /// reason XlsxRule leaves ranges to it - a second source of truth about what a column is would
    /// be one more thing to keep in step, and it would be the one that goes stale.
    /// </summary>
    private static void RequireColumn(string column)
    {
        ArgumentNullException.ThrowIfNull(column);

        if (column.Length == 0 || !column.All(char.IsAsciiLetter))
        {
            throw new ArgumentException(
                $"Column must be one or more letters, such as \"B\". Got \"{column}\".",
                nameof(column));
        }
    }

    /// <summary>
    /// The sheet-qualifier check <see cref="XlsxRule"/> and <see cref="XlsxTable"/> each carry their
    /// own copy of, for <see cref="WithMergedCells"/>'s own range argument.
    /// </summary>
    private static string RequireRange(string range, string paramName)
    {
        ArgumentNullException.ThrowIfNull(range, paramName);
        ArgumentException.ThrowIfNullOrWhiteSpace(range, paramName);

        if (range.Contains('!'))
        {
            throw new ArgumentException(
                $"\"{range}\" names a sheet, and the sheet qualifier is silently discarded rather "
                + "than honoured. Pass the range alone; Format's own sheetName parameter chooses "
                + "the sheet.",
                paramName);
        }

        return range;
    }
}
