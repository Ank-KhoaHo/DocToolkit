namespace DocToolkit;

/// <summary>How a conditional format draws attention to a cell.</summary>
/// <remarks>
/// <b>Four named intents, not a colour.</b> Exposing a colour would put an open-ended surface into
/// this API, and the boundary this feature is held to is a vocabulary that can be enumerated,
/// measured and guaranteed. A colour picker cannot be — see <see cref="XlsxFormat"/>.
/// </remarks>
public enum XlsxHighlight
{
    /// <summary>Something is wrong and needs attention.</summary>
    Red = 0,

    /// <summary>Something is worth checking.</summary>
    Amber = 1,

    /// <summary>Something is as it should be.</summary>
    Green = 2,

    /// <summary>Something is unremarkable and can be skimmed past.</summary>
    Grey = 3,
}

/// <summary>Which comparison an <see cref="XlsxRule"/> makes.</summary>
public enum XlsxRuleKind
{
    /// <summary>The cell's number is greater than <see cref="XlsxRule.Value"/>.</summary>
    GreaterThan = 0,

    /// <summary>The cell's number is less than <see cref="XlsxRule.Value"/>.</summary>
    LessThan = 1,

    /// <summary>
    /// The cell's number is between <see cref="XlsxRule.Value"/> and <see cref="XlsxRule.High"/>.
    /// </summary>
    Between = 2,

    /// <summary>The cell equals <see cref="XlsxRule.Text"/>.</summary>
    EqualTo = 3,

    /// <summary>The cell contains <see cref="XlsxRule.Text"/>.</summary>
    Contains = 4,

    /// <summary>The cell is empty.</summary>
    Blank = 5,
}

/// <summary>
/// A conditional format: highlight the cells in a range that meet a condition.
/// </summary>
/// <remarks>
/// Six conditions, all measured to survive a save and reload. The library beneath offers nine;
/// <c>EqualOrGreaterThan</c>, <c>EqualOrLessThan</c> and <c>NotEquals</c> are omitted because each is
/// expressible with what is here, and every extra factory is a member to test, document and support
/// forever. They persist, so adding one later is a smaller decision than removing it.
/// </remarks>
public sealed class XlsxRule
{
    private XlsxRule(
        string range, XlsxRuleKind kind, double value, double high, string? text,
        XlsxHighlight highlight)
    {
        Range = range;
        Kind = kind;
        Value = value;
        High = high;
        Text = text;
        Highlight = highlight;
    }

    /// <summary>The cells this applies to, such as <c>B2:B99</c>.</summary>
    public string Range { get; }

    /// <summary>Which comparison is made.</summary>
    public XlsxRuleKind Kind { get; }

    /// <summary>The number compared against, for the numeric kinds.</summary>
    public double Value { get; }

    /// <summary>The upper bound, for <see cref="XlsxRuleKind.Between"/>.</summary>
    public double High { get; }

    /// <summary>
    /// The text compared against, for <see cref="XlsxRuleKind.EqualTo"/> and
    /// <see cref="XlsxRuleKind.Contains"/>. Null for every other kind.
    /// </summary>
    public string? Text { get; }

    /// <summary>How a matching cell is drawn.</summary>
    public XlsxHighlight Highlight { get; }

    /// <summary>Highlights cells whose number is greater than <paramref name="value"/>.</summary>
    /// <param name="range">The cells this applies to, such as <c>B2:B99</c>.</param>
    /// <param name="value">The number to compare against.</param>
    /// <param name="highlight">How a matching cell is drawn.</param>
    /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="range"/> is blank.</exception>
    public static XlsxRule GreaterThan(string range, double value, XlsxHighlight highlight)
        => new(Require(range), XlsxRuleKind.GreaterThan, value, 0, null, highlight);

    /// <summary>Highlights cells whose number is less than <paramref name="value"/>.</summary>
    /// <param name="range">The cells this applies to, such as <c>B2:B99</c>.</param>
    /// <param name="value">The number to compare against.</param>
    /// <param name="highlight">How a matching cell is drawn.</param>
    /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="range"/> is blank.</exception>
    public static XlsxRule LessThan(string range, double value, XlsxHighlight highlight)
        => new(Require(range), XlsxRuleKind.LessThan, value, 0, null, highlight);

    /// <summary>Highlights cells whose number falls between two bounds, inclusive.</summary>
    /// <remarks>
    /// <paramref name="high"/> must not be below <paramref name="low"/>: inverted bounds describe an
    /// empty range, so the rule would never fire and the sheet would look formatted while nothing was
    /// highlighted.
    /// </remarks>
    /// <param name="range">The cells this applies to, such as <c>B2:B99</c>.</param>
    /// <param name="low">The lower bound.</param>
    /// <param name="high">The upper bound.</param>
    /// <param name="highlight">How a matching cell is drawn.</param>
    /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="range"/> is blank.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="high"/> is below <paramref name="low"/>.
    /// </exception>
    public static XlsxRule Between(string range, double low, double high, XlsxHighlight highlight)
    {
        string checkedRange = Require(range);
        if (high < low)
        {
            throw new ArgumentOutOfRangeException(
                nameof(high), high, "The upper bound is below the lower bound, so the rule could never match.");
        }

        return new XlsxRule(checkedRange, XlsxRuleKind.Between, low, high, null, highlight);
    }

    /// <summary>Highlights cells equal to <paramref name="value"/>.</summary>
    /// <param name="range">The cells this applies to, such as <c>A2:A99</c>.</param>
    /// <param name="value">The text to compare against.</param>
    /// <param name="highlight">How a matching cell is drawn.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="range"/> or <paramref name="value"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException"><paramref name="range"/> is blank.</exception>
    public static XlsxRule EqualTo(string range, string value, XlsxHighlight highlight)
    {
        string checkedRange = Require(range);
        ArgumentNullException.ThrowIfNull(value);
        return new XlsxRule(checkedRange, XlsxRuleKind.EqualTo, 0, 0, value, highlight);
    }

    /// <summary>Highlights cells containing <paramref name="text"/>.</summary>
    /// <param name="range">The cells this applies to, such as <c>A2:A99</c>.</param>
    /// <param name="text">The text to look for.</param>
    /// <param name="highlight">How a matching cell is drawn.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="range"/> or <paramref name="text"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException"><paramref name="range"/> is blank.</exception>
    public static XlsxRule Contains(string range, string text, XlsxHighlight highlight)
    {
        string checkedRange = Require(range);
        ArgumentNullException.ThrowIfNull(text);
        return new XlsxRule(checkedRange, XlsxRuleKind.Contains, 0, 0, text, highlight);
    }

    /// <summary>Highlights empty cells.</summary>
    /// <param name="range">The cells this applies to, such as <c>A2:A99</c>.</param>
    /// <param name="highlight">How a matching cell is drawn.</param>
    /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="range"/> is blank.</exception>
    public static XlsxRule Blank(string range, XlsxHighlight highlight)
        => new(Require(range), XlsxRuleKind.Blank, 0, 0, null, highlight);

    /// <summary>
    /// Checks the range here, where the caller supplied it, so the exception names their argument.
    /// </summary>
    /// <remarks>
    /// Whether the string is a <i>valid</i> range is left to the library beneath, deliberately — a
    /// second range parser here would be a second source of truth about what a range is.
    /// </remarks>
    private static string Require(string range)
    {
        // Stryker disable once Statement : equivalent - ThrowIfNullOrWhiteSpace below throws
        // ArgumentNullException for null itself, so deleting this line changes nothing observable.
        // Kept because it states the contract at the top, where a reader looks for it.
        ArgumentNullException.ThrowIfNull(range);
        ArgumentException.ThrowIfNullOrWhiteSpace(range);
        // A sheet-qualified range is REFUSED rather than accepted, because ClosedXML discards the
        // qualifier instead of honouring it. Measured on a two-sheet workbook: a rule on
        // "Other!A2:B2" - and even on "NoSuchSheet!A2:B2" - landed on the sheet Format names, with
        // no error. A caller who writes a qualifier means it, so silently ignoring it is the worst
        // available answer. Format already takes the sheet as its own parameter.
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
}
