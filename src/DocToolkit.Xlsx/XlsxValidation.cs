namespace DocToolkit;

/// <summary>What an <see cref="XlsxValidation"/> restricts a cell to.</summary>
public enum XlsxValidationKind
{
    /// <summary>A whole number within bounds.</summary>
    WholeNumber = 0,

    /// <summary>A decimal number within bounds.</summary>
    Decimal = 1,

    /// <summary>Text whose length is within bounds.</summary>
    TextLength = 2,

    /// <summary>A date within bounds.</summary>
    Date = 3,

    /// <summary>One of a fixed list of options.</summary>
    List = 4,
}

/// <summary>
/// A data validation: what a person may type into a range of cells once the workbook is open.
/// </summary>
/// <remarks>
/// This is the half of a generated workbook that survives a human editing it. Five kinds, all
/// measured to persist through a save and reload.
/// </remarks>
public sealed class XlsxValidation
{
    private XlsxValidation(
        string range, XlsxValidationKind kind, double min, double max,
        DateTime minDate, DateTime maxDate, IReadOnlyList<string> options)
    {
        Range = range;
        Kind = kind;
        Min = min;
        Max = max;
        MinDate = minDate;
        MaxDate = maxDate;
        Options = options;
    }

    /// <summary>The cells this applies to, such as <c>B2:B99</c>.</summary>
    public string Range { get; }

    /// <summary>What the cells are restricted to.</summary>
    public XlsxValidationKind Kind { get; }

    /// <summary>The lower bound, for the numeric and text-length kinds.</summary>
    public double Min { get; }

    /// <summary>The upper bound, for the numeric and text-length kinds.</summary>
    public double Max { get; }

    /// <summary>The earliest date, for <see cref="XlsxValidationKind.Date"/>.</summary>
    public DateTime MinDate { get; }

    /// <summary>The latest date, for <see cref="XlsxValidationKind.Date"/>.</summary>
    public DateTime MaxDate { get; }

    /// <summary>
    /// The permitted options, for <see cref="XlsxValidationKind.List"/>. Empty for every other kind.
    /// </summary>
    public IReadOnlyList<string> Options { get; }

    /// <summary>Restricts the range to a whole number between two bounds, inclusive.</summary>
    /// <remarks>
    /// <c>int</c> rather than <c>long</c> because that is what the file format takes — a wider
    /// parameter would only move the truncation somewhere the caller cannot see it.
    /// </remarks>
    /// <param name="range">The cells this applies to.</param>
    /// <param name="min">The lowest permitted value.</param>
    /// <param name="max">The highest permitted value.</param>
    /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="range"/> is blank.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="max"/> is below <paramref name="min"/>.
    /// </exception>
    public static XlsxValidation WholeNumberBetween(string range, int min, int max)
        => Bounded(range, XlsxValidationKind.WholeNumber, min, max);

    /// <summary>Restricts the range to a decimal number between two bounds, inclusive.</summary>
    /// <param name="range">The cells this applies to.</param>
    /// <param name="min">The lowest permitted value.</param>
    /// <param name="max">The highest permitted value.</param>
    /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="range"/> is blank.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="max"/> is below <paramref name="min"/>.
    /// </exception>
    public static XlsxValidation DecimalBetween(string range, double min, double max)
        => Bounded(range, XlsxValidationKind.Decimal, min, max);

    /// <summary>Restricts the range to text whose length is between two bounds, inclusive.</summary>
    /// <param name="range">The cells this applies to.</param>
    /// <param name="min">The shortest permitted length.</param>
    /// <param name="max">The longest permitted length.</param>
    /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="range"/> is blank.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="max"/> is below <paramref name="min"/>.
    /// </exception>
    public static XlsxValidation TextLengthBetween(string range, int min, int max)
        => Bounded(range, XlsxValidationKind.TextLength, min, max);

    /// <summary>Restricts the range to a date between two bounds, inclusive.</summary>
    /// <param name="range">The cells this applies to.</param>
    /// <param name="min">The earliest permitted date.</param>
    /// <param name="max">The latest permitted date.</param>
    /// <exception cref="ArgumentNullException"><paramref name="range"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="range"/> is blank.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="max"/> is before <paramref name="min"/>.
    /// </exception>
    public static XlsxValidation DateBetween(string range, DateTime min, DateTime max)
    {
        string checkedRange = Require(range);
        if (max < min)
        {
            throw new ArgumentOutOfRangeException(
                nameof(max), max, "The latest date is before the earliest, so nothing could satisfy the rule.");
        }

        return new XlsxValidation(checkedRange, XlsxValidationKind.Date, 0, 0, min, max, []);
    }

    /// <summary>Restricts the range to one of a fixed list of options.</summary>
    /// <remarks>
    /// <b>An option may not contain a comma or a double quote, and may not be blank.</b> The list is
    /// written into the file as one quoted, comma-joined formula, so a comma inside an option would
    /// silently become two options and a quote would produce a malformed formula — measured. This
    /// type's premise is a vocabulary that can be enumerated and guaranteed, and accepting input it
    /// cannot faithfully encode is the wrong failure.
    /// </remarks>
    /// <param name="range">The cells this applies to.</param>
    /// <param name="options">The permitted values. At least one is required.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="range"/> or <paramref name="options"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="range"/> is blank, <paramref name="options"/> is empty, or an option is null,
    /// blank, or contains a comma or a double quote.
    /// </exception>
    public static XlsxValidation OneOf(string range, params string[] options)
    {
        string checkedRange = Require(range);
        ArgumentNullException.ThrowIfNull(options);
        if (options.Length == 0)
        {
            throw new ArgumentException(
                "At least one option is required; an empty list would leave a cell nobody can fill.",
                nameof(options));
        }

        foreach (string option in options)
        {
            if (string.IsNullOrWhiteSpace(option))
            {
                throw new ArgumentException(
                    "An option was null or blank, which would become an unnamed choice nobody can pick.",
                    nameof(options));
            }

            if (option.Contains(',') || option.Contains('"'))
            {
                throw new ArgumentException(
                    $"The option \"{option}\" contains a comma or a double quote. The list is written as one "
                    + "quoted, comma-joined formula, so either would corrupt it silently.",
                    nameof(options));
            }
        }

        return new XlsxValidation(
            checkedRange, XlsxValidationKind.List, 0, 0, default, default,
            new System.Collections.ObjectModel.ReadOnlyCollection<string>([.. options]));
    }

    private static XlsxValidation Bounded(
        string range, XlsxValidationKind kind, double min, double max)
    {
        string checkedRange = Require(range);
        if (max < min)
        {
            throw new ArgumentOutOfRangeException(
                nameof(max), max, "The upper bound is below the lower bound, so nothing could satisfy the rule.");
        }

        return new XlsxValidation(checkedRange, kind, min, max, default, default, []);
    }

    /// <summary>
    /// Checks the range here, where the caller supplied it, so the exception names their argument.
    /// Whether it is a <i>valid</i> range is left to the library beneath — see <see cref="XlsxRule"/>.
    /// </summary>
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
