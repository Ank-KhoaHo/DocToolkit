namespace DocToolkit;

/// <summary>
/// One named sheet and its rows, for building a multi-sheet workbook from data rather than a
/// template.
///
/// <para>Content comes from data rather than a template, so there is no source file to edit. Cell
/// typing and culture rules are identical to the single-sheet
/// <see cref="WorkbookEditor.Create(string, System.Collections.Generic.IEnumerable{System.Collections.Generic.IEnumerable{object}})"/>;
/// a cell holding an <see cref="XlsxFormula"/> is written as a formula.</para>
///
/// <para>There is deliberately no separate header row. A header is styling, and this type carries
/// content — the first row is a header only in the sense that you put your headings in it.</para>
/// </summary>
public sealed class XlsxSheet
{
    private XlsxSheet(string name, IReadOnlyList<IReadOnlyList<object?>> rows)
    {
        Name = name;
        Rows = rows;
    }

    /// <summary>The sheet's name, as it appears on the tab.</summary>
    public string Name { get; }

    /// <summary>The rows, materialised at construction so a lazy sequence cannot be enumerated twice.</summary>
    public IReadOnlyList<IReadOnlyList<object?>> Rows { get; }

    /// <summary>
    /// Creates a sheet. <paramref name="rows"/> is materialised immediately, so a lazy or
    /// single-pass sequence is safe to pass.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="name"/> or <paramref name="rows"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="name"/> is empty, longer than 31 characters, or contains a character Excel
    /// forbids in a sheet name; or an element of <paramref name="rows"/> is null.
    /// </exception>
    public static XlsxSheet Named(string name, IEnumerable<IEnumerable<object?>> rows)
    {
        WorkbookEditor.ValidateSheetName(name, nameof(name));
        ArgumentNullException.ThrowIfNull(rows);

        // Validated and copied up front so a null row surfaces as the ArgumentException it is,
        // rather than as a NullReferenceException wrapped in a conversion failure later.
        var materialised = rows
            .Select((row, index) => (IReadOnlyList<object?>)(row
                    ?? throw new ArgumentException($"Row {index + 1} was null.", nameof(rows)))
                .ToList())
            .ToList();

        return new XlsxSheet(name, materialised);
    }
}
