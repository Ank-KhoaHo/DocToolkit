namespace DocToolkit;

/// <summary>A comment (a note) on one cell.</summary>
public sealed class XlsxComment
{
    private XlsxComment(string cell, string text)
    {
        Cell = cell;
        Text = text;
    }

    /// <summary>The cell carrying the comment, such as <c>B2</c>. A single cell, not a range.</summary>
    public string Cell { get; }

    /// <summary>The comment's text.</summary>
    public string Text { get; }

    /// <summary>Puts a comment on <paramref name="cell"/>.</summary>
    /// <param name="cell">The cell carrying the comment, such as <c>B2</c>. A single cell, not a range.</param>
    /// <param name="text">The comment's text.</param>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="cell"/> is blank or names a sheet, or <paramref name="text"/> is blank.
    /// </exception>
    public static XlsxComment On(string cell, string text)
    {
        ArgumentNullException.ThrowIfNull(cell);
        ArgumentException.ThrowIfNullOrWhiteSpace(cell);
        ArgumentNullException.ThrowIfNull(text);
        ArgumentException.ThrowIfNullOrWhiteSpace(text);

        if (cell.Contains('!'))
        {
            throw new ArgumentException(
                $"\"{cell}\" names a sheet, and the sheet qualifier is silently discarded rather "
                + "than honoured. Pass the cell alone; Format's own sheetName parameter chooses "
                + "the sheet.",
                nameof(cell));
        }

        return new XlsxComment(cell, text);
    }
}
