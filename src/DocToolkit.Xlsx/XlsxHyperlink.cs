namespace DocToolkit;

/// <summary>A hyperlink on one cell, pointing at an external URL.</summary>
public sealed class XlsxHyperlink
{
    private XlsxHyperlink(string cell, string url)
    {
        Cell = cell;
        Url = url;
    }

    /// <summary>The cell carrying the link, such as <c>B2</c>. A single cell, not a range.</summary>
    public string Cell { get; }

    /// <summary>The external URL the cell links to.</summary>
    public string Url { get; }

    /// <summary>Links <paramref name="cell"/> to <paramref name="url"/>.</summary>
    /// <param name="cell">The cell carrying the link, such as <c>B2</c>. A single cell, not a range.</param>
    /// <param name="url">The external URL. Must be an absolute URI.</param>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="cell"/> is blank or names a sheet, or <paramref name="url"/> is blank or not
    /// an absolute URI — checked here, where the caller supplied it, rather than left to surface
    /// later as a write failure with no argument to blame.
    /// </exception>
    public static XlsxHyperlink To(string cell, string url)
    {
        ArgumentNullException.ThrowIfNull(cell);
        ArgumentException.ThrowIfNullOrWhiteSpace(cell);
        ArgumentNullException.ThrowIfNull(url);
        ArgumentException.ThrowIfNullOrWhiteSpace(url);

        if (cell.Contains('!'))
        {
            throw new ArgumentException(
                $"\"{cell}\" names a sheet, and the sheet qualifier is silently discarded rather "
                + "than honoured. Pass the cell alone; Format's own sheetName parameter chooses "
                + "the sheet.",
                nameof(cell));
        }

        if (!Uri.TryCreate(url, UriKind.Absolute, out _))
            throw new ArgumentException($"\"{url}\" is not an absolute URI.", nameof(url));

        return new XlsxHyperlink(cell, url);
    }
}
