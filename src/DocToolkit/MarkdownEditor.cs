using OfficeIMO.Markdown;

namespace DocToolkit;

/// <summary>
/// Reads and updates an existing Markdown document — front matter, headings, tables, and one
/// section's content — without converting to another format first.
/// </summary>
/// <remarks>
/// <b>No <c>Stream source</c> or async overload, on any method here.</b> The input is a
/// <see cref="string"/> rather than a document, the same reason
/// <see cref="MarkdownToDocxConverter"/> has none.
/// </remarks>
public static class MarkdownEditor
{
    private const string FailureMessage =
        "Failed to read Markdown. See the inner exception for details.";

    /// <summary>
    /// Every front-matter key in <paramref name="markdown"/>, with its parsed value. A document
    /// with no front matter returns an empty dictionary, never <see langword="null"/>.
    /// </summary>
    /// <param name="markdown">The Markdown to read.</param>
    /// <remarks>
    /// A YAML scalar's runtime type follows the underlying reader: a quoted or bare word is a
    /// <see cref="string"/>, a number is a <see cref="double"/> (never <c>int</c> or
    /// <c>long</c> — <c>version: 3</c> comes back as <c>3.0</c>), and <c>true</c>/<c>false</c> is
    /// a <see cref="bool"/>.
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="markdown"/> is null.</exception>
    /// <exception cref="DocumentConversionException">The Markdown could not be parsed.</exception>
    public static IReadOnlyDictionary<string, object> ReadFrontMatter(string markdown)
    {
        ArgumentNullException.ThrowIfNull(markdown);

        var doc = ParseOrThrow(markdown);

        var result = new Dictionary<string, object>();
        foreach (var entry in doc.FrontMatterEntries)
        {
            result[entry.Key] = entry.Value!;
        }

        return result;
    }

    /// <summary>
    /// The heading in <paramref name="markdown"/> whose text matches <paramref name="headingText"/>,
    /// or <see langword="null"/> if none does. When more than one heading shares the same text,
    /// the first one in document order is returned.
    /// </summary>
    /// <param name="markdown">The Markdown to search.</param>
    /// <param name="headingText">
    /// The heading's text to match, without the leading <c>#</c> markers.
    /// </param>
    /// <param name="comparison">How <paramref name="headingText"/> is compared. Case-insensitive by default.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="markdown"/> or <paramref name="headingText"/> is null.
    /// </exception>
    /// <exception cref="DocumentConversionException">The Markdown could not be parsed.</exception>
    public static MarkdownHeading? FindHeading(
        string markdown, string headingText,
        StringComparison comparison = StringComparison.OrdinalIgnoreCase)
    {
        ArgumentNullException.ThrowIfNull(markdown);
        ArgumentNullException.ThrowIfNull(headingText);

        var doc = ParseOrThrow(markdown);
        var found = doc.FindHeading(headingText, comparison);

        return found is null ? null : new MarkdownHeading(found.Level, found.Text, found.Anchor);
    }

    /// <summary>The number of tables in <paramref name="markdown"/>, in document order.</summary>
    /// <param name="markdown">The Markdown to read.</param>
    /// <exception cref="ArgumentNullException"><paramref name="markdown"/> is null.</exception>
    /// <exception cref="DocumentConversionException">The Markdown could not be parsed.</exception>
    public static int TableCount(string markdown)
    {
        ArgumentNullException.ThrowIfNull(markdown);

        var doc = ParseOrThrow(markdown);
        return doc.DescendantTables().Count();
    }

    /// <summary>
    /// The table at <paramref name="index"/>, as rows of cell text — the header row is row 0,
    /// followed by every data row in document order. A row is returned with the shape it has: a
    /// row with fewer or more cells than its neighbours is not padded into a rectangle.
    /// </summary>
    /// <param name="markdown">The Markdown to read.</param>
    /// <param name="index">
    /// <b>0-based</b>, indexing what <see cref="TableCount(string)"/> reports.
    /// </param>
    /// <exception cref="ArgumentNullException"><paramref name="markdown"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="index"/> is negative, or at or beyond <see cref="TableCount(string)"/>.
    /// </exception>
    /// <exception cref="DocumentConversionException">The Markdown could not be parsed.</exception>
    public static IReadOnlyList<IReadOnlyList<string>> ReadTable(string markdown, int index)
    {
        ArgumentNullException.ThrowIfNull(markdown);
        ArgumentOutOfRangeException.ThrowIfNegative(index);

        var doc = ParseOrThrow(markdown);
        var tables = doc.DescendantTables().ToList();

        if (index >= tables.Count)
        {
            throw new ArgumentOutOfRangeException(
                nameof(index), index,
                $"Table {index} was requested from a document with {tables.Count} table(s).");
        }

        var table = tables[index];
        var rows = new List<IReadOnlyList<string>> { table.Headers };
        rows.AddRange(table.Rows);
        return rows;
    }

    /// <summary>
    /// Parses <paramref name="markdown"/>, wrapping any failure the reader itself raises in a
    /// <see cref="DocumentConversionException"/> — the one place every method in this class does
    /// so, matching <see cref="MarkdownToDocxConverter.ConvertCore"/>'s own wrapping around the
    /// same call.
    /// </summary>
    private static MarkdownDoc ParseOrThrow(string markdown)
    {
        try
        {
            return MarkdownReader.Parse(markdown);
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException(
                MarkdownFailureDiagnosis.Describe(ex, markdown) ?? FailureMessage, ex);
        }
    }
}

/// <summary>One heading found in a Markdown document by <see cref="MarkdownEditor.FindHeading"/>.</summary>
public sealed class MarkdownHeading
{
    internal MarkdownHeading(int level, string text, string anchor)
    {
        Level = level;
        Text = text;
        Anchor = anchor;
    }

    /// <summary>The heading's level: 1 for <c>#</c>, 2 for <c>##</c>, and so on.</summary>
    public int Level { get; }

    /// <summary>The heading's text, with any inline formatting markers stripped.</summary>
    public string Text { get; }

    /// <summary>
    /// The slug an in-document anchor link to this heading would use — for example
    /// <c>changed</c> for a heading reading "Changed".
    /// </summary>
    public string Anchor { get; }
}
