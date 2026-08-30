namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Reads and updates an existing Markdown document — front matter, headings, tables, and one
/// section's content — without converting to another format first. Registered by
/// <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.
/// </summary>
/// <remarks>
/// <b>No <c>Stream</c> or async overload on any method here</b>, matching the static API: the
/// input is a <see cref="string"/> rather than a document, and every method is CPU-bound rather
/// than genuinely I/O-bound. See <see cref="DocToolkit.MarkdownEditor"/>'s own remarks.
/// </remarks>
public interface IMarkdownEditor
{
    /// <summary>
    /// Every front-matter key in <paramref name="markdown"/>, with its parsed value. A document
    /// with no front matter returns an empty dictionary, never <see langword="null"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="markdown"/> is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The Markdown could not be parsed.</exception>
    IReadOnlyDictionary<string, object> ReadFrontMatter(string markdown);

    /// <summary>
    /// The heading in <paramref name="markdown"/> whose text matches <paramref name="headingText"/>,
    /// or <see langword="null"/> if none does. When more than one heading shares the same text,
    /// the first one in document order is returned.
    /// </summary>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="markdown"/> or <paramref name="headingText"/> is null.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The Markdown could not be parsed.</exception>
    DocToolkit.MarkdownHeading? FindHeading(
        string markdown, string headingText,
        StringComparison comparison = StringComparison.OrdinalIgnoreCase);

    /// <summary>The number of tables in <paramref name="markdown"/>, in document order.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="markdown"/> is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The Markdown could not be parsed.</exception>
    int TableCount(string markdown);

    /// <summary>
    /// The table at <paramref name="index"/>, as rows of cell text — the header row is row 0,
    /// followed by every data row in document order.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="markdown"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="index"/> is negative, or at or beyond <see cref="TableCount(string)"/>.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The Markdown could not be parsed.</exception>
    IReadOnlyList<IReadOnlyList<string>> ReadTable(string markdown, int index);

    /// <summary>
    /// Replaces the content of the section under the heading matching <paramref name="headingText"/>
    /// with <paramref name="newContent"/>, and returns the whole updated document. Front matter and
    /// every other section are left untouched.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// No <b>top-level</b> heading matches <paramref name="headingText"/>.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The Markdown could not be parsed.</exception>
    string ReplaceSection(
        string markdown, string headingText, string newContent,
        StringComparison comparison = StringComparison.OrdinalIgnoreCase);
}
