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
