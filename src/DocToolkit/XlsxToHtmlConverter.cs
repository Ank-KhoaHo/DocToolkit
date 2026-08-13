using System.Text;

namespace DocToolkit;

/// <summary>
/// Exports one sheet of a workbook as an HTML table.
/// </summary>
/// <remarks>
/// <b>A fragment, not a document</b> — a bare <c>&lt;table&gt;</c> with no <c>&lt;html&gt;</c>
/// wrapper, deliberately the opposite of <see cref="DocxToHtmlConverter"/>. A sheet is a component
/// of a page rather than a page, so the common case is embedding it; wrapping is one line for a
/// caller who wants a whole document, while unwrapping means parsing HTML.
///
/// <b>Cell text is culture-invariant</b>, matching <see cref="XlsxToCsvConverter"/> so the two
/// exporters cannot disagree about what a cell says. See
/// <see cref="WorkbookEditor.ReadSheet(byte[], string)"/> for why that differs from reading a sheet
/// as data.
///
/// <b>Every cell is escaped.</b> A workbook is untrusted input: a cell containing
/// <c>&lt;script&gt;</c> must arrive as text, not as markup.
/// </remarks>
public static class XlsxToHtmlConverter
{
    private const string FailureMessage = "Failed to convert XLSX to HTML.";

    /// <summary>Exports <paramref name="sheetName"/> as an HTML <c>&lt;table&gt;</c> fragment.</summary>
    /// <remarks>
    /// The sheet's first row becomes <c>&lt;th&gt;</c> cells inside a <c>&lt;thead&gt;</c>. That is
    /// a presentation guess — a spreadsheet does not record whether its first row is a header — and
    /// it is the guess that is right far more often than not. A caller who disagrees can restyle
    /// or replace the element; the alternative, emitting no header at all, cannot be undone from
    /// the output.
    /// </remarks>
    /// <param name="xlsx">The workbook to read.</param>
    /// <param name="sheetName">The sheet to export.</param>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="xlsx"/> is empty, or <paramref name="sheetName"/> is blank.
    /// </exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, or the sheet does not exist.
    /// </exception>
    public static string Convert(byte[] xlsx, string sheetName)
        => Format(WorkbookEditor.ReadSheetInvariant(xlsx, sheetName));

    /// <summary>
    /// Reads a workbook from <paramref name="source"/> and exports <paramref name="sheetName"/> as
    /// an HTML <c>&lt;table&gt;</c> fragment.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and is neither disposed, closed nor
    /// sought, so it may be forward-only.
    /// </summary>
    /// <inheritdoc cref="Convert(byte[], string)"/>
    /// <param name="source">The stream the workbook is read from.</param>
    /// <param name="sheetName">The sheet to export.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    public static async Task<string> ConvertAsync(
        Stream source, string sheetName, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ct.ThrowIfCancellationRequested();

        using var xlsx = await StreamPipeline
            .DrainAsync(source, "XLSX content was empty.", nameof(source), FailureMessage, ct)
            .ConfigureAwait(false);

        return Convert(xlsx.ToArray(), sheetName);
    }

    private static string Format(IReadOnlyList<IReadOnlyList<string>> rows)
    {
        var html = new StringBuilder("<table>\n");

        for (var r = 0; r < rows.Count; r++)
        {
            var isHeader = r == 0;
            if (isHeader) html.Append("  <thead>\n");
            if (r == 1) html.Append("  <tbody>\n");

            html.Append("    <tr>");
            foreach (var cell in rows[r])
            {
                html.Append(isHeader ? "<th>" : "<td>");
                Escape(html, cell);
                html.Append(isHeader ? "</th>" : "</td>");
            }

            html.Append("</tr>\n");
            if (isHeader) html.Append("  </thead>\n");
        }

        if (rows.Count > 1) html.Append("  </tbody>\n");
        html.Append("</table>");

        return html.ToString();
    }

    /// <summary>
    /// Escapes the five characters that can change how markup parses.
    /// </summary>
    /// <remarks>
    /// Both quote forms are escaped even though these values only ever land in element content,
    /// where neither can terminate anything. It costs nothing, and it means a value copied from
    /// this output into an attribute by some later change cannot break out of it — the class of
    /// bug that only shows up after the code that made it safe has moved.
    /// </remarks>
    private static void Escape(StringBuilder html, string text)
    {
        foreach (var c in text)
        {
            switch (c)
            {
                case '&': html.Append("&amp;"); break;
                case '<': html.Append("&lt;"); break;
                case '>': html.Append("&gt;"); break;
                case '"': html.Append("&quot;"); break;
                case '\'': html.Append("&#39;"); break;
                default: html.Append(c); break;
            }
        }
    }
}
