using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit;

/// <summary>
/// Locates the table rows a repeating-row fill should expand.
///
/// Deliberately does <b>not</b> use <c>Descendants&lt;TableRow&gt;()</c>. That also yields the rows
/// of tables nested inside cells, which would do two wrong things at once: mark a container row as
/// a template row because text further down happens to hold the marker, and sweep an inner row into
/// the outer row's expansion so it is cloned once per record. The same shape of bug — reaching
/// through <c>Descendants&lt;Paragraph&gt;()</c> into <c>w:txbxContent</c> — once deleted text-box
/// content and relocated it into the enclosing paragraph, schema-valid and without an exception.
/// See CLAUDE.md, "Traps in this codebase".
///
/// Detection is scoped the same way as discovery: a row owns the marker only when one of its own
/// cells' <i>direct child</i> paragraphs contains it. Text belonging to a nested table belongs to
/// that table's rows.
///
/// Rows come back innermost-first, so a nested template row is expanded before any row that
/// contains it.
/// </summary>
internal static class TableRowFinder
{
    /// <summary>
    /// Every row at or below <paramref name="scope"/> whose own cells hold <paramref name="marker"/>,
    /// innermost first.
    /// </summary>
    public static IReadOnlyList<TableRow> Find(OpenXmlElement scope, string marker)
    {
        var found = new List<TableRow>();
        Collect(scope, marker, found);
        return found;
    }

    private static void Collect(OpenXmlElement scope, string marker, List<TableRow> found)
    {
        foreach (var table in scope.ChildElements.OfType<Table>())
        {
            // Materialised: expansion mutates the row collection while callers iterate this result.
            foreach (var row in table.ChildElements.OfType<TableRow>().ToList())
            {
                foreach (var cell in row.ChildElements.OfType<TableCell>())
                    Collect(cell, marker, found);

                if (OwnsMarker(row, marker)) found.Add(row);
            }
        }
    }

    private static bool OwnsMarker(TableRow row, string marker)
    {
        foreach (var cell in row.ChildElements.OfType<TableCell>())
        {
            foreach (var paragraph in cell.ChildElements.OfType<Paragraph>())
            {
                if (paragraph.InnerText.Contains(marker, StringComparison.Ordinal))
                    return true;
            }
        }

        return false;
    }
}
