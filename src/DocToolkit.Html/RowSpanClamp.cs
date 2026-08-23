using System.Globalization;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;

namespace DocToolkit;

/// <summary>
/// Clamps a table cell's <c>rowspan</c> to the rows that actually exist, so a document carrying one
/// that overruns can be converted instead of crashing.
/// </summary>
/// <remarks>
/// <b>Without this, 14 of 179 real `.gov` pages - 7.7% - cannot be converted at all.</b> Measured
/// 2026-08-17 against a public crawl; it was the single most frequent conversion failure in the set.
/// The HTML parser this package uses sizes an accumulator to the rows a table part actually has and
/// then indexes it with the rowspan, unclamped, so a cell reaching past the last row writes out of
/// bounds and raises an <see cref="IndexOutOfRangeException"/> naming no table.
///
/// <b>The markup is not malformed, which is the whole reason this is worth repairing rather than
/// refusing.</b> Browsers clamp such a rowspan to the rows that exist, so these pages render
/// correctly in a browser and the author had no reason to think anything was wrong. The corpus held
/// spans of 2, 3, 14, 100 and 103 against tables of one to three rows - a <c>rowspan="100"</c> is
/// somebody wanting a cell to reach the bottom of a table and picking a number bigger than the row
/// count. Clamping is therefore not a guess about intent: it produces exactly what every browser
/// already shows.
///
/// <b>It cannot change any document that converts today, by construction.</b> The input string is
/// returned untouched unless a cell is actually found to overrun - and every document containing one
/// currently throws, so there is no input whose output changes, only inputs that begin to have one.
/// That is stronger than the same argument for <see cref="ListMarkerSubstitution"/>, which had to
/// justify a visible glyph change; here the rendered result is what the page already looked like.
///
/// <b>Why AngleSharp rather than scanning the string.</b> The clamp has to agree with the PARSER
/// about which cells sit in which part and how many rows that part has, and real pages disagree with
/// any simpler model: they nest tables, omit <c>&lt;/td&gt;</c>, and put stray cells outside a row.
/// A regex is not merely fragile here but measurably wrong - a non-greedy <c>&lt;table&gt;.*?
/// &lt;/table&gt;</c> stops at the first close tag and miscounts every nested table, which produced
/// a false negative on the smallest reproduction in the corpus during this investigation. AngleSharp
/// is the parser HtmlToOpenXml itself uses, so parsing with it is the only way to see the document
/// the way the code downstream will see it.
/// </remarks>
internal static class RowSpanClamp
{
    /// <summary>
    /// Returns <paramref name="html"/> with overrunning row spans clamped, or the same string when
    /// there is nothing to change.
    /// </summary>
    /// <remarks>
    /// Returning the input unchanged matters more than the usual "do not pay for what you do not
    /// use": it is what makes this incapable of altering a document that already converts. The
    /// parse-and-serialise round trip happens ONLY for documents that would otherwise throw.
    /// </remarks>
    internal static string Apply(string html)
    {
        // Cheap reject on the raw string before parsing anything. Most documents have no rowspan at
        // all, and parsing every one of them to discover that would be a real cost on the common
        // path - HTML parsing is not free the way a substring search is.
        if (html.IndexOf("rowspan", StringComparison.OrdinalIgnoreCase) < 0) return html;

        IHtmlDocument document;
        try
        {
            document = new HtmlParser().ParseDocument(html);
        }
        catch
        {
            // If it cannot be parsed here it will not be parsed downstream either. Hand the original
            // on and let the real converter produce the real diagnostic, rather than inventing one
            // from a pre-pass the caller never asked for.
            return html;
        }

        var clamped = false;
        foreach (var table in document.QuerySelectorAll("table").OfType<IHtmlTableElement>())
            clamped |= ClampTable(table);

        return clamped ? document.DocumentElement.OuterHtml : html;
    }

    /// <summary>
    /// Clamps every overrunning cell in one table, counting rows PER SECTION.
    /// </summary>
    /// <remarks>
    /// Per section, not per table, because that is how the failing code counts: it allocates one
    /// accumulator per table part and indexes it with the row's position within that part. A cell in
    /// a <c>thead</c> of one row with <c>rowspan="2"</c> overruns even when the <c>tbody</c> below it
    /// has fifty rows.
    /// </remarks>
    private static bool ClampTable(IHtmlTableElement table)
    {
        var clamped = false;
        foreach (var section in Sections(table))
        {
            var rows = section.ToList();
            for (var i = 0; i < rows.Count; i++)
            {
                var remaining = rows.Count - i;
                var span = remaining.ToString(CultureInfo.InvariantCulture);

                foreach (var cell in rows[i].Cells.Where(c => c.RowSpan > remaining))
                {
                    // SetAttribute rather than the RowSpan property: the property is what AngleSharp
                    // computed, and writing the attribute is what the re-parse downstream will read.
                    cell.SetAttribute("rowspan", span);
                    clamped = true;
                }
            }
        }
        return clamped;
    }

    /// <summary>
    /// The row groups of a table: <c>thead</c>, each <c>tbody</c>, then <c>tfoot</c>.
    /// </summary>
    /// <remarks>
    /// A table whose rows sit directly under <c>&lt;table&gt;</c> still has a <c>tbody</c> here - the
    /// HTML parsing specification inserts one, and AngleSharp implements that - so the fallback below
    /// is for the case where a section-less table somehow yields no bodies rather than for ordinary
    /// markup. Falling back to the table's own row list keeps a cell from being missed entirely,
    /// which would put the crash back.
    /// </remarks>
    private static IEnumerable<IEnumerable<IHtmlTableRowElement>> Sections(IHtmlTableElement table)
    {
        var any = false;

        if (table.Head is not null) { any = true; yield return table.Head.Rows; }
        foreach (var body in table.Bodies) { any = true; yield return body.Rows; }
        if (table.Foot is not null) { any = true; yield return table.Foot.Rows; }

        if (!any) yield return table.Rows;
    }
}
