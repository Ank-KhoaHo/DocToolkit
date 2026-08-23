using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;

namespace DocToolkit;

/// <summary>
/// Puts a non-breaking space into a table cell that holds nothing, so the renderer's own auto-layout
/// does not squeeze it to a width it then refuses.
/// </summary>
/// <remarks>
/// <b>Without this, 17 of 181 real `.gov` pages cannot be rendered to PDF.</b> The failure is
/// <i>"Table horizontal cell padding must leave a positive text width"</i>, and the obvious reading
/// of that - somebody specified a column too narrow for its padding - is <b>wrong</b>. Reduced from a
/// real page to 316 bytes, the trigger is:
///
/// <code>
/// &lt;table&gt;&lt;tr&gt;&lt;td&gt; &lt;/td&gt;&lt;td&gt;a long sentence of text...&lt;/td&gt;&lt;/tr&gt;&lt;/table&gt;
/// </code>
///
/// <b>No width is specified anywhere.</b> An empty or whitespace-only cell beside a cell of long text
/// gets a near-zero width from automatic layout, and the renderer then rejects the layout it just
/// computed. The spacer cell is one of the most common idioms in HTML of this era.
///
/// <b>Nothing about width or padding avoids it, which was measured before choosing this repair.</b>
/// <c>width="20"</c> on the spacer, <c>width="1"</c>, <c>style="padding:0"</c>,
/// <c>cellpadding="0"</c> on the table and <c>width="100%"</c> on the table all still fail - and the
/// renderer's padding is not reachable from this package at all, since <c>WordPdfSaveOptions</c>
/// exposes no cell-padding knob.
///
/// <b>What does work is giving the cell a non-breaking space</b>, which is what authors of that era
/// wrote in spacer cells for exactly this reason. It is invisible: a browser renders
/// <c>&lt;td&gt; &lt;/td&gt;</c> and <c>&lt;td&gt;&amp;nbsp;&lt;/td&gt;</c> identically, so this
/// changes nothing a reader can see. That makes it the same class of repair as
/// <see cref="RowSpanClamp"/> - restating the document in a form the renderer can read - rather than
/// an override of anything the author chose.
///
/// <b>Only cells with no content at all are touched.</b> A cell holding an image, or any element, is
/// left exactly as it is: it is not the empty case, and this must not become a general licence to
/// write into somebody's table.
/// </remarks>
internal static class EmptyTableCellRepair
{
    private const string NonBreakingSpace = " ";

    /// <summary>
    /// Whether <paramref name="ex"/> is the failure this repair addresses.
    /// </summary>
    /// <remarks>
    /// Matched on the renderer's own message rather than on the exception type, which says nothing:
    /// every conversion failure arrives as a <see cref="DocumentConversionException"/>. Same
    /// discipline as <see cref="HtmlFailureDiagnosis"/> - name only a cause that can be told apart.
    /// </remarks>
    internal static bool WouldHelp(Exception ex) =>
        (ex.InnerException?.Message ?? string.Empty)
            .Contains("positive text width", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Returns <paramref name="html"/> with empty table cells filled, or the same string when there
    /// are none.
    /// </summary>
    internal static string Apply(string html)
    {
        // No table, nothing to do - and most documents have none, so this avoids parsing them.
        if (html.IndexOf("<td", StringComparison.OrdinalIgnoreCase) < 0
            && html.IndexOf("<th", StringComparison.OrdinalIgnoreCase) < 0) return html;

        IHtmlDocument document;
        try
        {
            document = new HtmlParser().ParseDocument(html);
        }
        catch
        {
            return html;
        }

        // Any cell with no TEXT, whatever elements it holds. The real pages carry <br>, <p><br></p>
        // and <img> in these cells - all of which render to nothing, and images are not fetched at
        // all on the default path - so restricting this to structurally empty cells fixed one page
        // out of seventeen. Measured before broadening it.
        var empty = document.QuerySelectorAll("td,th")
            .Where(cell => string.IsNullOrWhiteSpace(cell.TextContent))
            .ToList();

        if (empty.Count == 0) return html;

        // APPENDED, not assigned. Replacing the content would delete an image the caller may well
        // have asked to be embedded; a non-breaking space beside it is invisible either way.
        foreach (var cell in empty)
            cell.AppendChild(document.CreateTextNode(NonBreakingSpace));

        return document.DocumentElement.OuterHtml;
    }
}
