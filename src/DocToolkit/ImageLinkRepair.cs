using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;

namespace DocToolkit;

/// <summary>
/// Gives a link that wraps only an image something to be labelled with, so a table cell holding one
/// can be rendered.
/// </summary>
/// <remarks>
/// <b>Without this, 13 of 181 real `.gov` pages cannot be rendered to PDF</b>, with
/// <i>"Parameter 'linkContents' cannot be empty or whitespace"</i>. Measured 2026-08-17.
///
/// The markup is ordinary: a logo linking home, a "skip navigation" button, a banner. Measured, the
/// same link is fine <b>outside</b> a table and fine <b>inside</b> one when it wraps text - only an
/// image-only link in a cell is refused, and <b>the renderer does not look at the image's
/// <c>alt</c></b> even when one is present.
///
/// <b>This is the first repair here with no browser behaviour to appeal to</b>, which is why it was
/// a maintainer decision rather than a measurement. Both candidates were measured and both work, and
/// they lose different things: unwrapping the link keeps the image and loses the navigation, while
/// using the <c>alt</c> keeps the navigation and replaces the image with words.
///
/// <b>The maintainer chose the alt text</b>, and the default path makes that close to free: this
/// package does not fetch remote images, so on the default path <b>the image is not in the output
/// anyway</b> and unwrapping would leave an empty cell. With
/// <see cref="RemoteImageOptions"/> in play the trade reverses, and a link labelled with its own alt
/// text is still the accessible name the author wrote for it.
///
/// <b>An image with no usable <c>alt</c> falls back to unwrapping</b> - there is nothing to label the
/// link with, and inventing a placeholder would put words in somebody's document that they never
/// wrote.
/// </remarks>
internal static class ImageLinkRepair
{
    /// <summary>
    /// Whether <paramref name="ex"/> is the failure this repair addresses.
    /// </summary>
    /// <remarks>
    /// On the renderer's own message, not the exception type - every conversion failure arrives as a
    /// <see cref="DocumentConversionException"/>, so the type discriminates nothing. Two frames raise
    /// it, <c>PdfTableCell.Validate</c> and the <c>PdfTableCell</c> constructor, and both carry the
    /// same parameter name.
    /// </remarks>
    internal static bool WouldHelp(Exception ex) =>
        (ex.InnerException?.Message ?? string.Empty)
            .Contains("linkContents", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Returns <paramref name="html"/> with image-only links in table cells labelled or unwrapped,
    /// or the same string when there are none.
    /// </summary>
    internal static string Apply(string html)
    {
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

        // Only inside a cell. The same link at the top of a document renders perfectly well, and
        // rewriting it would be an edit with no purpose.
        var textless = document.QuerySelectorAll("td a[href], th a[href]")
            .Where(link => string.IsNullOrWhiteSpace(link.TextContent))
            .ToList();

        if (textless.Count == 0) return html;

        foreach (var link in textless)
        {
            var alt = link.QuerySelectorAll("img")
                .Select(img => img.GetAttribute("alt"))
                .FirstOrDefault(a => !string.IsNullOrWhiteSpace(a));

            if (alt is not null)
            {
                // Appended, so the image survives for a caller who did ask for images to be embedded.
                link.AppendChild(document.CreateTextNode(alt));
            }
            else
            {
                // Nothing to label it with. Unwrap: the children stay exactly where they are, the
                // link goes. Inventing a placeholder would put words in somebody's document.
                while (link.FirstChild is not null)
                    link.Parent!.InsertBefore(link.FirstChild, link);
                link.Remove();
            }
        }

        return document.DocumentElement.OuterHtml;
    }
}
