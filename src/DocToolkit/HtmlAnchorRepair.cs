using System.Text.RegularExpressions;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;

namespace DocToolkit;

/// <summary>
/// Makes a document's internal links resolvable before it is rendered to PDF: promotes obsolete
/// <c>&lt;a name&gt;</c> anchors to <c>id</c>, and drops a link to a target that genuinely is not
/// there.
/// </summary>
/// <remarks>
/// <b>Without this, 27 of 181 real `.gov` pages - 15% - cannot be rendered to PDF at all.</b>
/// Measured 2026-08-17. The renderer validates that every internal link resolves and refuses the
/// whole document when one does not: <i>"PDF bookmark link target 'content' was not found."</i>
///
/// <b>The links are not broken. The anchors are declared the old way.</b> Measured across all 27:
/// <b>every one</b> declares its targets as <c>&lt;a name="x"&gt;</c>, and <b>none</b> is purely
/// dangling. HTML5 made <c>name</c> on <c>&lt;a&gt;</c> obsolete and <c>id</c> its replacement, and
/// browsers still honour both - so these pages navigate correctly in a browser and their authors had
/// no reason to think otherwise. What the converter produces from them is a hyperlink to a bookmark
/// it never created, which is a dangling link in the DOCX too; the PDF stage is merely the first
/// thing strict enough to say so.
///
/// So the repair is to <b>promote, not to delete</b>: the link keeps working, which dropping it
/// would not. Measured on the same corpus, 20 of the 23 pages carrying internal targets are fixed by
/// promotion alone and 3 also hold a genuinely absent target.
///
/// <b>A target that exists nowhere in either form has its link dropped, keeping the text.</b> That
/// is what a browser effectively does - clicking navigates nowhere - and it is the only option left,
/// since there is no anchor to point at.
///
/// <b>This runs on the PDF path only, and deliberately not inside
/// <see cref="HtmlToDocxConverter"/>.</b> Promoting an anchor there would also repair the DOCX, which
/// sounds strictly better and is not free: a document that converts today would start carrying
/// bookmarks it did not carry before, so output would change for conversions that currently succeed.
/// Every property this class claims rests on the opposite - a document reaching the PDF renderer with
/// an unresolvable link fails today, so nothing that works can change. Extending it to the DOCX path
/// is a separate decision, filed rather than taken.
/// </remarks>
internal static class HtmlAnchorRepair
{
    /// <summary>
    /// Only documents that actually contain an internal link are parsed. A bare '#' test would be
    /// useless - every page with a CSS colour has one - so this looks for the href specifically.
    /// </summary>
    private static readonly Regex InternalLink =
        new(@"href\s*=\s*[""']?#", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Returns <paramref name="html"/> with its internal links made resolvable, or the same string
    /// when there is nothing to change.
    /// </summary>
    internal static string Apply(string html)
    {
        if (!InternalLink.IsMatch(html)) return html;

        IHtmlDocument document;
        try
        {
            document = new HtmlParser().ParseDocument(html);
        }
        catch
        {
            // Same reasoning as RowSpanClamp: if it cannot be parsed here it will not be parsed
            // downstream either, and the real converter should produce the real diagnostic.
            return html;
        }

        var changed = Promote(document);
        changed |= DropUnresolvable(document);

        return changed ? document.DocumentElement.OuterHtml : html;
    }

    /// <summary>
    /// Whether a link to <paramref name="target"/> will actually find a bookmark.
    /// </summary>
    /// <remarks>
    /// <b>Not "does any element carry that id".</b> That was the first rule here and it was wrong in
    /// the expensive direction: an id on an <c>&lt;a&gt;</c> or a <c>&lt;td&gt;</c> makes the target
    /// look present while no bookmark is ever created, so the link was left in place and the document
    /// still failed. Only the blocks in <see cref="BookmarkableBlocks"/> produce one.
    /// </remarks>
    private static bool Satisfied(IHtmlDocument document, string target) =>
        document.QuerySelectorAll(BlockSelector)
            .Any(b => string.Equals(b.Id, target, StringComparison.Ordinal));

    /// <summary>
    /// Elements a bookmark is actually created from. <b>Measured, not assumed</b>, because the
    /// obvious repair does not work: putting the <c>id</c> on the <c>&lt;a&gt;</c> itself changes
    /// nothing at all.
    /// </summary>
    /// <remarks>
    /// Measured 2026-08-17 by giving one document the same link and moving the target's <c>id</c>
    /// around. A bookmark appears for an <c>id</c> on <c>&lt;p&gt;</c>, <c>&lt;div&gt;</c> and
    /// <c>&lt;h2&gt;</c>, and does NOT appear for one on <c>&lt;a&gt;</c>, <c>&lt;span&gt;</c> or
    /// <c>&lt;td&gt;</c> - so it is block elements that become bookmarks, and the first attempt at
    /// this repair (<c>&lt;a name="x"&gt;</c> to <c>&lt;a name="x" id="x"&gt;</c>) had no effect on
    /// the corpus whatsoever. <b>Anything added here must be measured the same way</b>: a wrong entry
    /// silently produces a document that still fails.
    /// </remarks>
    private static readonly string[] BookmarkableBlocks =
        ["p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "li", "blockquote", "pre"];

    private static readonly string BlockSelector = string.Join(",", BookmarkableBlocks);

    /// <summary>
    /// Moves each <c>&lt;a name="x"&gt;</c>'s identity onto the nearest block that can carry a
    /// bookmark.
    /// </summary>
    /// <remarks>
    /// <b>The anchor point moves to the start of that block</b>, which is a real if small change:
    /// a link that pointed mid-paragraph now arrives at the paragraph. That is what the author meant
    /// - these are section links - and the alternative is injecting an element into somebody's
    /// document, which is a larger liberty than relabelling one.
    ///
    /// Skipped when the block already has an id: an id must be unique, and creating a duplicate to
    /// fix a link would trade one malformed document for another. Such a link falls through to
    /// <see cref="DropUnresolvable"/>.
    /// </remarks>
    private static bool Promote(IHtmlDocument document)
    {
        // Only targets something actually links to. Relabelling a block nobody points at would be
        // an edit with no purpose, and every edit here is a liberty taken with somebody's document.
        var wanted = document.QuerySelectorAll("a[href]")
            .Select(a => a.GetAttribute("href"))
            .Where(h => h is { Length: > 1 } && h[0] == '#')
            .Select(h => h!.Substring(1))
            .ToHashSet(StringComparer.Ordinal);

        if (wanted.Count == 0) return false;

        var changed = false;
        foreach (var target in wanted)
        {
            if (Satisfied(document, target)) continue;

            // The element that declares the identity, in either form. `[id]` is included because an
            // id on an inline element or a cell reads as present and produces no bookmark.
            var declaring = document.QuerySelectorAll($"a[name], [id]")
                .FirstOrDefault(e => string.Equals(e.GetAttribute("name"), target, StringComparison.Ordinal)
                                  || string.Equals(e.Id, target, StringComparison.Ordinal));
            if (declaring is null) continue;

            var block = declaring.Closest(BlockSelector);
            if (block is null || !string.IsNullOrEmpty(block.Id)) continue;

            block.Id = target;
            changed = true;
        }

        return changed;
    }

    /// <summary>Removes the <c>href</c> from a link whose target exists in no form, keeping its text.</summary>
    private static bool DropUnresolvable(IHtmlDocument document)
    {
        // Materialised before mutating: RemoveAttribute("href") makes an element stop matching the
        // "a[href]" selector it was found by, and AngleSharp's result is live.
        var unresolved = document.QuerySelectorAll("a[href]")
            .Select(link => (link, href: link.GetAttribute("href")))
            .Where(x => x.href is { Length: > 1 } && x.href[0] == '#')
            .Where(x => !Satisfied(document, x.href!.Substring(1)))
            .Select(x => x.link)
            .ToList();

        foreach (var link in unresolved)
            link.RemoveAttribute("href");

        return unresolved.Count > 0;
    }
}
