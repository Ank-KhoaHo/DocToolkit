using System.Text.RegularExpressions;
using AngleSharp.Dom;
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
    /// <remarks>
    /// <b>The fragment need not be at the start of the href</b>, and matching only <c>href="#</c>
    /// made this gate narrower than <see cref="Fragment"/> - so <c>page.html#name</c> was recognised
    /// as internal by the logic and then never reached it, because the document was rejected here
    /// first. Caught by a test rather than by the corpus, which is the cheaper way round.
    /// </remarks>
    private static readonly Regex InternalLink =
        new(@"href\s*=\s*(""[^""]*#|'[^']*#|[^\s>""']*#)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

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

        // ONE pass over the document's blocks, shared by both repairs. Promote may satisfy a
        // target by relabelling a block, and DropUnresolvable must see that - otherwise it strips
        // the href off a link that was just made to work.
        var satisfied = SatisfiedIds(document);

        var changed = Promote(document, satisfied);
        changed |= DropUnresolvable(document, satisfied);

        return changed ? document.DocumentElement.OuterHtml : html;
    }

    /// <summary>Every id that would actually produce a bookmark, collected in ONE pass.</summary>
    /// <remarks>
    /// <b>Not "every id in the document".</b> That was the first rule here and it was wrong in the
    /// expensive direction: an id on an <c>&lt;a&gt;</c> or a <c>&lt;td&gt;</c> makes a target look
    /// present while no bookmark is ever created, so the link was left in place and the document
    /// still failed. Only the blocks in <see cref="BookmarkableBlocks"/> produce one.
    ///
    /// <b>This replaced a per-target predicate that re-queried the WHOLE document every time.</b>
    /// <see cref="Promote"/> called it once per distinct target and <see cref="DropUnresolvable"/>
    /// once per link, so a page with 300 in-page links ran 600 full-document queries. It is paid
    /// on the common path: <see cref="Apply"/> gates only on a regex, so any page carrying a
    /// single in-page link enters it.
    ///
    /// <para><b>Measured 2026-08-22 on a page of 2000 blocks and 300 links: 258 ms to 20 ms.</b>
    /// 14 ms of that 20 is AngleSharp parsing the document, which no change here can avoid, so the
    /// repair work itself went from roughly 245 ms to roughly 6 ms. The AFTER figure repeats
    /// stably - 20 and 21 ms on consecutive best-of-5 runs - while the BEFORE figure did not,
    /// ranging 142-258 ms on the same code. So read this as an order of magnitude, not a ratio;
    /// this repository has been burned twice by wall-clock numbers taken on a busy machine.</para>
    ///
    /// <para>The deterministic version of the same claim, which is the one to trust: the document
    /// was queried once per link plus once per distinct target, and is now queried <b>once</b>.</para>
    ///
    /// <para><b>Exactly equivalent to the predicate it replaced, not an approximation.</b> A
    /// target comes from <see cref="Fragment"/>, which returns null rather than an empty string
    /// for a bare <c>#</c>, so no target is ever empty and skipping empty ids cannot change an
    /// answer. <see cref="YieldsBookmark"/> is still only asked about blocks that carry an id,
    /// which is what the old short-circuiting <c>&amp;&amp;</c> already did.</para>
    ///
    /// <para><b>Safe to MAINTAIN across Promote's mutations.</b> The only edit that loop makes is
    /// <c>block.Id = target</c>, on a block that had no id and does yield a bookmark - so no id is
    /// ever removed, the set never shrinks, and it grows by exactly the target just promoted.</para>
    /// </remarks>
    private static HashSet<string> SatisfiedIds(IHtmlDocument document)
    {
        return document.QuerySelectorAll(BlockSelector)
            .Where(b => !string.IsNullOrEmpty(b.Id) && YieldsBookmark(b))
            .Select(b => b.Id)
            .ToHashSet(StringComparer.Ordinal);
    }

    /// <summary>
    /// Whether a block carrying an id would actually produce a bookmark.
    /// </summary>
    /// <remarks>
    /// <b>Being a block is not enough, and the two ways it is not enough have opposite costs.</b>
    /// Measured 2026-08-17:
    ///
    /// <list type="bullet">
    /// <item><description>A block holding <b>no text of its own</b> - an empty <c>&lt;div&gt;</c>, or
    /// one containing only an <c>&lt;a name&gt;</c> - does not merely fail to make a bookmark, it
    /// makes the render throw <c>InvalidOperationException</c> from
    /// <c>ThrowNoElementsException</c>. So labelling one turns a legible "target not found" into an
    /// opaque crash, which is a worse document than the one we started with. <b>This is not
    /// hypothetical: the first version of this class did exactly that to eight corpus pages.</b>
    /// </description></item>
    /// <item><description>A block whose only content is a <b>table</b> produces no paragraph of its
    /// own, so no bookmark - it simply stays broken.</description></item>
    /// </list>
    ///
    /// Both are avoided by the same test: the block must have text that is not inside a nested table.
    /// </remarks>
    private static bool YieldsBookmark(IElement block) =>
        !string.IsNullOrWhiteSpace(TextOutsideTables(block));

    private static string TextOutsideTables(IElement block)
    {
        // Concat rather than a StringBuilder loop: only emptiness is ever asked of the result, and
        // the filter is the point of the method - saying it as a Where makes that legible.
        return string.Concat(block.Descendants<IText>()
            .Where(node => node.ParentElement?.Closest("table") is null)
            .Select(node => node.Text));
    }

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
        ["p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "li", "blockquote"];

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
    private static bool Promote(IHtmlDocument document, HashSet<string> satisfied)
    {
        // Only targets something actually links to. Relabelling a block nobody points at would be
        // an edit with no purpose, and every edit here is a liberty taken with somebody's document.
        var wanted = document.QuerySelectorAll("a[href]")
            .Select(a => Fragment(a.GetAttribute("href")))
            .Where(t => t is not null)
            .Select(t => t!)
            .ToHashSet(StringComparer.Ordinal);

        if (wanted.Count == 0) return false;

        var changed = false;
        foreach (var target in wanted)
        {
            if (satisfied.Contains(target)) continue;

            // The element that declares the identity, in either form. `[id]` is included because an
            // id on an inline element or a cell reads as present and produces no bookmark.
            var declaring = document.QuerySelectorAll($"a[name], [id]")
                .FirstOrDefault(e => string.Equals(e.GetAttribute("name"), target, StringComparison.Ordinal)
                                  || string.Equals(e.Id, target, StringComparison.Ordinal));
            if (declaring is null) continue;

            var block = declaring.Closest(BlockSelector);
            if (block is null || !string.IsNullOrEmpty(block.Id)) continue;
            if (!YieldsBookmark(block)) continue;

            block.Id = target;

            // LOAD-BEARING, not bookkeeping: DropUnresolvable reads this same set afterwards, and
            // without this line it would strip the href off the link this promotion just fixed.
            satisfied.Add(target);

            changed = true;
        }

        return changed;
    }

    /// <summary>
    /// The bookmark a link points at, or <see langword="null"/> when it points outside the document.
    /// </summary>
    /// <remarks>
    /// <b>A bare <c>#name</c> is not the only internal form, and assuming it was left four corpus
    /// pages failing with no visible cause.</b> Measured: <c>page.html#privacy</c> - a RELATIVE url
    /// carrying a fragment - is turned into an internal bookmark link too, while
    /// <c>https://host/page.html#privacy</c> is left as an ordinary external link. So the test is the
    /// presence of a fragment on anything that is not absolute, which is what the converter itself
    /// does.
    /// </remarks>
    private static string? Fragment(string? href)
    {
        if (string.IsNullOrEmpty(href)) return null;
        if (Uri.TryCreate(href, UriKind.Absolute, out _)) return null;

        var hash = href.IndexOf('#');
        if (hash < 0 || hash == href.Length - 1) return null;
        return href.Substring(hash + 1);
    }

    /// <summary>Removes the <c>href</c> from a link whose target exists in no form, keeping its text.</summary>
    private static bool DropUnresolvable(IHtmlDocument document, HashSet<string> satisfied)
    {
        // Materialised before mutating: RemoveAttribute("href") makes an element stop matching the
        // "a[href]" selector it was found by, and AngleSharp's result is live.
        var unresolved = document.QuerySelectorAll("a[href]")
            .Select(link => (link, target: Fragment(link.GetAttribute("href"))))
            .Where(x => x.target is not null)
            .Where(x => !satisfied.Contains(x.target!))
            .Select(x => x.link)
            .ToList();

        foreach (var link in unresolved)
            link.RemoveAttribute("href");

        return unresolved.Count > 0;
    }
}
