using OfficeIMO.Markdown;

namespace DocToolkit;

/// <summary>
/// Reads and updates an existing Markdown document — front matter, headings, tables, and one
/// section's content — without converting to another format first.
/// </summary>
/// <remarks>
/// <para>
/// <b>No <c>Stream source</c> overload, on any method here.</b> The input is a
/// <see cref="string"/> rather than a document, so a caller holding bytes decides their own
/// encoding — the same reason <see cref="MarkdownToDocxConverter"/> has no <c>Stream source</c>
/// overload either.
/// </para>
/// <para>
/// <b>No async overload either, for a different reason.</b> An overload here is async only where
/// there is real I/O to await — that is what a <c>Stream</c> overload earns, because draining or
/// emitting a stream genuinely awaits. Every method on this class is CPU-bound: it parses a string
/// and returns a value. Wrapping that in a <see cref="System.Threading.Tasks.Task"/> would make it
/// look async without making it so.
/// </para>
/// <para>
/// <b>No <c>*Core</c> split.</b> That convention exists so a <c>byte[]</c> overload and a
/// <c>Stream</c> overload cannot drift apart; no method here has more than one overload, so there
/// is nothing for a split to keep in sync.
/// </para>
/// <para>
/// <b>No dependency-injection mirror yet, deliberately deferred rather than argued by analogy.</b>
/// <c>DocToolkit.Extensions.DependencyInjection</c> references the <b>published</b> core package
/// (see that project's own notes), so the service delegating to this class — the mirror itself —
/// cannot be implemented before this class has shipped. That is a scheduling constraint, not a
/// design decision to leave it unmirrored.
/// </para>
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
    /// <para>
    /// <b>Values are whatever the underlying reader produced, and not every YAML shape survives.</b>
    /// Every statement below was measured against the pinned <c>OfficeIMO.Markdown</c> 3.2.6 rather
    /// than inferred from YAML in general.
    /// </para>
    /// <para>
    /// A <b>scalar</b> arrives as one of three runtime types: a quoted or bare word is a
    /// <see cref="string"/>, a number is a <see cref="double"/> (never <c>int</c> or <c>long</c> —
    /// <c>version: 3</c> comes back as <c>3.0</c>), and <c>true</c>/<c>false</c> is a
    /// <see cref="bool"/>. An absent value is the empty string: <c>key:</c> with nothing after it,
    /// <c>key: null</c>, <c>key: ~</c> and <c>key: ""</c> are read as <c>""</c>, <c>"null"</c>,
    /// <c>"~"</c> and <c>""</c> respectively. No value is ever <see langword="null"/>.
    /// </para>
    /// <para>
    /// An <b>inline sequence</b> — <c>tags: [alpha, beta]</c> — is a fourth runtime type, a
    /// <c>List&lt;string&gt;</c>. Its items are always strings, so <c>nums: [1, 2]</c> yields
    /// <c>"1"</c> and <c>"2"</c> rather than two <see cref="double"/>s, and <c>tags: []</c> yields
    /// an empty list.
    /// </para>
    /// <para>
    /// A <b>block sequence</b> — a <c>tags:</c> line with indented <c>- alpha</c> / <c>- beta</c>
    /// items beneath it — is <b>not read</b>. The key is present and maps to an empty string; the
    /// items are lost entirely, with nothing raised. Write the sequence inline if you need to read
    /// it back.
    /// </para>
    /// <para>
    /// A <b>nested mapping</b> does not nest — it <b>flattens</b>. Given an <c>author:</c> line with
    /// indented <c>name:</c> and <c>email:</c> beneath it, <c>author</c> maps to an empty string
    /// while <c>name</c> and <c>email</c> appear as their own top-level keys of the returned
    /// dictionary; deeper indentation flattens the same way, to the same one level. If a flattened
    /// key collides with another key in the same front matter, only one entry survives and it
    /// carries the <b>later</b> value — silently, exactly as a duplicate top-level key does. A
    /// single-line <b>inline</b> mapping is neither parsed nor flattened: <c>author: {name: Ada}</c>
    /// comes back as the literal string <c>{name: Ada}</c>.
    /// </para>
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
            // The `!` is measured rather than assumed, so a future reader need not re-derive it.
            // `Entry.Value` came back non-null for every front-matter shape tried against
            // OfficeIMO.Markdown 3.2.6: an empty value, the literal `null`, `~`, an empty quoted
            // string, duplicate keys, a block sequence, an inline sequence (including an empty
            // one), a nested mapping and an inline mapping. An absent value is the empty string.
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
    /// <remarks>
    /// <b>Every heading is searched, including a nested one</b> — a heading inside a blockquote or a
    /// list item is found here like any other. That is deliberately wider than
    /// <see cref="ReplaceSection"/>, which considers top-level headings only because it has to index
    /// the document's own block list to find a section's boundaries. Reading a heading has no such
    /// constraint, so it does not inherit the restriction.
    /// </remarks>
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

        // Genuine read-only copies, not the parser's own collections. `table.Headers` is a
        // concrete `List<string>` and each entry of `table.Rows` is a `string[]`; handing either
        // straight back behind an `IReadOnlyList<string>` lets a caller cast the reference and
        // mutate the parsed document (a List via Add/Remove/indexer, an array via its indexer) —
        // not what read-only means. `DocxEditor.ReadTableCore` builds a fresh list per row for the
        // same reason, and this matches that convention.
        var rows = new List<IReadOnlyList<string>> { new List<string>(table.Headers).AsReadOnly() };
        foreach (var row in table.Rows)
        {
            rows.Add(new List<string>(row).AsReadOnly());
        }

        return rows;
    }

    /// <summary>
    /// Replaces the content of the section under the heading matching <paramref name="headingText"/>
    /// with <paramref name="newContent"/>, and returns the whole updated document. Front matter and
    /// every other section are left untouched.
    /// </summary>
    /// <param name="markdown">The Markdown to edit.</param>
    /// <param name="headingText">
    /// The target heading's text, without the leading <c>#</c> markers, matched against the
    /// document's <b>top-level</b> headings only. This is deliberately narrower than
    /// <see cref="FindHeading"/>, which also matches a heading nested inside a blockquote or a list
    /// item — see the remarks for why the two differ.
    /// </param>
    /// <param name="newContent">
    /// The section's new body, inserted verbatim in place of the <b>blocks</b> between the heading's
    /// own line and the start of the next section. Include your own surrounding newlines — this
    /// method does not add or normalise whitespace around what you pass. "Blocks" is meant
    /// literally: a CommonMark link reference definition leaves no block behind, so one immediately
    /// after the heading is kept rather than replaced. See the remarks.
    /// </param>
    /// <param name="comparison">How <paramref name="headingText"/> is compared. Case-insensitive by default.</param>
    /// <remarks>
    /// <para>
    /// A section runs from immediately after the target heading's own line to the start of the
    /// next heading at the <b>same or a shallower level</b> (a level-2 target's section can only be
    /// closed by another level-1 or level-2 heading, never by a level-3 one), or to the end of the
    /// document if there is no such heading.
    /// </para>
    /// <para>
    /// <b>Line endings that came from <paramref name="markdown"/> are normalised to <c>\n</c></b>,
    /// and only <c>\r\n</c> is recognised as a line ending here — a lone <c>\r</c> is left exactly
    /// as it is. Normalising is not a stylistic choice: every boundary above is an offset that
    /// <c>OfficeIMO.Markdown</c> computes against <b>LF-normalised</b> text, so splicing those
    /// offsets into an original <c>\r\n</c> string would misalign every one of them by the number of
    /// line breaks preceding it — silently truncating the heading and corrupting an unrelated line
    /// further down. The input is normalised once, then parsed and spliced as one consistent string.
    /// <paramref name="newContent"/> is inserted verbatim and keeps whatever line endings you pass
    /// it, so a <c>\r\n</c> written there survives into the result.
    /// </para>
    /// <para>
    /// <b>Only top-level headings are ever considered.</b> A heading nested inside a blockquote or a
    /// list item is not a candidate — it is not searched, rather than found and then refused — even
    /// when no top-level heading shares its text, in which case the call reports that no heading
    /// matched. It has to be invisible rather than merely refused: the section boundaries are
    /// computed by walking the document's own top-level block list, and searching the whole document
    /// first made a legitimate top-level heading uneditable whenever a nested heading earlier in the
    /// document happened to share its text.
    /// </para>
    /// <para>
    /// <b>One measured limitation.</b> A CommonMark link reference definition
    /// (<c>[label]: https://example.com</c>) is consumed into the parser's link registry and leaves
    /// no block in the document, so the body's start is found past it. One sitting immediately after
    /// the target heading is therefore not treated as part of the section body and survives
    /// untouched even when the rest of the section is replaced. One sitting after other content in
    /// the same section falls inside the replaced range and does go. Rewriting the splice to track
    /// them was judged the worse trade — the offset arithmetic here is correct and heavily pinned,
    /// and link reference definitions are rare in the section bodies this method exists for.
    /// </para>
    /// </remarks>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// No <b>top-level</b> heading matches <paramref name="headingText"/>.
    /// </exception>
    /// <exception cref="DocumentConversionException">The Markdown could not be parsed.</exception>
    public static string ReplaceSection(
        string markdown, string headingText, string newContent,
        StringComparison comparison = StringComparison.OrdinalIgnoreCase)
    {
        ArgumentNullException.ThrowIfNull(markdown);
        ArgumentNullException.ThrowIfNull(headingText);
        ArgumentNullException.ThrowIfNull(newContent);

        // Normalise ONCE, up front, and then parse and splice against this same string
        // throughout. See the <remarks> above for why: SourceSpan offsets are measured against
        // LF-normalised text, so `markdown` and the offsets taken from it are two different
        // coordinate systems the moment the input carries a single \r\n.
        //
        // This is deliberately local to ReplaceSection. ReadFrontMatter, FindHeading, TableCount
        // and ReadTable read parsed structure and never touch a raw offset, so they are unaffected
        // and must keep returning what the reader gives them.
        var normalized = markdown.Replace("\r\n", "\n", StringComparison.Ordinal);

        var doc = ParseOrThrow(normalized);

        // The document's own top-level blocks are searched directly, rather than through
        // doc.FindHeading. FindHeading matches ANY heading, including one nested inside a
        // blockquote or a list item — and the boundary walk below indexes doc.ChildObjects using
        // the heading's IndexInParent, which for a nested heading counts within its QuoteBlock or
        // ListItem instead. Measured against OfficeIMO.Markdown 3.2.6: both `> ## Deep` and
        // `- # Nested` report IndexInParent = 0, a perfectly valid index into an unrelated list,
        // and the splice duplicates the document's tail. Nothing raises.
        //
        // Searching only top-level blocks makes a nested heading INVISIBLE here rather than merely
        // refused, which is the honest contract and also closes a bug the refusing form had: a
        // nested heading sharing text with a real top-level heading later in the document matched
        // first, and rejecting that match made the legitimate heading impossible to edit, with no
        // workaround available to a caller.
        //
        // OfficeIMO.Markdown.HeadingBlock is fully qualified, and the hazard is general rather than
        // specific to this one type: DocToolkit.Docx declares internal block types in this same
        // `DocToolkit` namespace (see DocxBlock.cs) named HeadingBlock, ParagraphBlock, TableBlock
        // and ImageBlock — every one of which also exists as a public type in OfficeIMO.Markdown —
        // and InternalsVisibleTo makes them visible here. A type in the enclosing namespace always
        // wins over one brought in by `using`, silently and with no ambiguity warning. So ANY
        // OfficeIMO.Markdown type named anywhere in this file whose name a DocToolkit.Docx block
        // type shares must be written out in full — a count of how many places do this today would
        // only go stale the next time one is added.
        OfficeIMO.Markdown.HeadingBlock? headingBlock = null;
        foreach (var child in doc.ChildObjects)
        {
            if (child is OfficeIMO.Markdown.HeadingBlock candidate
                && string.Equals(candidate.Text, headingText, comparison))
            {
                headingBlock = candidate;
                break;
            }
        }

        if (headingBlock is null)
        {
            throw new ArgumentException(
                $"No heading matching '{headingText}' was found.", nameof(headingText));
        }

        var level = headingBlock.Level;

        // The body starts where the very next block starts, which cleanly includes any blank
        // line between the heading and its content. A heading with nothing after it at all (no
        // NextSibling) has an empty body sitting at the end of the document — falling back to
        // the heading's own SourceSpan.EndOffset + 1 here is measured to be off by one and drops
        // the heading line's trailing newline.
        var bodyStart = headingBlock.NextSibling?.SourceSpan?.StartOffset ?? normalized.Length;

        // The section ends at the next heading of the SAME OR SHALLOWER level, found by walking
        // the document's flat top-level block list from just after this heading. Absent such a
        // heading, the section runs to the end of the document.
        var sectionEnd = normalized.Length;
        var siblings = doc.ChildObjects;
        var startIndex = (headingBlock.IndexInParent ?? -1) + 1;
        for (var i = startIndex; i < siblings.Count; i++)
        {
            // Fully qualified for the reason spelled out above the top-level search: an
            // unqualified `HeadingBlock` binds to DocToolkit.Docx's internal type, silently.
            if (siblings[i] is OfficeIMO.Markdown.HeadingBlock sibling && sibling.Level <= level)
            {
                sectionEnd = sibling.SourceSpan?.StartOffset ?? normalized.Length;
                break;
            }
        }

        return normalized.Substring(0, bodyStart) + newContent + normalized.Substring(sectionEnd);
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
            // Describe cannot return non-null for anything THIS class can raise, and that is
            // expected rather than an untested branch to close. Both shapes it recognises are
            // identified by a stack frame from the DOCX/PDF *conversion* path
            // (MarkdownToWordConverter.AddRun, NumberedListBlock), which a plain
            // MarkdownReader.Parse never produces — so every failure reachable from here falls
            // through to FailureMessage, and no test written against MarkdownEditor can exercise
            // the diagnosed arm. It is still called, so this class cannot drift out of step with
            // MarkdownToDocxConverter if the reader ever starts surfacing those frames itself.
            throw new DocumentConversionException(
                MarkdownFailureDiagnosis.Describe(ex, markdown) ?? FailureMessage, ex);
        }
    }
}
