using System.Text;

namespace DocToolkit;

/// <summary>
/// Shared placeholder-substitution core for the OOXML editors.
///
/// Both Word (<c>w:t</c> inside <c>w:p</c>) and PowerPoint (<c>a:t</c> inside <c>a:p</c>) split a
/// single visible word across several text runs whenever formatting, spell-check state or a
/// language change intervenes, so a per-run <c>string.Replace</c> misses any placeholder that
/// straddles a run boundary. The naive fix — concatenate the paragraph, substitute, then dump the
/// whole result onto the first run — loses every other run's formatting and empties hyperlinks.
///
/// This does the substitution against the concatenated paragraph text but maps the resulting
/// offsets back onto the individual runs, so a run the match does not overlap is never touched at
/// all. Matching is a single left-to-right pass (longest key wins at any given offset), so a
/// substituted value is never itself rescanned for further placeholders.
/// </summary>
internal static class RunTextSplicer
{
    private readonly record struct Match(int Start, int End, string Value);

    /// <summary>
    /// Applies <paramref name="replacements"/> across the concatenation of <paramref name="nodes"/>,
    /// writing back only to the nodes a match actually overlaps.
    /// </summary>
    /// <returns><c>true</c> when at least one node's text changed.</returns>
    public static bool Apply<T>(
        IReadOnlyList<T> nodes,
        Func<T, string> read,
        Action<T, string> write,
        IReadOnlyDictionary<string, string> replacements)
    {
        if (nodes.Count == 0) return false;

        var originals = new string[nodes.Count];
        var builder = new StringBuilder();
        for (var i = 0; i < nodes.Count; i++)
        {
            originals[i] = read(nodes[i]) ?? string.Empty;
            builder.Append(originals[i]);
        }

        var merged = builder.ToString();
        if (merged.Length == 0) return false;

        var matches = FindMatches(merged, replacements);
        if (matches.Count == 0) return false;

        var changed = false;
        var matchIndex = 0;
        var nodeStart = 0;

        for (var i = 0; i < nodes.Count; i++)
        {
            var nodeEnd = nodeStart + originals[i].Length;
            var rebuilt = new StringBuilder();
            var pos = nodeStart;

            while (pos < nodeEnd)
            {
                if (matchIndex < matches.Count && matches[matchIndex].Start <= pos)
                {
                    var match = matches[matchIndex];

                    // The replacement value is inserted into the run that owns the *start* of the
                    // match, so it inherits that run's formatting. Runs the match merely spans
                    // simply lose their share of the matched characters.
                    if (match.Start == pos) rebuilt.Append(match.Value);

                    pos = Math.Min(match.End, nodeEnd);
                    if (match.End <= nodeEnd) matchIndex++;
                    continue;
                }

                var copyEnd = matchIndex < matches.Count
                    ? Math.Min(nodeEnd, matches[matchIndex].Start)
                    : nodeEnd;
                rebuilt.Append(merged, pos, copyEnd - pos);
                pos = copyEnd;
            }

            var updated = rebuilt.ToString();
            if (!string.Equals(updated, originals[i], StringComparison.Ordinal))
            {
                write(nodes[i], updated);
                changed = true;
            }

            nodeStart = nodeEnd;
        }

        return changed;
    }

    /// <summary>
    /// Non-overlapping, left-to-right matches of any key in <paramref name="replacements"/>.
    /// At a given offset the longest matching key wins, which makes the outcome independent of
    /// dictionary enumeration order.
    /// </summary>
    private static List<Match> FindMatches(string merged, IReadOnlyDictionary<string, string> replacements)
    {
        var keys = replacements.Keys
            .Where(k => !string.IsNullOrEmpty(k))
            .OrderByDescending(k => k.Length)
            .ThenBy(k => k, StringComparer.Ordinal)
            .ToArray();

        var matches = new List<Match>();
        if (keys.Length == 0) return matches;

        var i = 0;
        while (i < merged.Length)
        {
            string? hit = null;
            foreach (var key in keys)
            {
                if (key.Length <= merged.Length - i &&
                    merged.AsSpan(i).StartsWith(key.AsSpan(), StringComparison.Ordinal))
                {
                    hit = key;
                    break;
                }
            }

            if (hit is null)
            {
                i++;
                continue;
            }

            matches.Add(new Match(i, i + hit.Length, replacements[hit] ?? string.Empty));
            i += hit.Length;
        }

        return matches;
    }
}
