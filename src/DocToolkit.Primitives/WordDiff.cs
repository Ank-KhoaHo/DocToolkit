namespace DocToolkit;

/// <summary>What happened to one run of words between two documents.</summary>
internal enum WordDiffKind
{
    /// <summary>Present in both, unchanged.</summary>
    Same = 0,

    /// <summary>Present only in the revised document.</summary>
    Inserted = 1,

    /// <summary>Present only in the original.</summary>
    Deleted = 2,
}

/// <summary>One contiguous stretch of words with the same fate.</summary>
internal readonly record struct WordDiffSpan(WordDiffKind Kind, IReadOnlyList<string> Words);

/// <summary>
/// The word-sequence diff behind document comparison (A118), with no OOXML in it at all.
/// </summary>
/// <remarks>
/// <b>Kept free of OpenXml deliberately.</b> A diff is not document-specific, and this is the half
/// most likely to be wrong in ways a document-level test cannot localise — so it is tested directly
/// against string sequences, where a failure names the input rather than a .docx. That paid
/// immediately: see the note on the algorithm below.
///
/// <para>
/// <b>Common prefix and suffix are stripped first, then a longest-common-subsequence table solves
/// what is left.</b> The stripping is what makes this practical rather than an optimisation
/// detail: people compare two revisions of ONE document, so the differing middle is a small
/// fraction of the whole, and the table is O(n·m) over that middle rather than over the documents.
/// Two entirely unrelated documents are the worst case and are bounded — see
/// <see cref="MaxTableCells"/>.
/// </para>
///
/// <para>
/// <b>Myers' O(ND) algorithm was written first and was WRONG, and the property test caught it.</b>
/// `DroppingEachSideRebuildsTheOtherDocumentExactly` failed on three inputs including generated
/// ones — words were being lost, which every span list still looked plausible under. Myers is the
/// better algorithm on paper and the middle-snake recursion is easy to get subtly wrong; a
/// correct diff whose cost is bounded beats a faster one that loses a word, and nothing here
/// measured a case where the difference mattered.
/// </para>
/// </remarks>
internal static class WordDiff
{
    /// <summary>
    /// The largest LCS table this will build, past which the two texts are reported as a wholesale
    /// replacement instead.
    /// </summary>
    /// <remarks>
    /// <b>The fallback is a correct answer, not a refusal</b> — deleting everything and inserting
    /// everything genuinely describes the difference, it is merely coarse. That is the right
    /// trade for two documents with nothing in common, which is the only way to reach this bound
    /// after the common prefix and suffix have been removed: 4 million cells is roughly 2,000
    /// differing words on each side.
    /// </remarks>
    private const int MaxTableCells = 4_000_000;

    /// <summary>
    /// Splits text into the words a diff compares, keeping the whitespace that follows each.
    /// </summary>
    /// <remarks>
    /// <b>Trailing whitespace rides with its word</b>, so re-joining the pieces reproduces the
    /// original exactly. A split that dropped it would make every reconstruction lossy, and the
    /// loss would show up as spurious differences rather than as an error.
    /// </remarks>
    public static List<string> Split(string text)
    {
        var words = new List<string>();
        var i = 0;

        while (i < text.Length)
        {
            var start = i;
            while (i < text.Length && !char.IsWhiteSpace(text[i])) i++;
            while (i < text.Length && char.IsWhiteSpace(text[i])) i++;
            words.Add(text[start..i]);
        }

        return words;
    }

    /// <summary>
    /// The edits turning <paramref name="original"/> into <paramref name="revised"/>, in order.
    /// </summary>
    /// <remarks>
    /// Adjacent words sharing a fate are merged into one span, so a caller emits one
    /// <c>w:ins</c> per inserted phrase rather than one per word.
    /// </remarks>
    public static List<WordDiffSpan> Diff(IReadOnlyList<string> original, IReadOnlyList<string> revised)
    {
        var steps = new List<(WordDiffKind Kind, string Word)>();

        var start = 0;
        while (start < original.Count && start < revised.Count && original[start] == revised[start])
        {
            steps.Add((WordDiffKind.Same, original[start]));
            start++;
        }

        var endA = original.Count;
        var endB = revised.Count;
        while (endA > start && endB > start && original[endA - 1] == revised[endB - 1])
        {
            endA--;
            endB--;
        }

        Middle(original, revised, start, endA, start, endB, steps);

        for (var i = endA; i < original.Count; i++) steps.Add((WordDiffKind.Same, original[i]));

        var spans = new List<WordDiffSpan>();
        foreach (var step in steps)
        {
            if (spans.Count > 0 && spans[^1].Kind == step.Kind)
            {
                ((List<string>)spans[^1].Words).Add(step.Word);
                continue;
            }

            spans.Add(new WordDiffSpan(step.Kind, new List<string> { step.Word }));
        }

        return spans;
    }

    /// <summary>
    /// Diffs the differing middle, once the shared prefix and suffix are gone.
    /// </summary>
    private static void Middle(
        IReadOnlyList<string> a, IReadOnlyList<string> b,
        int aStart, int aEnd, int bStart, int bEnd,
        List<(WordDiffKind, string)> steps)
    {
        var n = aEnd - aStart;
        var m = bEnd - bStart;

        if (n == 0)
        {
            for (var i = bStart; i < bEnd; i++) steps.Add((WordDiffKind.Inserted, b[i]));
            return;
        }

        if (m == 0)
        {
            for (var i = aStart; i < aEnd; i++) steps.Add((WordDiffKind.Deleted, a[i]));
            return;
        }

        if ((long)n * m > MaxTableCells)
        {
            // Nothing in common worth finding. Deleting all and inserting all is correct, just
            // coarse - see MaxTableCells.
            for (var i = aStart; i < aEnd; i++) steps.Add((WordDiffKind.Deleted, a[i]));
            for (var i = bStart; i < bEnd; i++) steps.Add((WordDiffKind.Inserted, b[i]));
            return;
        }

        // lcs[i, j] is the length of the longest common subsequence of a[aStart+i..] and
        // b[bStart+j..], filled from the end so the walk below can go forwards and emit in
        // document order.
        var lcs = new int[n + 1, m + 1];
        for (var i = n - 1; i >= 0; i--)
        {
            for (var j = m - 1; j >= 0; j--)
            {
                lcs[i, j] = a[aStart + i] == b[bStart + j]
                    ? lcs[i + 1, j + 1] + 1
                    : Math.Max(lcs[i + 1, j], lcs[i, j + 1]);
            }
        }

        var x = 0;
        var y = 0;
        while (x < n && y < m)
        {
            if (a[aStart + x] == b[bStart + y])
            {
                steps.Add((WordDiffKind.Same, a[aStart + x]));
                x++;
                y++;
            }
            else if (lcs[x + 1, y] >= lcs[x, y + 1])
            {
                steps.Add((WordDiffKind.Deleted, a[aStart + x]));
                x++;
            }
            else
            {
                steps.Add((WordDiffKind.Inserted, b[bStart + y]));
                y++;
            }
        }

        for (; x < n; x++) steps.Add((WordDiffKind.Deleted, a[aStart + x]));
        for (; y < m; y++) steps.Add((WordDiffKind.Inserted, b[bStart + y]));
    }
}
