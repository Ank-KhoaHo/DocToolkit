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

        // The open span is held as its own List rather than reached back through spans[^1].Words,
        // which would mean casting an IReadOnlyList down to the concrete type it happens to be.
        // That cast is only sound while every span in this method was built here - a fragility with
        // no upside, since the span already holds a reference to this same list.
        var spans = new List<WordDiffSpan>();
        List<string>? open = null;
        var openKind = default(WordDiffKind);

        foreach (var step in steps)
        {
            if (open is not null && openKind == step.Kind)
            {
                open.Add(step.Word);
                continue;
            }

            open = [step.Word];
            openKind = step.Kind;
            spans.Add(new WordDiffSpan(openKind, open));
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

        Walk(a, b, aStart, n, bStart, m, Table(a, b, aStart, n, bStart, m), steps);
    }

    /// <summary>
    /// The longest-common-subsequence lengths for the window, filled from the END.
    /// </summary>
    /// <remarks>
    /// <c>table[i, j]</c> is the length of the longest common subsequence of the windows starting
    /// at <c>i</c> and <c>j</c>. Filling backwards is what lets <see cref="Walk"/> run FORWARDS and
    /// therefore emit in document order - a table filled from the front has to be walked in reverse
    /// and the result reversed again, which is one more place to get an off-by-one wrong.
    /// </remarks>
    private static int[,] Table(
        IReadOnlyList<string> a, IReadOnlyList<string> b, int aStart, int n, int bStart, int m)
    {
        var table = new int[n + 1, m + 1];

        for (var i = n - 1; i >= 0; i--)
        {
            for (var j = m - 1; j >= 0; j--)
            {
                table[i, j] = a[aStart + i] == b[bStart + j]
                    ? table[i + 1, j + 1] + 1
                    : Math.Max(table[i + 1, j], table[i, j + 1]);
            }
        }

        return table;
    }

    /// <summary>
    /// Reads the table forwards, emitting one step per word.
    /// </summary>
    /// <remarks>
    /// The tie-break is <c>&gt;=</c> rather than <c>&gt;</c>, so a word replaced by another is
    /// reported deleted-then-inserted rather than the other way round. Either is a correct diff;
    /// this order is the one <c>DocxCompare</c> relies on to place <c>w:del</c> before <c>w:ins</c>.
    /// </remarks>
    private static void Walk(
        IReadOnlyList<string> a, IReadOnlyList<string> b, int aStart, int n, int bStart, int m,
        int[,] table, List<(WordDiffKind, string)> steps)
    {
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
            else if (table[x + 1, y] >= table[x, y + 1])
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
