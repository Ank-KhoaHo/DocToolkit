namespace DocToolkit;

/// <summary>
/// The repairs that run only AFTER OfficeIMO 3.5 has refused a DOCX with their specific message.
/// </summary>
/// <remarks>
/// <b>Why they wait for the refusal.</b> Each step below changes a layout Word accepts: a column
/// width, a cell padding, an indent, a list marker's tab. A document 3.5 renders as it stands has
/// nothing to fix, so repairing it up front would only change a rendering that was already right.
/// Waiting costs a second render for the documents that need it, and nothing for the rest.
///
/// <b>One refusal can need several steps, and one document several refusals.</b> A step runs at
/// most once per conversion (<c>used</c>) and a step that finds nothing to change is skipped, so the
/// ladder always ends: at most one render per step. html-000079 needs padding zeroed first and the
/// tiny percentage width repaired second, because 3.5 only reaches the second refusal once the first
/// is out of the way.
///
/// <b>Steps for the same message run in order, and the first that changes something wins the
/// round.</b> The empty-cell width comes before the zero padding because it keeps the document's
/// padding when it is enough on its own (measured: 45 documents that zero padding rescued keep their
/// padding that way).
///
/// <b>Matching is on the renderer's message, anywhere in the exception chain</b>, for the reason
/// <see cref="DocxPdfFailureDiagnosis"/> gives: these all arrive as <see cref="ArgumentException"/>.
/// </remarks>
internal static class DocxPdfRefusalLadder
{
    private sealed record Step(string Message, Func<byte[], byte[]> Repair);

    // Each Repair returns the SAME array when it has nothing to change (the DocxPartRewrite
    // contract), which NextOrNull turns into "skip this step".
    private static readonly Step[] Steps =
    [
        // R1 and R6 come first: they answer refusals that stop the render before layout reaches the
        // padding checks, and both look at the whole document because they run only on a refusal.
        new("cannot be empty or whitespace", HyperlinkBlankRunSplit.Apply),
        new("header offset must not exceed the top margin", HeaderOffsetRepair.Apply),
        new("preferred width must be a positive finite value", TinyPercentWidthRepair.Apply),
        new("cell padding must leave a positive text width", EmptyCellMinWidth.Apply),
        new("cell padding must leave a positive text width", CellPaddingZero.Apply),
        new("indents must leave a positive text width", IndentFitToColumn.Apply),
        new("Inline element width exceeds the available paragraph width", ListMarkerSuffixSpace.Apply),
    ];

    /// <summary>Whether some step answers this failure (or one inside it).</summary>
    internal static bool Knows(Exception ex) => Steps.Any(s => Mentions(ex, s.Message));

    /// <summary>
    /// The package with the first unused matching step applied, or <see langword="null"/> when no
    /// unused step matches <paramref name="failure"/> or every matching one finds nothing to change.
    /// </summary>
    /// <param name="failure">What the renderer just refused with.</param>
    /// <param name="current">The package that was just rendered. Never modified.</param>
    /// <param name="used">Step numbers already tried in this conversion; updated here.</param>
    internal static byte[]? NextOrNull(Exception failure, byte[] current, ISet<int> used)
    {
        for (var i = 0; i < Steps.Length; i++)
        {
            if (used.Contains(i) || !Mentions(failure, Steps[i].Message)) continue;
            used.Add(i);

            var repaired = Steps[i].Repair(current);
            if (!ReferenceEquals(repaired, current)) return repaired;
        }

        return null;
    }

    private static bool Mentions(Exception ex, string text)
    {
        for (var e = ex; e is not null; e = e.InnerException)
            if (e.Message.Contains(text, StringComparison.Ordinal)) return true;

        return false;
    }
}
