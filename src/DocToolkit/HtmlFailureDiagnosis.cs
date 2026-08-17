namespace DocToolkit;

/// <summary>
/// Turns an HTML-conversion failure the package can actually identify into a message that names it.
/// </summary>
/// <remarks>
/// <b>Every message here must name a cause the exception can DISTINGUISH</b>, which is the lesson
/// this package has already had to learn twice: a timeout message that asserted "a TCP connect that
/// will never complete" as fact cost real time chasing a network regression that did not exist, and
/// <c>PdfEditor</c>'s read failure now names three candidates rather than guessing between them. So
/// the test here is a stack frame in a specific type, not the exception's type alone - an
/// <see cref="IndexOutOfRangeException"/> by itself says nothing about tables.
///
/// <b>Anything not recognised falls through to the generic wrapper</b>, deliberately. A diagnosis
/// that is right most of the time is worse than no diagnosis, because it sends the reader somewhere
/// specific and wrong.
/// </remarks>
internal static class HtmlFailureDiagnosis
{
    /// <summary>
    /// Returns a message naming the cause, or <see langword="null"/> when the failure is not one of
    /// the recognised shapes and the caller should use its generic wrapper.
    /// </summary>
    internal static string? Describe(Exception ex)
    {
        if (OverhangingRowSpan(ex)) return OverhangingRowSpanMessage;
        return null;
    }

    /// <summary>
    /// A table cell whose <c>rowspan</c> reaches past the last row of its table.
    /// </summary>
    /// <remarks>
    /// <b>Measured 2026-08-17 across 179 real .gov pages: 14 of them - 7.7% - fail this way</b>,
    /// which made it the single most frequent conversion failure in the set, and the caller was told
    /// only "see the inner exception" over a bare <see cref="IndexOutOfRangeException"/> naming no
    /// table.
    ///
    /// The parser sizes an accumulator to the rows a table actually has and then indexes it with the
    /// rowspan, unclamped, so any cell reaching past the last row writes out of bounds. The boundary
    /// is exact and was measured as a grid: one row breaks at <c>rowspan=2</c>, two rows at 3, three
    /// rows at 4, four rows survives 4. <c>colspan</c> is unaffected at every value tried, and
    /// <c>rowspan="0"</c> is fine.
    ///
    /// <b>The markup is not malformed</b> - browsers clamp such a rowspan to the rows that exist, so
    /// it renders correctly everywhere else, which is exactly why real pages carry it. The corpus
    /// held spans of 2, 3, 14, 100 and 103 against tables of one to three rows; a <c>rowspan="100"</c>
    /// is somebody wanting a cell to reach the bottom of a table and picking a number bigger than the
    /// row count.
    ///
    /// <b>The frame is the discriminator, not the exception type.</b> Matching on
    /// <see cref="IndexOutOfRangeException"/> alone would put this message on any index error
    /// anywhere in the conversion, which is precisely the over-claiming this class exists to avoid.
    /// </remarks>
    private const string OverhangingRowSpanMessage =
        "Failed to convert HTML to DOCX: a table cell has a rowspan reaching past the last row of "
        + "its table, and the HTML parser this package uses cannot read that - it raises an index "
        + "error rather than reporting it. The markup is not invalid: browsers clamp such a rowspan "
        + "to the rows that exist, so the page renders correctly in a browser. To convert it, reduce "
        + "the rowspan to at most the number of rows below the cell, or remove the attribute. See "
        + "the inner exception for the parser's own error.";

    /// <summary>The frame that identifies it. Private to the parser, so this is a heuristic on a
    /// name - but a specific one, and it fails closed: no match means no claim.</summary>
    private const string RowSpanFrame = "TableExpression.GuessColumnsCount";

    private static bool OverhangingRowSpan(Exception ex) =>
        ex is IndexOutOfRangeException
        && (ex.StackTrace ?? string.Empty).Contains(RowSpanFrame, StringComparison.Ordinal);
}
