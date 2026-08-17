namespace DocToolkit;

/// <summary>
/// Turns the two Markdown conversion failures this package can identify into messages that name
/// them.
/// </summary>
/// <remarks>
/// <b>Measured against the CommonMark 0.31.2 conformance suite - 652 examples - plus four real
/// project READMEs.</b> Markdown is the best-performing capability here (655/656 to DOCX, 645/656 to
/// PDF), and the two things it does reject are both spec-valid input that arrives as an unhandled
/// exception from inside a dependency: a bare <see cref="NullReferenceException"/> and an
/// <see cref="ArgumentOutOfRangeException"/>, neither naming a construct.
///
/// <b>Neither is repaired, and that is deliberate rather than unfinished.</b> Both repairs would
/// change what the document says:
///
/// <list type="bullet">
/// <item><description><c>&amp;#10;</c> is a line feed the author wrote as a character. Turning it
/// into a real newline is not equivalent - two of them in a row are a paragraph break in source and
/// are not as entities - and turning it into a space deletes a character somebody
/// wrote.</description></item>
/// <item><description>An ordered list starting at <c>0</c> renumbers if it is made to start at
/// <c>1</c>. CommonMark permits the zero, and the numbers are the content.</description></item>
/// </list>
///
/// So what improves is the diagnosis, on the same discipline as
/// <see cref="HtmlFailureDiagnosis"/>: name only a cause that can be told apart, and let everything
/// else keep the generic wrapper.
/// </remarks>
internal static class MarkdownFailureDiagnosis
{
    /// <summary>
    /// Returns a message naming the cause, or <see langword="null"/> when the failure is not one of
    /// the recognised shapes.
    /// </summary>
    /// <param name="ex">The exception the conversion raised.</param>
    /// <param name="markdown">
    /// The source. <b>Required, because the frame alone does not identify the line-feed case</b> -
    /// <c>AddRun</c> is where every inline run is built and could fail for reasons having nothing to
    /// do with character references. The frame narrows it; the input confirms it.
    /// </param>
    internal static string? Describe(Exception ex, string markdown)
    {
        if (LineFeedEntity(ex, markdown)) return LineFeedMessage;
        if (ListStartingBelowOne(ex)) return ListStartMessage;
        return null;
    }

    /// <summary>
    /// A numeric character reference for a line feed, <c>&amp;#10;</c>, in inline text.
    /// </summary>
    /// <remarks>
    /// Measured 2026-08-17, and the boundary is narrow: <c>&amp;#9;</c> (tab), <c>&amp;#13;</c>
    /// (carriage return), <c>&amp;#32;</c> (space) and <c>&amp;#0;</c> all convert, and a REAL
    /// newline in the source converts. Only the entity form of U+000A crashes, in both converters,
    /// at the same frame.
    /// </remarks>
    private static bool LineFeedEntity(Exception ex, string markdown) =>
        ex is NullReferenceException
        && (ex.StackTrace ?? string.Empty).Contains("MarkdownToWordConverter.AddRun", StringComparison.Ordinal)
        && (markdown.Contains("&#10;", StringComparison.OrdinalIgnoreCase)
            || markdown.Contains("&#x0a;", StringComparison.OrdinalIgnoreCase)
            || markdown.Contains("&#xa;", StringComparison.OrdinalIgnoreCase));

    private const string LineFeedMessage =
        "Failed to convert Markdown: it contains a numeric character reference for a line feed "
        + "(&#10;), and the Markdown reader this package uses raises a null reference on it. The "
        + "input is valid CommonMark. To convert it, write the line break directly instead of as a "
        + "character reference - a real newline converts. Other references are unaffected: &#9;, "
        + "&#13;, &#32; and &#0; all convert. See the inner exception for the reader's own error.";

    /// <summary>
    /// An ordered list whose first number is below 1.
    /// </summary>
    /// <remarks>
    /// <b>The PDF path only.</b> Measured: <c>0. ok</c> and <c>0) ok</c> convert to DOCX perfectly
    /// well and are refused by the PDF renderer's numbered-list block, so a caller who only ever
    /// converts to DOCX will never meet this.
    /// </remarks>
    private static bool ListStartingBelowOne(Exception ex) =>
        ex is ArgumentOutOfRangeException
        && (ex.StackTrace ?? string.Empty).Contains("NumberedListBlock", StringComparison.Ordinal);

    private const string ListStartMessage =
        "Failed to convert Markdown to PDF: it contains an ordered list starting below 1 (for "
        + "example \"0. item\"), and the PDF renderer this package uses requires a start of 1 or "
        + "more. The input is valid CommonMark, which permits any starting number. Converting to "
        + "DOCX works - it is only the PDF stage that refuses - or renumber the list from 1. See "
        + "the inner exception for the renderer's own error.";
}
