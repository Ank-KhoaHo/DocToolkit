namespace DocToolkit;

/// <summary>
/// Names the two most common reasons a real Word document cannot be rendered to PDF.
/// </summary>
/// <remarks>
/// <b>Measured 2026-08-18 over 99 documents carrying real content</b> - govdocs1 predates
/// <c>.docx</c>, so these are the corpus's 111 genuine <c>.doc</c> files converted first.
/// <b>DOCX to PDF succeeded on 71.7% of them</b>, and 15 of the 28 failures are the two causes named
/// here: 8 for a negative paragraph indent and 7 for header or footer content wider than the page.
///
/// <b>Both are legal in Word, which is the part worth telling somebody.</b> The renderer's own
/// messages are accurate - <i>"Paragraph right indent must be a non-negative finite value"</i> - and
/// leave a reader looking for a mistake in a document that does not contain one. Content set outside
/// the margin, and a wide header, are ordinary things to find in a letterhead or a report.
///
/// <b>Nothing here is repaired.</b> Clamping a negative indent to zero pulls content back inside a
/// margin the author deliberately put it outside of, and shrinking a header changes a layout
/// somebody chose - unlike the HTML repairs, where a browser's own behaviour said what the right
/// answer was. Those are decisions rather than measurements, and they are filed rather than taken.
///
/// <b>An ordinary hanging indent is unaffected</b>, which was measured before this message was
/// written so it could say so: <c>w:hanging</c> and a negative <c>w:firstLine</c> both convert. Only
/// a negative <c>w:left</c> or <c>w:right</c> is refused, at any magnitude.
/// </remarks>
internal static class DocxPdfFailureDiagnosis
{
    /// <summary>
    /// Returns a message naming the cause, or <see langword="null"/> when the failure is not one of
    /// the recognised shapes and the caller should use its generic wrapper.
    /// </summary>
    /// <remarks>
    /// Matched on the renderer's own message rather than on the exception type, which says nothing:
    /// these arrive as <see cref="ArgumentException"/>, as does a great deal else. Same discipline as
    /// <see cref="HtmlFailureDiagnosis"/> - name only a cause that can be told apart, and let
    /// anything unrecognised keep the generic wrapper.
    /// </remarks>
    internal static string? Describe(Exception ex)
    {
        var message = ex.Message ?? string.Empty;

        if (message.Contains("indent must be a non-negative", StringComparison.OrdinalIgnoreCase))
            return NegativeIndentMessage;

        if (message.Contains("must fit inside the page content width", StringComparison.OrdinalIgnoreCase)
            || message.Contains("header zones must not overlap", StringComparison.OrdinalIgnoreCase))
            return HeaderFooterMessage;

        return null;
    }

    private const string NegativeIndentMessage =
        "Failed to convert DOCX to PDF: a paragraph has a negative left or right indent, and the "
        + "PDF renderer this package uses requires both to be zero or more. The document is not "
        + "invalid - Word allows a paragraph to be set outside the margin, which is how a letterhead "
        + "or a pull-quote is often laid out. An ordinary hanging indent is unaffected and converts "
        + "normally. To render this document, set the negative indent to 0. See the inner exception "
        + "for the renderer's own error.";

    private const string HeaderFooterMessage =
        "Failed to convert DOCX to PDF: the header or footer is wider than the page content area, or "
        + "its zones overlap, and the PDF renderer this package uses refuses that rather than "
        + "clipping it. Word lays such a header out without complaint, so the document is not "
        + "invalid. To render it, narrow the header or footer content, or widen the page margins. "
        + "See the inner exception for the renderer's own error.";
}
