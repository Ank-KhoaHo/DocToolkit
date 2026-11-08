namespace DocToolkit;

/// <summary>
/// Names the reasons a real Word document cannot be rendered to PDF that a reader can act on.
/// </summary>
/// <remarks>
/// <b>Measured 2026-08-18 over 99 documents carrying real content</b> - govdocs1 predates
/// <c>.docx</c>, so these are the corpus's 111 genuine <c>.doc</c> files converted first, on
/// OfficeIMO 3.2.6. <b>DOCX to PDF succeeded on 71.7% of them</b>, and 15 of the 28 failures were a
/// negative paragraph indent (8) or header or footer content wider than the page (7).
///
/// <b>Both were legal in Word, which is the part worth telling somebody.</b> The renderer's own
/// messages were accurate and left a reader looking for a mistake in a document that does not
/// contain one. <b>OfficeIMO 3.5 renders a negative indent</b> (measured: 300+ negative right
/// indents in the corpus convert), so that message and its retry are gone; the header and footer
/// message remains.
///
/// <b>Nothing here is repaired by this class.</b> It only describes. Shrinking a header changes a
/// layout somebody chose - unlike the HTML repairs, where a browser's own behaviour said what the
/// right answer was. Those are decisions rather than measurements, and they are filed rather than
/// taken. The refusals that do have a repair are answered by <see cref="DocxPdfRefusalLadder"/>.
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

        if (message.Contains("must fit inside the page content width", StringComparison.OrdinalIgnoreCase)
            || message.Contains("header zones must not overlap", StringComparison.OrdinalIgnoreCase))
            return HeaderFooterMessage;

        if (message.Contains("not covered by any embedded font", StringComparison.OrdinalIgnoreCase))
            return MissingGlyphMessage;

        return null;
    }

    /// <summary>
    /// A character no font on this machine can encode.
    /// </summary>
    /// <remarks>
    /// <b>The renderer's own message is good - it names the character and the fallbacks it tried -
    /// and stops one step short of the useful part.</b> It says "add a fallback font" without saying
    /// that this package now takes one, so a reader hits a wall the API can already get them past.
    ///
    /// <b>It also does not say that this is host-dependent</b>, which is the thing most likely to
    /// mislead: the same document renders on a machine whose fonts happen to cover the script, so
    /// somebody testing on Linux and deploying to Windows meets it only in production.
    /// </remarks>
    private const string MissingGlyphMessage =
        "Failed to convert DOCX to PDF: the document contains a character no font available to the "
        + "renderer can encode - typically non-Latin text such as Cyrillic, Greek or CJK. "
        + "IMPORTANT: this depends on the machine. The same document renders on a host whose "
        + "installed fonts cover the script, so it can pass in development and fail in production. "
        + "To remove the machine from the answer, pass a PdfFontOptions carrying fonts that cover "
        + "the text - the bytes come from you, so nothing is fetched. Supply fonts covering "
        + "EVERYTHING the document uses, not only the script that failed: the fonts you pass "
        + "replace the host's own fallbacks rather than adding to them, so too few is worse than "
        + "none. See the inner exception for the character and the fallbacks that were tried.";

    private const string HeaderFooterMessage =
        "Failed to convert DOCX to PDF: the header or footer is wider than the page content area, or "
        + "its zones overlap, and the PDF renderer this package uses refuses that rather than "
        + "clipping it. Word lays such a header out without complaint, so the document is not "
        + "invalid. To render it, narrow the header or footer content, or widen the page margins. "
        + "See the inner exception for the renderer's own error.";
}
