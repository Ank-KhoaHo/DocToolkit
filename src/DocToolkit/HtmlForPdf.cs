namespace DocToolkit;

/// <summary>
/// The one place HTML is prepared for, and rendered on, the PDF path.
/// </summary>
/// <remarks>
/// <b>There are four call sites and nothing structural keeps them in step</b>, which is the same
/// hand-maintained-inventory hazard that let all eight <c>PdfEditor</c> stream overloads go
/// unguarded. Routing them through one method means a repair added here reaches every overload, and
/// a test exercising every public entry point catches a site that was missed. Adding a preparation
/// call directly at a call site is how that guarantee gets lost - add it here instead.
///
/// <b>Repairs come in two kinds, and the difference is what keeps this safe.</b>
///
/// <list type="number">
/// <item><description><b>Applied up front</b>, when every document they touch is already known to
/// fail. <see cref="HtmlAnchorRepair"/> qualifies: it changes nothing unless a link resolves
/// nowhere, and such a document always fails to render.</description></item>
/// <item><description><b>Applied only after a failure</b>, when the repair cannot tell in advance
/// whether the document needed it. <see cref="EmptyTableCellRepair"/> is this: a table cell with no
/// text is completely ordinary and usually renders perfectly well, so filling every one of them up
/// front would edit documents that were fine.</description></item>
/// </list>
///
/// The second kind is what <see cref="RenderAsync"/> exists for. **A document that renders on the
/// first attempt is never repaired at all**, so no working conversion can change - which is the
/// property the whole PDF-path repair strategy rests on, held here by construction rather than by
/// argument.
///
/// <b>Not applied to the DOCX path.</b> These repairs would improve a .docx too - a dangling
/// internal link is dangling in Word as well - but a document that converts today would start
/// carrying bookmarks and cell content it did not carry before. That is a separate decision.
/// </remarks>
internal static class HtmlForPdf
{
    /// <summary>Repairs that are safe to apply before knowing whether the document would fail.</summary>
    internal static string Prepare(string html) => HtmlAnchorRepair.Apply(html);

    /// <summary>
    /// Renders <paramref name="html"/> to PDF, retrying once with a repair if it fails in a way that
    /// repair is known to address.
    /// </summary>
    /// <param name="html">The markup to render.</param>
    /// <param name="toDocx">Converts prepared HTML to a .docx, carrying whatever options and
    /// cancellation token the calling overload was given.</param>
    /// <remarks>
    /// <b>The retry costs a second full conversion, and only a document that already failed pays
    /// it.</b> That is the right way round: the alternative is editing every document in the hope
    /// that some of them needed it.
    ///
    /// A repair that returns its input unchanged means there was nothing of its kind to fix, so the
    /// original failure is rethrown rather than re-running an identical conversion to fail again.
    /// </remarks>
    internal static async Task<byte[]> RenderAsync(string html, Func<string, Task<byte[]>> toDocx)
    {
        var prepared = Prepare(html);

        try
        {
            return DocxToPdfConverter.Convert(await toDocx(prepared).ConfigureAwait(false));
        }
        catch (DocumentConversionException ex) when (EmptyTableCellRepair.WouldHelp(ex))
        {
            var repaired = EmptyTableCellRepair.Apply(prepared);
            if (ReferenceEquals(repaired, prepared)) throw;

            return DocxToPdfConverter.Convert(await toDocx(repaired).ConfigureAwait(false));
        }
    }
}
