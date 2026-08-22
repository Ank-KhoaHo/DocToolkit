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
    /// Repairs applied only when the render fails in the way each one addresses.
    /// </summary>
    /// <remarks>
    /// Each is tried <b>at most once per render</b>, which is what bounds the loop below: a repair
    /// that runs and does not help cannot be selected again, so the worst case is one attempt per
    /// entry plus the first.
    /// </remarks>
    private static readonly (Func<Exception, bool> Matches, Func<string, string> Repair)[] Retryable =
    [
        (EmptyTableCellRepair.WouldHelp, EmptyTableCellRepair.Apply),
        (ImageLinkRepair.WouldHelp, ImageLinkRepair.Apply),
    ];

    /// <summary>
    /// Renders <paramref name="html"/> to PDF, repairing and retrying when it fails in a way one of
    /// the repairs above is known to address.
    /// </summary>
    /// <param name="html">The markup to render.</param>
    /// <param name="toDocx">Converts prepared HTML to a .docx, carrying whatever options and
    /// cancellation token the calling overload was given.</param>
    /// <param name="fonts">Fonts to fall back to when rendering, or <see langword="null"/> for none.</param>
    /// <param name="ct">Observed before each attempt and between the two stages of one.</param>
    /// <remarks>
    /// <b>A retry costs a full second conversion, and only a document that already failed pays
    /// it.</b> That is the right way round: the alternative is editing every document in the hope
    /// that some of them needed it.
    /// </remarks>
    internal static async Task<byte[]> RenderAsync(
        string html, Func<string, Task<byte[]>> toDocx, PdfFontOptions? fonts = null,
        CancellationToken ct = default)
    {
        var current = Prepare(html);
        var used = new HashSet<int>();

        while (true)
        {
            // BEFORE EACH ATTEMPT, because a repair costs a whole further conversion and this
            // loop can run several. A caller who cancelled during attempt one should not pay for
            // attempt two.
            ct.ThrowIfCancellationRequested();

            try
            {
                var docx = await toDocx(current).ConfigureAwait(false);

                // BETWEEN THE STAGES, and this is the check the whole parameter exists for.
                // DocxToPdfConverter.Convert is synchronous, CPU-bound and the expensive half; it
                // takes no token and cannot be interrupted once entered. Until 2026-08-22 the
                // token reached only the HTML to DOCX stage, so every HTML to PDF overload
                // documented an OperationCanceledException it could not raise over most of its
                // own runtime.
                //
                // An already-cancelled token could never have caught that: the first stage
                // refuses immediately, so the suite went green either way. Same shape as the
                // seven PdfEditor overloads that passed the cancellation suite only because
                // destination.WriteAsync refused at the end.
                ct.ThrowIfCancellationRequested();

                return DocxToPdfConverter.Convert(docx, fonts);
            }
            catch (DocumentConversionException ex)
            {
                // A real page can need more than one of these: fixing its empty cells reveals that
                // its image links are also unlabelled. So this loops rather than retrying once,
                // taking whichever repair matches the failure actually raised.
                var index = -1;
                for (var i = 0; i < Retryable.Length; i++)
                {
                    if (used.Contains(i) || !Retryable[i].Matches(ex)) continue;
                    index = i;
                    break;
                }

                if (index < 0) throw;
                used.Add(index);

                var repaired = Retryable[index].Repair(current);
                // Nothing of that kind to fix, so re-running would fail identically. Rethrow the
                // original rather than paying for a second conversion to learn nothing.
                if (ReferenceEquals(repaired, current)) throw;

                current = repaired;
            }
        }
    }
}
