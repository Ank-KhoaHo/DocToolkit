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
    /// <remarks>
    /// <b>A retry costs a full second conversion, and only a document that already failed pays
    /// it.</b> That is the right way round: the alternative is editing every document in the hope
    /// that some of them needed it.
    /// </remarks>
    internal static async Task<byte[]> RenderAsync(string html, Func<string, Task<byte[]>> toDocx)
    {
        var current = Prepare(html);
        var used = new HashSet<int>();

        while (true)
        {
            try
            {
                return DocxToPdfConverter.Convert(await toDocx(current).ConfigureAwait(false));
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
