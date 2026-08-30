using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit;

/// <summary>
/// Checks a Word document <b>before</b> converting it, and reports what
/// <see cref="DocxToPdfConverter"/> may not carry into the PDF.
/// </summary>
/// <remarks>
/// <b>Why this exists.</b> The renderer beneath <see cref="DocxToPdfConverter"/> produces no report
/// of its own, so a conversion that drops something drops it silently and returns a valid PDF. A
/// caller converting documents they did not author has no way to tell which ones need a human to
/// look at them. This answers that, and only that.
///
/// <b>It reads the SOURCE and reports presence, never loss.</b> The conversion has not run. A
/// finding means "your document contains this, and this renderer may not represent it" — not "this
/// was removed". The weaker claim is the one that can be honestly made from the input alone, and it
/// is what a caller triaging a batch actually needs.
///
/// <b>What it does NOT cover, stated because an empty report otherwise reads as a clean bill of
/// health.</b> Charts, SmartArt, embedded objects and shape effects are all plausible risks and are
/// <b>not</b> detected: authoring a fixture for each is substantial work, and this library does not
/// list a construct it has not watched fail. Detection is also limited to the main document body —
/// a nested table inside a header or footer is not reported, because that case is unmeasured.
///
/// <b>It lives in DocToolkit rather than DocToolkit.Docx, beside the converter it is about.</b>
/// <see cref="DocxReview"/> is in the Docx project because it reads a document; this is about a
/// CONVERSION, and the conversions stayed in the core project when the per-concern split ran. The
/// dependency direction settles it either way - core references Docx, not the reverse, so a
/// <c>cref</c> to <see cref="DocxToPdfConverter"/> cannot resolve from there at all.
///
/// <b>What is deliberately absent is decided by measurement, because a report that fires on
/// constructs the renderer handles teaches a caller to ignore it.</b> A text box
/// (<c>w:txbxContent</c>) renders, and so does a content control (<c>w:sdt</c>) at body level — so
/// neither is reported there, and a control inside a text box's table is not reported either.
///
/// <b>A content control in a TABLE is the exception, and it was found by re-measuring.</b> The same
/// construct that survives at body level loses its text inside a cell, or wrapping a cell or a row.
/// This class excluded content controls outright until 2026-08-27, on the body-level evidence
/// alone — correct for what had been measured, and incomplete.
/// </remarks>
public static class DocxToPdfPreflight
{
    /// <summary>Reports what <paramref name="docx"/> contains that may not reach a PDF.</summary>
    /// <param name="docx">The document that is about to be converted.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The document could not be opened or read.</exception>
    public static DocxToPdfPreflightReport Inspect(byte[] docx)
    {
        ArgumentNullException.ThrowIfNull(docx);
        if (docx.Length == 0) throw new ArgumentException(EmptySource, nameof(docx));

        using var source = new MemoryStream(docx, writable: false);
        return InspectCore(source);
    }

    /// <summary>
    /// Reports what the .docx in <paramref name="source"/> contains that may not reach a PDF.
    /// <paramref name="source"/> is <b>read</b> to its end and is neither disposed, closed nor
    /// sought.
    /// </summary>
    /// <inheritdoc cref="Inspect(byte[])" path="/remarks"/>
    /// <param name="source">The document that is about to be converted.</param>
    /// <param name="ct">Cancels before the document is read.</param>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be opened or read.</exception>
    public static async Task<DocxToPdfPreflightReport> InspectAsync(
        Stream source, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, EmptySource, nameof(source), ReadFailure, ct)
            .ConfigureAwait(false);

        return InspectCore(docx);
    }

    private const string EmptySource = "DOCX content was empty.";
    private const string ReadFailure =
        "Failed to inspect the document. See the inner exception for details.";

    /// <summary>Stable codes, so a caller can branch on one without matching prose.</summary>
    internal const string FootnoteCode = "Footnote";
    internal const string NestedTableCode = "NestedTable";
    internal const string ControlInCellCode = "ControlInCell";

    private static DocxToPdfPreflightReport InspectCore(Stream source)
    {
        try
        {
            using var doc = WordprocessingDocument.Open(source, false);
            var main = doc.MainDocumentPart;
            var body = main?.Document?.Body;
            var findings = new List<DocxToPdfPreflightFinding>();

            // Read the PACKAGE, never ExtractText. That method has its own blind spots - measured
            // 2026-08-25, it returns nothing for content-control text that IS in the PDF - and a
            // detector built on it would inherit them.

            // Measured (A94): survival is decided entirely by the BODY's own FootnoteReference
            // run, never by anything in the footnote's own definition inside FootnotesPart. A
            // reference is lost from the render specifically when its run lacks
            // RunStyle="FootnoteReference" - confirmed against OfficeIMO.Word's own footnote
            // authoring (which always applies it) and against DocxEditor.AddFootnote's real
            // output (same). No id filter is needed here, unlike the old definition-scanning
            // version: the separator/continuation-separator entries a FootnotesPart always
            // carries are never themselves referenced from the body, so they can never appear in
            // this scan regardless of id.
            int footnotes = body is null ? 0 : body.Descendants<FootnoteReference>()
                .Count(r => (r.Parent as Run)?.RunProperties?.RunStyle?.Val?.Value != "FootnoteReference");

            if (footnotes > 0)
            {
                findings.Add(new DocxToPdfPreflightFinding(
                    FootnoteCode,
                    "Footnotes",
                    $"{footnotes} footnote reference(s) without the character style a "
                    + "normally-authored footnote carries. Their text does not reach the PDF - "
                    + "measured. Word itself, and this library's own AddFootnote, always apply "
                    + "that style, so this fires only for a footnote built by hand or by another "
                    + "tool without it.",
                    footnotes,
                    DocxToPdfRisk.Known));
            }

            // A table that is a DIRECT child of a cell, at any depth of nesting. Descendants finds
            // the cells at every level; ContentControls.Tables keeps the table bound to the cell
            // that holds it - so a sibling table further down the body is not counted as nested -
            // while still seeing one wrapped in a w:sdt.
            //
            // It read c.Elements<Table>() until 2026-08-27, which missed a wrapped one entirely:
            // measured, the same document reported 1 nested table unwrapped and 0 wrapped. A
            // preflight that under-reports is worse than one that does not run, because a caller
            // reads silence as "nothing will be lost".
            int nested = body is null ? 0 : Cells(body).SelectMany(ContentControls.Tables).Count();

            if (nested > 0)
            {
                findings.Add(new DocxToPdfPreflightFinding(
                    NestedTableCode,
                    "Nested tables",
                    $"{nested} table(s) nested inside a table cell. Their content is lost on render "
                    + "- measured. Tables used for page layout are the common case.",
                    nested,
                    DocxToPdfRisk.Known));
            }

            // A content control that wraps, or sits inside, part of a TABLE. NOT content
            // controls in general: A67 measured a body-level w:sdt SURVIVING the render and
            // excluded controls on that evidence, correctly for what it measured. Re-measured
            // 2026-08-27, the same construct one level in is lost while the body-level one still
            // renders, so the finding is scoped to the table cases.
            //
            // ALL THREE table wrapper positions, because a code review measured that the first
            // version reported only one of them and stayed silent on the other two:
            //
            //     w:tc > w:sdt        a control inside a cell        was reported
            //     w:tr > w:sdt > w:tc a cell wrapped in a control    was SILENT, and is lost
            //     w:tbl > w:sdt > w:tr a row wrapped in a control    was SILENT, and is lost
            //
            // The neighbouring plain cell renders in both silent cases, so this is a real loss
            // rather than a broken fixture - and a Word form laid out in a table, which is what
            // the finding's own message advertises, is usually built from cell-level controls.
            int controlsInTables = body is null ? 0 : ControlsInTables(body);

            if (controlsInTables > 0)
            {
                findings.Add(new DocxToPdfPreflightFinding(
                    ControlInCellCode,
                    "Content controls in tables",
                    $"{controlsInTables} content control(s) in a table - inside a cell, or wrapping "
                    + "a cell or a row. Their text does not reach the PDF - measured, while the "
                    + "same control at body level renders normally. A form or template laid out in "
                    + "a table is the common case.",
                    controlsInTables,
                    DocxToPdfRisk.Known));
            }

            return new DocxToPdfPreflightReport(findings);
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException(ReadFailure, ex);
        }
    }
    /// <summary>
    /// Every table cell in the body that the PDF renderer actually lays out — which excludes the
    /// ones inside a text box.
    /// </summary>
    /// <remarks>
    /// <b>A text box is the reason this is not simply <c>Descendants&lt;TableCell&gt;()</c>.</b>
    /// Measured 2026-08-27: a table inside <c>w:txbxContent</c> renders to the PDF with its content
    /// intact — both a nested table and a content control inside one survive — yet a plain
    /// <c>Descendants</c> sweep walks straight into it and reported both as losses. A preflight
    /// that cries wolf on content the reader can see is worse than one reporting nothing, because
    /// it is the findings a caller stops believing.
    ///
    /// <para>This is the same <c>Descendants</c> trap <c>CLAUDE.md</c> records twice — once where
    /// it deleted text-box content in <c>DocxEditor</c>, and once where it swept a nested table's
    /// rows into its container's expansion.</para>
    /// </remarks>
    private static IEnumerable<TableCell> Cells(Body body) =>
        body.Descendants<TableCell>().Where(c => !c.Ancestors<TextBoxContent>().Any());

    /// <summary>
    /// How many content controls sit in a table position the renderer drops, counting the
    /// OUTERMOST one at each site.
    /// </summary>
    /// <remarks>
    /// Outermost wins because that is what the render does: measured, a control nested inside
    /// another in the same cell loses its text exactly once — the outer one is dropped whole and
    /// takes the inner with it. Counting both would inflate a number a caller uses to decide
    /// whether the document is worth opening.
    /// </remarks>
    private static int ControlsInTables(Body body)
    {
        int count = 0;

        foreach (var table in body.Descendants<Table>().Where(t => !t.Ancestors<TextBoxContent>().Any()))
        {
            foreach (var child in table.Elements())
            {
                if (child is SdtRow)
                {
                    // The whole row is wrapped. Nothing inside it renders, so nothing inside it
                    // is counted again.
                    count++;
                    continue;
                }

                if (child is not TableRow row) continue;

                foreach (var cell in row.Elements())
                {
                    if (cell is SdtCell)
                    {
                        count++;
                        continue;
                    }

                    if (cell is TableCell plain) count += plain.Elements<SdtBlock>().Count();
                }
            }
        }

        return count;
    }
}
