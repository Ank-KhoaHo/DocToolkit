using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit;

/// <summary>
/// Compares two versions of a document and returns the later one with the differences marked as
/// tracked changes (A118).
/// </summary>
/// <remarks>
/// <b>The result is an ordinary .docx carrying revisions</b>, not a report — so Word shows it the
/// way it shows any tracked-changes document, and <c>DocxEditor.Revisions</c>,
/// <c>AcceptRevisions</c> and <c>RejectRevisions</c> all read and apply it without knowing it came
/// from here.
///
/// <para>
/// <b>Paragraph text is what gets compared, and everything else is REPORTED rather than silently
/// skipped.</b> Tables, lists and formatting changes are named in the warnings of
/// <see cref="CompareWithReport(byte[], byte[], string)"/>. A comparison that quietly mis-marked a
/// table would be worse than one that says what it did not look at — the same contract
/// <c>DocToDocxConverter</c> offers through <c>LegacyDocOptions.AllowContentLoss</c>.
/// </para>
///
/// <para>
/// <b>A formatting-only change is not detected at all</b>, and is not reported as a text change
/// either. <c>DocxRevisionKind</c> has no formatting member, and inventing one to describe
/// something this does not measure would be worse than the gap.
/// </para>
/// </remarks>
public static class DocxCompare
{
    /// <summary>Code on the warning raised when a document contains tables.</summary>
    internal const string TablesNotCompared = "COMPARE-TABLES-NOT-COMPARED";

    /// <summary>Code on the warning raised when the paragraph counts differ structurally.</summary>
    internal const string StructureChanged = "COMPARE-STRUCTURE-CHANGED";

    /// <summary>Code on the warning that formatting is never compared.</summary>
    internal const string FormattingNotCompared = "COMPARE-FORMATTING-NOT-COMPARED";

    private const string CompareFailure =
        "Failed to compare the documents. See the inner exception for details.";

    /// <summary>
    /// Returns <paramref name="revised"/> with its differences from <paramref name="original"/>
    /// marked as tracked insertions and deletions.
    /// </summary>
    /// <remarks>
    /// <b>Only paragraph text is compared.</b> Use
    /// <see cref="CompareWithReport(byte[], byte[], string)"/> to find out what was not — this
    /// overload discards that, and a caller who does not know whether the documents contain tables
    /// should not be using it.
    ///
    /// <b>Comparing a document with itself produces no revisions</b>, rather than a document marked
    /// entirely rewritten.
    ///
    /// The cost is dominated by the differing text rather than by document size: the shared prefix
    /// and suffix are removed first, so two revisions of one document are cheap and two unrelated
    /// documents are the bounded worst case.
    /// </remarks>
    /// <param name="original">The earlier version.</param>
    /// <param name="revised">The later version, which the result is built from.</param>
    /// <param name="author">The name recorded against each revision.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">Either document is empty, or <paramref name="author"/> is blank.</exception>
    /// <exception cref="DocumentConversionException">Either package could not be opened or edited.</exception>
    public static byte[] Compare(byte[] original, byte[] revised, string author) =>
        CompareWithReport(original, revised, author).Value;

    /// <summary>
    /// Compares two documents and reports what the comparison did not look at.
    /// </summary>
    /// <inheritdoc cref="Compare(byte[], byte[], string)" path="/remarks|/param|/exception"/>
    /// <returns>
    /// The marked-up document, with a warning for every construct present but not compared.
    /// <c>HasLoss</c> is true whenever anything was skipped, which is the signal that the verdict
    /// covers less than the document.
    /// </returns>
    public static ConversionResult<byte[]> CompareWithReport(byte[] original, byte[] revised, string author)
    {
        ArgumentNullException.ThrowIfNull(original);
        ArgumentNullException.ThrowIfNull(revised);
        ArgumentException.ThrowIfNullOrWhiteSpace(author);
        if (original.Length == 0) throw new ArgumentException("DOCX content was empty.", nameof(original));
        if (revised.Length == 0) throw new ArgumentException("DOCX content was empty.", nameof(revised));

        try
        {
            var before = Paragraphs(original, out var originalTables);

            using var ms = new MemoryStream();
            ms.Write(revised, 0, revised.Length);
            ms.Position = 0;

            var warnings = new List<ConversionWarning>();
            var revisedTables = 0;

            using (var document = WordprocessingDocument.Open(ms, true))
            {
                var body = Body(document);

                revisedTables = body.Elements<Table>().Count();

                var after = body.Elements<Paragraph>().ToList();
                var stamp = DateTime.UtcNow;

                // Paragraphs are paired by position. Pairing them by content would be a second
                // diff over a different alphabet, and A118 scoped this to text within a paragraph
                // rather than to paragraph-level moves - which is why a count change is REPORTED
                // rather than papered over.
                for (var i = 0; i < after.Count; i++)
                {
                    var originalText = i < before.Count ? before[i] : string.Empty;
                    MarkParagraph(after[i], originalText, author, stamp);
                }

                document.MainDocumentPart!.Document!.Save();
            }

            var afterCount = CountParagraphs(ms);
            if (before.Count != afterCount)
            {
                warnings.Add(new ConversionWarning(
                    StructureChanged,
                    $"The documents have different paragraph counts ({before.Count} and "
                    + $"{afterCount}). Paragraphs are paired by position, so a paragraph "
                    + "inserted or removed in the middle shifts every pair after it and its "
                    + "neighbours will read as heavily rewritten.",
                    ConversionLossKind.Approximation));
            }

            if (originalTables > 0 || revisedTables > 0)
            {
                warnings.Add(new ConversionWarning(
                    TablesNotCompared,
                    $"{Math.Max(originalTables, revisedTables)} table(s) are present and were not "
                    + "compared. A row inserted mid-table is not the same edit as its text appearing "
                    + "elsewhere, and marking it as one would be a wrong answer rather than a coarse "
                    + "one.",
                    ConversionLossKind.Omission));
            }

            warnings.Add(new ConversionWarning(
                FormattingNotCompared,
                "Formatting changes are never compared. A paragraph whose text is unchanged but "
                + "whose styling differs produces no revision, because DocxRevisionKind has no "
                + "formatting member and reporting one would describe something this did not measure.",
                ConversionLossKind.Omission));

            return new ConversionResult<byte[]>(ms.ToArray(), warnings);
        }
        catch (Exception ex) when (ex is not DocumentConversionException and not ArgumentException)
        {
            throw new DocumentConversionException(CompareFailure, ex);
        }
    }

    /// <summary>
    /// The document body, or a typed failure naming the likely cause.
    /// </summary>
    /// <remarks>
    /// <b>One place decides what "this is not really a .docx" means</b>, rather than the same
    /// null-check written at each of the three points that needs it. Three copies of one rule is
    /// three chances for them to disagree about the message a caller sees, and three branches
    /// describing one fact.
    /// </remarks>
    private static Body Body(WordprocessingDocument document) =>
        document.MainDocumentPart?.Document?.Body
        ?? throw new DocumentConversionException(
            "Document has no body. This usually means the file is not really a .docx (for example "
            + "it was renamed from another format) or the upload is corrupt.");

    private static int CountParagraphs(MemoryStream docx)
    {
        var at = docx.Position;
        docx.Position = 0;
        try
        {
            using var document = WordprocessingDocument.Open(docx, false);
            return Body(document).Elements<Paragraph>().Count();
        }
        finally
        {
            docx.Position = at;
        }
    }

    /// <summary>Each top-level paragraph's text, plus how many tables the document holds.</summary>
    private static List<string> Paragraphs(byte[] docx, out int tables)
    {
        using var ms = new MemoryStream(docx);
        using var document = WordprocessingDocument.Open(ms, false);

        var body = Body(document);

        tables = body.Elements<Table>().Count();

        // Elements, never Descendants - the rule ContentControls records and this repository has
        // paid for twice. A paragraph inside a table or a text box is not a top-level paragraph.
        return [.. body.Elements<Paragraph>().Select(TextOf)];
    }

    private static string TextOf(Paragraph paragraph) =>
        string.Concat(paragraph.Descendants<Text>().Select(t => t.Text));

    /// <summary>
    /// Rewrites one paragraph's runs so its difference from <paramref name="originalText"/> is
    /// marked.
    /// </summary>
    /// <remarks>
    /// <b>An unchanged paragraph is left completely alone</b> — not rewritten with identical text —
    /// so its runs keep their formatting and a document compared with itself comes back byte-for-byte
    /// equivalent in every way that matters. That is also what makes the zero-revision case
    /// meaningful rather than accidental.
    ///
    /// <b>A deletion becomes <c>w:del</c> wrapping <c>w:delText</c></b>, which is the only form Word
    /// accepts: a deleted run keeping a plain <c>w:t</c> renders as ordinary text with a revision
    /// mark around it, and accepting the change then leaves the text behind.
    /// </remarks>
    private static void MarkParagraph(Paragraph paragraph, string originalText, string author, DateTime stamp)
    {
        var revisedText = TextOf(paragraph);
        if (string.Equals(originalText, revisedText, StringComparison.Ordinal)) return;

        var spans = WordDiff.Diff(WordDiff.Split(originalText), WordDiff.Split(revisedText));

        // The first run's properties are carried onto every rebuilt run, so the paragraph keeps its
        // look. Per-run formatting within a rewritten paragraph is not preserved - that is the
        // formatting limitation this class documents, made concrete.
        var template = paragraph.Elements<Run>().FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;

        foreach (var run in paragraph.Elements<Run>().ToList()) run.Remove();
        foreach (var ins in paragraph.Elements<InsertedRun>().ToList()) ins.Remove();
        foreach (var del in paragraph.Elements<DeletedRun>().ToList()) del.Remove();

        var id = 1;
        foreach (var span in spans)
        {
            var text = string.Concat(span.Words);
            if (text.Length == 0) continue;

            switch (span.Kind)
            {
                case WordDiffKind.Same:
                    paragraph.AppendChild(NewRun(text, template));
                    break;

                case WordDiffKind.Inserted:
                    paragraph.AppendChild(new InsertedRun(NewRun(text, template))
                    {
                        Author = author,
                        Date = stamp,
                        Id = (id++).ToString(System.Globalization.CultureInfo.InvariantCulture),
                    });
                    break;

                case WordDiffKind.Deleted:
                    var deleted = new Run(new DeletedText(text) { Space = SpaceProcessingModeValues.Preserve });
                    if (template is not null) deleted.RunProperties = (RunProperties)template.CloneNode(true);
                    paragraph.AppendChild(new DeletedRun(deleted)
                    {
                        Author = author,
                        Date = stamp,
                        Id = (id++).ToString(System.Globalization.CultureInfo.InvariantCulture),
                    });
                    break;
            }
        }
    }

    private static Run NewRun(string text, RunProperties? template)
    {
        var run = new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        if (template is not null) run.RunProperties = (RunProperties)template.CloneNode(true);
        return run;
    }
}
