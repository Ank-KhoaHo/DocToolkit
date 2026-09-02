using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// Comparing two versions of a document (A118).
///
/// <b>The result is read back through the shipped revision reader</b> rather than by inspecting
/// XML. <c>DocxReview.Inspect</c>, <c>AcceptRevisions</c> and <c>RejectRevisions</c> already
/// existed, so a comparison is verifiable by machinery that had to work anyway — and an assertion
/// on what Word will actually do is worth more than one on the markup.
///
/// The word-level diff itself is proved separately in <c>WordDiffTests</c>, against string
/// sequences where a failure names the input rather than a .docx.
/// </summary>
public class DocxCompareTests
{
    private const string Author = "Reviewer";

    private static byte[] Doc(params string[] paragraphs) =>
        DocxEditor.Create([.. paragraphs.Select(DocxBlock.Paragraph)]);

    private static IReadOnlyList<DocxRevision> RevisionsOf(byte[] docx) => DocxReview.Inspect(docx).Revisions;

    // ---------- the negative control comes first, because everything else leans on it ----------

    /// <summary>
    /// A document compared with itself must report ZERO revisions. Nothing else proves the comparer
    /// can return nothing — every assertion below is satisfied by one that marks everything changed.
    /// </summary>
    [Fact]
    public void ADocumentComparedWithItselfHasNoRevisions()
    {
        var docx = Doc("Acme Corporation", "Invoice 42", "Terms apply");

        var compared = DocxCompare.Compare(docx, docx, Author);

        Assert.Empty(RevisionsOf(compared));
    }

    [Fact]
    public void UnchangedTextSurvivesUnmarked()
    {
        var original = Doc("first paragraph", "second paragraph");
        var revised = Doc("first paragraph", "second paragraph CHANGED");

        var compared = DocxCompare.Compare(original, revised, Author);

        // The untouched paragraph contributes nothing, so every revision belongs to the second.
        Assert.All(RevisionsOf(compared), r => Assert.DoesNotContain("first", r.AffectedText, StringComparison.Ordinal));
    }

    // ---------- what a comparison actually reports ----------

    [Fact]
    public void ReportsAnInsertionForAddedText()
    {
        var compared = DocxCompare.Compare(Doc("the fox"), Doc("the quick fox"), Author);

        var inserted = Assert.Single(RevisionsOf(compared), r => r.Kind == DocxRevisionKind.Insertion);
        Assert.Contains("quick", inserted.AffectedText, StringComparison.Ordinal);
    }

    [Fact]
    public void ReportsADeletionForRemovedText()
    {
        var compared = DocxCompare.Compare(Doc("the quick fox"), Doc("the fox"), Author);

        var deleted = Assert.Single(RevisionsOf(compared), r => r.Kind == DocxRevisionKind.Deletion);
        Assert.Contains("quick", deleted.AffectedText, StringComparison.Ordinal);
    }

    /// <summary>
    /// THE discriminating case. A comparer that marks the whole paragraph changed passes any test
    /// asserting merely that a difference was found — only asserting that the UNCHANGED words are
    /// not marked fails against it.
    /// </summary>
    [Fact]
    public void MarksOnlyTheWordsThatChanged()
    {
        var compared = DocxCompare.Compare(Doc("the quick brown fox"), Doc("the quick red fox"), Author);

        var revisions = RevisionsOf(compared);
        var marked = string.Concat(revisions.Select(r => r.AffectedText));

        Assert.Contains("red", marked, StringComparison.Ordinal);
        Assert.Contains("brown", marked, StringComparison.Ordinal);
        Assert.DoesNotContain("quick", marked, StringComparison.Ordinal);
        Assert.DoesNotContain("fox", marked, StringComparison.Ordinal);
    }

    [Fact]
    public void RecordsTheAuthorOnEveryRevision()
    {
        var compared = DocxCompare.Compare(Doc("before"), Doc("after"), "Ada Lovelace");

        var revisions = RevisionsOf(compared);
        Assert.NotEmpty(revisions);
        Assert.All(revisions, r => Assert.Equal("Ada Lovelace", r.Author));
    }

    // ---------- the round trip through the existing appliers ----------

    /// <summary>
    /// Accepting every revision must yield the revised document's text, and rejecting every one must
    /// yield the original's. That is the strongest statement available about a comparison being
    /// right, and it costs nothing because both appliers already ship.
    /// </summary>
    [Fact]
    public void AcceptingEveryRevisionYieldsTheRevisedText()
    {
        var revised = Doc("the quick red fox jumps");
        var compared = DocxCompare.Compare(Doc("the quick brown fox"), revised, Author);

        Assert.Equal(DocxEditor.ExtractText(revised), DocxEditor.ExtractText(DocxReview.AcceptRevisions(compared)));
    }

    [Fact]
    public void RejectingEveryRevisionYieldsTheOriginalText()
    {
        var original = Doc("the quick brown fox");
        var compared = DocxCompare.Compare(original, Doc("the quick red fox jumps"), Author);

        Assert.Equal(DocxEditor.ExtractText(original), DocxEditor.ExtractText(DocxReview.RejectRevisions(compared)));
    }

    // ---------- the report, which is what makes the scope honest ----------

    [Fact]
    public void ReportsThatFormattingIsNeverCompared()
    {
        var result = DocxCompare.CompareWithReport(Doc("same"), Doc("same"), Author);

        Assert.Contains(result.Warnings, w => w.Code == DocxCompare.FormattingNotCompared);
        Assert.True(result.HasLoss, "a comparison that never looks at formatting has lost something");
    }

    [Fact]
    public void ReportsAParagraphCountChange()
    {
        var result = DocxCompare.CompareWithReport(Doc("one"), Doc("one", "two"), Author);

        Assert.Contains(result.Warnings, w => w.Code == DocxCompare.StructureChanged);
    }

    /// <summary>
    /// The positive control for the warning above: two documents with the same paragraph count must
    /// NOT raise it, or the warning says nothing.
    /// </summary>
    [Fact]
    public void PositiveControl_NoStructureWarningWhenTheCountsMatch()
    {
        var result = DocxCompare.CompareWithReport(Doc("one", "two"), Doc("one", "changed"), Author);

        Assert.DoesNotContain(result.Warnings, w => w.Code == DocxCompare.StructureChanged);
    }

    // ---------- argument guards ----------

    [Fact]
    public void RejectsNullDocumentsAndABlankAuthor()
    {
        var docx = Doc("x");

        Assert.Throws<ArgumentNullException>(() => DocxCompare.Compare(null!, docx, Author));
        Assert.Throws<ArgumentNullException>(() => DocxCompare.Compare(docx, null!, Author));
        Assert.Throws<ArgumentException>(() => DocxCompare.Compare(docx, docx, "  "));
    }

    [Fact]
    public void RejectsEmptyContent()
    {
        var docx = Doc("x");

        Assert.Equal("original", Assert.Throws<ArgumentException>(
            () => DocxCompare.Compare([], docx, Author)).ParamName);
        Assert.Equal("revised", Assert.Throws<ArgumentException>(
            () => DocxCompare.Compare(docx, [], Author)).ParamName);
    }

    [Fact]
    public void RejectsBytesThatAreNotADocx()
    {
        Assert.Throws<DocumentConversionException>(
            () => DocxCompare.Compare("not a docx"u8.ToArray(), Doc("x"), Author));
    }

    // ---------- the promise that an unchanged paragraph is left ALONE ----------

    /// <summary>
    /// An unchanged paragraph must keep its own runs, not be rebuilt from its text.
    /// </summary>
    /// <remarks>
    /// <b>This test exists because sabotage found nothing without it.</b> Deleting the early-out in
    /// <c>MarkParagraph</c> left every other test in this file green — identical text diffs to all
    /// <c>Same</c> spans, so the text comes back right either way. What is lost is per-run
    /// formatting: the rebuild carries the FIRST run's properties onto everything, so a paragraph
    /// that was half bold comes back uniformly bold with no revision to explain it.
    ///
    /// The class documents that behaviour as a promise, and a promise nothing tests is invisible to
    /// sabotage — the lesson A117 is filed on.
    /// </remarks>
    [Fact]
    public void AnUnchangedParagraphKeepsItsSeparateRuns()
    {
        var docx = TwoRunParagraph();
        Assert.Equal(2, RunCount(docx));

        var compared = DocxCompare.Compare(docx, docx, Author);

        Assert.Equal(2, RunCount(compared));
        Assert.Empty(RevisionsOf(compared));
    }

    /// <summary>
    /// And the same when a DIFFERENT paragraph changes: the untouched one is still not rebuilt.
    /// Comparing a document with itself alone would not distinguish "left alone" from "nothing to
    /// do at all".
    /// </summary>
    [Fact]
    public void AnUnchangedParagraphIsLeftAloneEvenWhenAnotherChanges()
    {
        var original = TwoRunParagraph("second paragraph");
        var revised = TwoRunParagraph("second paragraph CHANGED");

        var compared = DocxCompare.Compare(original, revised, Author);

        // The first paragraph's two runs survive; the second is the one carrying revisions.
        Assert.Equal(2, FirstParagraphRunCount(compared));
        Assert.NotEmpty(RevisionsOf(compared));
    }

    /// <summary>A document whose first paragraph is two runs with different formatting.</summary>
    private static byte[] TwoRunParagraph(params string[] extraParagraphs)
    {
        var docx = Doc([.. new[] { "bold plain" }.Concat(extraParagraphs)]);

        var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        ms.Position = 0;

        using (var document = WordprocessingDocument.Open(ms, true))
        {
            var paragraph = document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
            foreach (var run in paragraph.Elements<Run>().ToList()) run.Remove();

            paragraph.AppendChild(new Run(new RunProperties(new Bold()), new Text("bold ") { Space = SpaceProcessingModeValues.Preserve }));
            paragraph.AppendChild(new Run(new Text("plain") { Space = SpaceProcessingModeValues.Preserve }));

            document.MainDocumentPart.Document.Save();
        }

        return ms.ToArray();
    }

    private static int RunCount(byte[] docx) => FirstParagraphRunCount(docx);

    private static int FirstParagraphRunCount(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var document = WordprocessingDocument.Open(ms, false);
        return document.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>().First().Elements<Run>().Count();
    }

    // ---------- the table warning, and the paths a plain two-paragraph fixture never reaches ----------

    [Fact]
    public void ReportsThatTablesWereNotCompared()
    {
        var withTable = WithATable(Doc("intro"));

        var result = DocxCompare.CompareWithReport(withTable, withTable, Author);

        var warning = Assert.Single(result.Warnings, w => w.Code == DocxCompare.TablesNotCompared);
        Assert.Contains("not", warning.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// The positive control. A comparison that raised the table warning unconditionally would
    /// satisfy the test above and say nothing at all.
    /// </summary>
    [Fact]
    public void PositiveControl_NoTableWarningWhenNeitherDocumentHasOne()
    {
        var result = DocxCompare.CompareWithReport(Doc("intro"), Doc("intro"), Author);

        Assert.DoesNotContain(result.Warnings, w => w.Code == DocxCompare.TablesNotCompared);
    }

    /// <summary>
    /// Comparing an already-compared document must not stack revision markup on itself. The
    /// previous run's w:ins and w:del elements are cleared before the paragraph is rebuilt, and
    /// nothing else in this file puts them there to begin with.
    /// </summary>
    [Fact]
    public void ComparingAnAlreadyComparedDocumentDoesNotStackRevisions()
    {
        var once = DocxCompare.Compare(Doc("the quick brown fox"), Doc("the quick red fox"), Author);

        var twice = DocxCompare.Compare(Doc("the quick brown fox"), once, Author);

        // Accepting everything still lands on the text the second document actually holds.
        Assert.Equal(
            DocxEditor.ExtractText(DocxReview.AcceptRevisions(once)),
            DocxEditor.ExtractText(DocxReview.AcceptRevisions(twice)));
    }

    /// <summary>
    /// A deleted run carries the paragraph's formatting rather than arriving unstyled, so a
    /// reviewer sees the removed text as it looked.
    /// </summary>
    [Fact]
    public void ADeletedRunKeepsTheParagraphsFormatting()
    {
        // The template comes from the REVISED document, which is what the result is built from -
        // so that is the one that has to carry the formatting. The first version of this test had
        // them the wrong way round and was wrong about the fixture rather than about the code.
        var original = Doc("bold plain extra");
        var revised = TwoRunParagraph();

        var compared = DocxCompare.Compare(original, revised, Author);

        using var ms = new MemoryStream(compared);
        using var document = WordprocessingDocument.Open(ms, false);
        var deleted = document.MainDocumentPart!.Document!.Body!.Descendants<DeletedRun>().ToList();

        Assert.NotEmpty(deleted);
        Assert.All(deleted, d => Assert.NotNull(d.Descendants<Run>().First().RunProperties));
    }

    /// <summary>A document with a table after its paragraphs.</summary>
    private static byte[] WithATable(byte[] docx)
    {
        var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        ms.Position = 0;

        using (var document = WordprocessingDocument.Open(ms, true))
        {
            var body = document.MainDocumentPart!.Document!.Body!;
            body.AppendChild(new Table(
                new TableProperties(),
                new TableGrid(new GridColumn()),
                new TableRow(new TableCell(new Paragraph(new Run(new Text("cell")))))));
            document.MainDocumentPart.Document.Save();
        }

        return ms.ToArray();
    }
}
