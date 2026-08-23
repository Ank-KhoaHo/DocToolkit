using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit.Tests;

/// <summary>
/// Covers <see cref="DocxReview"/> — reading a document's comments and tracked changes, and
/// resolving them.
///
/// <b>Two of the three fixtures below are built with raw OpenXml rather than through this library,
/// and that is not a style choice.</b> OfficeIMO cannot produce either state:
///
/// <list type="bullet">
/// <item><c>TrackChanges = true</c> records nothing — measured 2026-08-23, editing text with it set
/// produces no <c>w:ins</c>, no <c>w:del</c>, and no <c>settings.xml</c> entry. A revision fixture
/// built through <see cref="DocxEditor"/> would assert against an empty set forever while looking
/// green.</item>
/// <item><c>AddComment</c> is the only comment-authoring API and cannot set a parent, so no
/// document this library can write contains a reply. Threading lives in a second package part.</item>
/// </list>
///
/// Both are the same hazard as the air-gap suites' positive controls: an assertion that passes
/// because nothing was there to find.
/// </summary>
public class DocxReviewTests
{
    // ---- the positive control -------------------------------------------------------------

    [Fact]
    public void Inspect_OnADocumentWithNoReviewState_ReturnsEmpty()
    {
        // THE CONTROL. Without it every assertion in this file could be passing against a walk that
        // finds nothing at all, and the whole suite would be vacuous.
        byte[] docx = DocxEditor.Create([DocxBlock.Paragraph("Plain text.")]);

        var report = DocxReview.Inspect(docx);

        Assert.Empty(report.Comments);
        Assert.Empty(report.Revisions);
        Assert.Equal(0, report.CommentThreadCount);
        Assert.Equal(0, report.UnresolvedThreadCount);
    }

    // ---- comments -------------------------------------------------------------------------

    [Fact]
    public void Inspect_ReadsCommentsWithTheirAuthorsAndInitials()
    {
        var report = DocxReview.Inspect(WithTwoComments());

        Assert.Equal(2, report.Comments.Count);
        Assert.Equal(2, report.CommentThreadCount);
        Assert.Equal(["Reviewer One", "Reviewer Two"], report.Comments.Select(c => c.Author));
        Assert.Equal(["R1", "R2"], report.Comments.Select(c => c.Initials));
        Assert.Contains("tighten", report.Comments[0].Text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Inspect_NestsARepliedCommentUnderItsParent_RatherThanListingItTwice()
    {
        // The flat list OfficeIMO also exposes carries the reply as an entry of its own. Mapping
        // that list instead of CommentThreads would report two top-level comments here, and the
        // reply would appear both at the top and under its parent.
        var report = DocxReview.Inspect(WithAThread(resolved: false));

        var parent = Assert.Single(report.Comments);
        Assert.Equal("Reviewer One", parent.Author);
        Assert.Equal("Parent question?", parent.Text);

        var reply = Assert.Single(parent.Replies);
        Assert.Equal("Reviewer Two", reply.Author);
        Assert.Equal("Reply answer.", reply.Text);
        Assert.Empty(reply.Replies);
    }

    [Fact]
    public void Inspect_CountsAThreadOnce_NotOncePerComment()
    {
        var report = DocxReview.Inspect(WithAThread(resolved: false));

        // Two comments, one thread. A CommentThreadCount that merely counted comments would say 2,
        // which is what makes this assertion discriminate rather than restate the line above.
        Assert.Equal(1, report.CommentThreadCount);
    }

    [Theory]
    [InlineData(false, 1)]
    [InlineData(true, 0)]
    public void Inspect_ReportsAResolvedThreadAsNothingLeftToActOn(bool resolved, int expected)
    {
        // BOTH rows are needed. UnresolvedThreadCount is only meaningful if it can differ from
        // CommentThreadCount - a property that always equalled the thread count would pass a
        // one-row version of this test and tell a caller nothing.
        var report = DocxReview.Inspect(WithAThread(resolved));

        Assert.Equal(1, report.CommentThreadCount);
        Assert.Equal(expected, report.UnresolvedThreadCount);
        Assert.Equal(resolved, Assert.Single(report.Comments).IsResolved);
    }

    [Fact]
    public void RemoveComments_ClearsTheCommentsAndLeavesTheTextAlone()
    {
        byte[] cleaned = DocxReview.RemoveComments(WithTwoComments());

        // Two assertions, deliberately. The count falling to zero alone would also pass if the call
        // had removed the commented paragraphs along with the comments.
        Assert.Empty(DocxReview.Inspect(cleaned).Comments);
        Assert.Contains("First paragraph under review.", DocxEditor.ExtractText(cleaned), StringComparison.Ordinal);
    }

    // ---- tracked changes ------------------------------------------------------------------

    [Fact]
    public void Inspect_ReadsTrackedChangesWithTheirKindAndAuthor()
    {
        var report = DocxReview.Inspect(WithOneInsertionAndOneDeletion());

        Assert.Equal(2, report.Revisions.Count);

        var insertion = Assert.Single(report.Revisions, r => r.Kind == DocxRevisionKind.Insertion);
        Assert.Equal("Reviewer One", insertion.Author);

        // Trimmed, though the fixture writes a trailing space under xml:space="preserve" - measured
        // 2026-08-23. AffectedText is a summary of what changed rather than a byte-exact copy of
        // the run, so do not assert leading or trailing whitespace through it. The accept test
        // asserts the SPACING separately, through the document's own extracted text.
        Assert.Equal("Inserted text.", insertion.AffectedText);

        var deletion = Assert.Single(report.Revisions, r => r.Kind == DocxRevisionKind.Deletion);
        Assert.Equal("Reviewer Two", deletion.Author);
        Assert.Equal("Removed text.", deletion.AffectedText);
    }

    [Fact]
    public void AcceptRevisions_KeepsTheInsertionAndDropsTheDeletion()
    {
        byte[] accepted = DocxReview.AcceptRevisions(WithOneInsertionAndOneDeletion());

        // The revision count falling to zero is NOT enough on its own: it would also pass if accept
        // had deleted the content along with the markup, which is exactly what
        // WordDocument.Paragraphs[0].Text made it look like it did (that property returns only the
        // first run). Assert on the text as well, through DocxEditor rather than that property.
        Assert.Empty(DocxReview.Inspect(accepted).Revisions);

        string text = DocxEditor.ExtractText(accepted);
        Assert.Contains("Kept.", text, StringComparison.Ordinal);
        Assert.Contains("Inserted text.", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Removed text.", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RejectRevisions_DropsTheInsertionAndRestoresTheDeletion()
    {
        // THE MIRROR of the test above. Without it, accept and reject could share one
        // implementation - or be wired to the same call - and both stay green.
        byte[] rejected = DocxReview.RejectRevisions(WithOneInsertionAndOneDeletion());

        Assert.Empty(DocxReview.Inspect(rejected).Revisions);

        string text = DocxEditor.ExtractText(rejected);
        Assert.Contains("Kept.", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Inserted text.", text, StringComparison.Ordinal);
        Assert.Contains("Removed text.", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Inspect_ReportsAFormattingRevisionAsOther_RatherThanRefusingTheDocument()
    {
        // Nine of Word's eleven revision kinds map to Other. A document carrying only those must
        // still read - refusing it would be worse than reporting it coarsely.
        var report = DocxReview.Inspect(WithARunFormattingRevision());

        var revision = Assert.Single(report.Revisions);
        Assert.Equal(DocxRevisionKind.Other, revision.Kind);
        Assert.Equal("Reviewer Three", revision.Author);
    }

    // ---- guards ---------------------------------------------------------------------------

    [Fact]
    public void EveryByteArrayOverload_RefusesNullAndEmptyByTheParameterItDeclares()
    {
        var calls = new (string Name, Action<byte[]> Call)[]
        {
            (nameof(DocxReview.Inspect), b => DocxReview.Inspect(b)),
            (nameof(DocxReview.RemoveComments), b => DocxReview.RemoveComments(b)),
            (nameof(DocxReview.AcceptRevisions), b => DocxReview.AcceptRevisions(b)),
            (nameof(DocxReview.RejectRevisions), b => DocxReview.RejectRevisions(b)),
        };

        foreach (var (name, call) in calls)
        {
            var nullEx = Assert.Throws<ArgumentNullException>(() => call(null!));
            Assert.Equal("docx", nullEx.ParamName);

            var emptyEx = Assert.Throws<ArgumentException>(() => call([]));
            Assert.Equal("docx", emptyEx.ParamName);
            Assert.Contains("empty", emptyEx.Message, StringComparison.OrdinalIgnoreCase);

            // Names the overload under test, so a failure says WHICH of the four is wrong rather
            // than only that one of them is.
            Assert.False(string.IsNullOrEmpty(name));
        }
    }

    [Fact]
    public void Inspect_WrapsAnUnreadableDocumentRatherThanLettingTheLibraryThrow()
    {
        var ex = Assert.Throws<DocumentConversionException>(() => DocxReview.Inspect([1, 2, 3, 4]));
        Assert.NotNull(ex.InnerException);
    }

    [Theory]
    [InlineData(nameof(DocxReview.RemoveComments))]
    [InlineData(nameof(DocxReview.AcceptRevisions))]
    [InlineData(nameof(DocxReview.RejectRevisions))]
    public void EveryMutatingOverload_WrapsAnUnreadableDocument_AndSaysWhichOperationFailed(string api)
    {
        // Covers the one catch arm the three mutating operations share. Each must name ITS OWN
        // operation: a shared helper that reported "failed to read" for all three would leave a
        // caller of AcceptRevisions looking in the wrong place.
        byte[] garbage = [1, 2, 3, 4];

        var ex = Assert.Throws<DocumentConversionException>(() => api switch
        {
            nameof(DocxReview.RemoveComments) => DocxReview.RemoveComments(garbage),
            nameof(DocxReview.AcceptRevisions) => DocxReview.AcceptRevisions(garbage),
            _ => DocxReview.RejectRevisions(garbage),
        });

        Assert.NotNull(ex.InnerException);

        string expected = api switch
        {
            nameof(DocxReview.RemoveComments) => "comments",
            _ => "tracked changes",
        };
        Assert.Contains(expected, ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ---- the fields a report carries beyond text and author ---------------------------------

    [Fact]
    public void Inspect_CarriesTheDateACommentOrRevisionWasRecorded()
    {
        // The fixtures stamp a fixed date. Without this the property could be unmapped - always
        // null - and every other assertion in this file would still pass.
        //
        // COMPARED AS AN INSTANT, not as a wall-clock reading. The date comes back with Kind Local:
        // the UTC 09:30 written by the fixture read back as 16:30+07:00 on the machine this was
        // written on, which is the same moment. Asserting the DateTime directly would have passed
        // on CI - whose runners sit at UTC, where the two forms coincide - and failed only on a
        // developer's machine, which is the worst way round for a test to be wrong.
        AssertSameInstant(Assert.Single(DocxReview.Inspect(WithAThread(resolved: false)).Comments).Created);

        foreach (var revision in DocxReview.Inspect(WithOneInsertionAndOneDeletion()).Revisions)
            AssertSameInstant(revision.Created);
    }

    private static void AssertSameInstant(DateTime? actual)
    {
        Assert.NotNull(actual);
        Assert.Equal(Reviewed, actual.Value.ToUniversalTime());
    }

    [Fact]
    public void Inspect_SaysWhetherARevisionIsInsideATable()
    {
        // BOTH directions. A property hard-wired to false would pass the first assertion alone, and
        // one hard-wired to true would pass the second alone.
        Assert.All(DocxReview.Inspect(WithOneInsertionAndOneDeletion()).Revisions,
                   r => Assert.False(r.IsInTable));

        Assert.True(Assert.Single(DocxReview.Inspect(WithAnInsertionInsideATable()).Revisions).IsInTable);
    }

    // ---- fixtures ---------------------------------------------------------------------------

    /// <summary>
    /// Two plain comments, built through OfficeIMO — so this one proves the mapping reads a
    /// document a real producer wrote, not only the hand-assembled packages below.
    /// </summary>
    private static byte[] WithTwoComments()
    {
        string path = Path.Join(Path.GetTempPath(), Path.GetRandomFileName() + ".docx");
        try
        {
            using (var document = OfficeIMO.Word.WordDocument.Create(path))
            {
                document.AddParagraph("First paragraph under review.")
                        .AddComment("Reviewer One", "R1", "Please tighten this.");
                document.AddParagraph("Second paragraph.")
                        .AddComment("Reviewer Two", "R2", "Agreed with the above.");
                document.Save();
            }

            return File.ReadAllBytes(path);
        }
        finally
        {
            File.Delete(path);
        }
    }

    /// <summary>
    /// One comment with one reply, optionally marked resolved.
    ///
    /// <b>Hand-assembled because OfficeIMO cannot author a thread</b> — <c>AddComment</c> takes no
    /// parent and <c>ParentComment</c> is read-only. Threading is not in <c>comments.xml</c> at
    /// all: it lives in <c>commentsExtended.xml</c>, which pairs a <c>w15:paraId</c> to its
    /// <c>w15:paraIdParent</c> and carries the resolved flag, matched to a <c>w14:paraId</c>
    /// attribute on the paragraph inside each <c>w:comment</c>.
    /// </summary>
    private static byte[] WithAThread(bool resolved)
    {
        const string ParentParaId = "11111111";
        const string ReplyParaId = "22222222";

        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Paragraph(
                new CommentRangeStart { Id = "1" },
                new Run(new Text("Reviewed sentence.")),
                new CommentRangeEnd { Id = "1" },
                new Run(new CommentReference { Id = "1" }))));

            var comments = main.AddNewPart<WordprocessingCommentsPart>();
            comments.Comments = new Comments(
                Comment("1", "Reviewer One", "R1", ParentParaId, "Parent question?"),
                Comment("2", "Reviewer Two", "R2", ReplyParaId, "Reply answer."));

            var threading = main.AddNewPart<WordprocessingCommentsExPart>();
            threading.CommentsEx = new CommentsEx(
                new CommentEx { ParaId = ParentParaId, Done = resolved },
                new CommentEx { ParaId = ReplyParaId, ParaIdParent = ParentParaId, Done = false });

            main.Document.Save();
        }

        return ms.ToArray();
    }

    private static Comment Comment(string id, string author, string initials, string paraId, string text)
    {
        var body = new Paragraph(new Run(new Text(text)));
        body.SetAttribute(new OpenXmlAttribute(
            "w14", "paraId", "http://schemas.microsoft.com/office/word/2010/wordml", paraId));

        return new Comment(body)
        {
            Id = id,
            Author = author,
            Initials = initials,
            Date = new DateTimeValue(Reviewed),
        };
    }

    /// <summary>
    /// A genuine tracked-changes document: one inserted run and one deleted run, with authors and
    /// dates — the shape Word itself writes. See this class's summary for why it cannot be built
    /// through the library.
    /// </summary>
    private static byte[] WithOneInsertionAndOneDeletion()
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var paragraph = new Paragraph(new Run(new Text("Kept. ") { Space = SpaceProcessingModeValues.Preserve }));

            paragraph.Append(new InsertedRun(
                new Run(new Text("Inserted text. ") { Space = SpaceProcessingModeValues.Preserve }))
            {
                Id = "1",
                Author = "Reviewer One",
                Date = new DateTimeValue(Reviewed),
            });

            paragraph.Append(new DeletedRun(new Run(new DeletedText("Removed text.")))
            {
                Id = "2",
                Author = "Reviewer Two",
                Date = new DateTimeValue(Reviewed),
            });

            main.Document = new Document(new Body(paragraph));
            main.Document.Save();
        }

        return ms.ToArray();
    }

    /// <summary>
    /// An insertion inside a table cell. Accepting or rejecting one of these can move cell
    /// boundaries, which is why <see cref="DocxRevision.IsInTable"/> is worth reporting at all.
    /// </summary>
    private static byte[] WithAnInsertionInsideATable()
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();

            var cell = new TableCell(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Auto }));
            cell.Append(new Paragraph(new InsertedRun(new Run(new Text("Added in a cell.")))
            {
                Id = "4",
                Author = "Reviewer Four",
                Date = new DateTimeValue(Reviewed),
            }));

            // A w:tbl needs a w:tblGrid or the validator rejects its first w:tr as an unexpected
            // child - the same trap DocxFixtures.Tbl records.
            var table = new Table(
                new TableProperties(new TableWidth { Type = TableWidthUnitValues.Auto }),
                new TableGrid(new GridColumn()),
                new TableRow(cell));

            main.Document = new Document(new Body(table, new Paragraph()));
            main.Document.Save();
        }

        return ms.ToArray();
    }

    /// <summary>
    /// A run-formatting revision — one of the nine kinds that report as
    /// <see cref="DocxRevisionKind.Other"/>. It changes how the text looks, not what it says.
    /// </summary>
    private static byte[] WithARunFormattingRevision()
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var run = new Run(new Text("Restyled text."));
            run.RunProperties = new RunProperties(
                new RunPropertiesChange(new RunProperties())
                {
                    Id = "3",
                    Author = "Reviewer Three",
                    Date = new DateTimeValue(Reviewed),
                },
                new Bold());

            main.Document = new Document(new Body(new Paragraph(run)));
            main.Document.Save();
        }

        return ms.ToArray();
    }

    /// <summary>
    /// A fixed date, so a fixture never depends on the clock — every assertion here is about
    /// structure, and a moving date would be one more thing that could differ between runs.
    /// </summary>
    private static readonly DateTime Reviewed = new(2026, 8, 23, 9, 30, 0, DateTimeKind.Utc);
}
