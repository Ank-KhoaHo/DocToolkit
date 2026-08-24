using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

/// <summary>
/// <see cref="DocxReviewService"/> — the dependency-injection mirror of
/// <see cref="DocToolkit.DocxReview"/>, which landed one release behind it because the extensions
/// package builds against the PUBLISHED core.
///
/// <b>This service is pure delegation, so the risk is not that the logic is wrong — it is that a
/// member delegates to the WRONG thing, or silently does nothing.</b> A test asserting only "it did
/// not throw" would miss both, and the DI package is held at 100% coverage precisely because an
/// uncovered member here is a member nobody checked was wired to anything.
///
/// So every assertion below is one a passthrough would fail:
///
/// <list type="bullet">
/// <item><c>Inspect</c> reads a comment and a tracked change out of a document that HAS them — a
/// stub returning an empty report passes a clean-document test and fails this one.</item>
/// <item><c>AcceptRevisions</c> and <c>RejectRevisions</c> are asserted on OPPOSITE outcomes, so
/// one implementation cannot satisfy both, and neither can a method that returns its input.</item>
/// <item><c>RemoveComments</c> is asserted to have removed them.</item>
/// <item>every guard is asserted to throw, which a member doing nothing would not.</item>
/// </list>
///
/// The fixtures are built with raw OpenXml — available transitively through
/// <c>Ank.DocToolkit</c>, so no new dependency — because the library cannot author either state:
/// <c>TrackChanges</c> records nothing for its own edits, and there is no comment-authoring API.
/// </summary>
public class DocxReviewServiceTests
{
    private static readonly DateTime Reviewed = new(2026, 8, 24, 9, 0, 0, DateTimeKind.Utc);

    private static IDocxReview Service() => new DocxReviewService();

    // ---- Inspect ---------------------------------------------------------------------------

    [Fact]
    public void Inspect_ReadsTheReviewStateRatherThanReportingNothing()
    {
        var report = Service().Inspect(WithACommentAndTwoRevisions());

        // A stub returning an empty report would pass a clean-document test. It fails here.
        Assert.Equal("Reviewer One", Assert.Single(report.Comments).Author);
        Assert.Equal(1, report.CommentThreadCount);
        Assert.Equal(2, report.Revisions.Count);
    }

    [Fact]
    public async Task InspectAsync_ReadsTheSameStateAndLeavesTheStreamOpen()
    {
        using var source = new MemoryStream(WithACommentAndTwoRevisions());

        var report = await Service().InspectAsync(source);

        Assert.Equal("Reviewer One", Assert.Single(report.Comments).Author);
        Assert.Equal(2, report.Revisions.Count);

        // DocToolkit did not open this stream and must not close it. A member that disposed the
        // caller's stream would throw on the next read.
        source.Position = 0;
        Assert.Equal(1, source.ReadByte() >= 0 ? 1 : 0);
    }

    // ---- RemoveComments --------------------------------------------------------------------

    [Fact]
    public void RemoveComments_ActuallyRemovesThem()
    {
        byte[] input = WithACommentAndTwoRevisions();
        Assert.NotEmpty(Service().Inspect(input).Comments);          // the fixture really has one

        byte[] cleaned = Service().RemoveComments(input);

        Assert.Empty(Service().Inspect(cleaned).Comments);
    }

    [Fact]
    public async Task RemoveCommentsAsync_WritesACleanedDocumentToTheDestination()
    {
        using var source = new MemoryStream(WithACommentAndTwoRevisions());
        using var destination = new MemoryStream();

        await Service().RemoveCommentsAsync(source, destination);

        // Something was written, and what was written is a document with no comments left.
        Assert.True(destination.Length > 0);
        Assert.Empty(Service().Inspect(destination.ToArray()).Comments);
    }

    // ---- Accept and reject, asserted on OPPOSITE outcomes ----------------------------------

    [Fact]
    public void AcceptRevisions_KeepsTheInsertionAndDropsTheDeletion()
    {
        byte[] accepted = Service().AcceptRevisions(WithACommentAndTwoRevisions());

        Assert.Empty(Service().Inspect(accepted).Revisions);

        string text = DocToolkit.DocxEditor.ExtractText(accepted);
        Assert.Contains("Inserted", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Removed", text, StringComparison.Ordinal);
    }

    [Fact]
    public void RejectRevisions_DropsTheInsertionAndRestoresTheDeletion()
    {
        // THE MIRROR. Wire both members to the same static method and one of these two fails.
        byte[] rejected = Service().RejectRevisions(WithACommentAndTwoRevisions());

        Assert.Empty(Service().Inspect(rejected).Revisions);

        string text = DocToolkit.DocxEditor.ExtractText(rejected);
        Assert.DoesNotContain("Inserted", text, StringComparison.Ordinal);
        Assert.Contains("Removed", text, StringComparison.Ordinal);
    }

    [Fact]
    public async Task AcceptRevisionsAsync_AndRejectRevisionsAsync_DisagreeTheSameWay()
    {
        Assert.Contains("Inserted", await ApplyAsync(Service().AcceptRevisionsAsync), StringComparison.Ordinal);
        Assert.DoesNotContain("Inserted", await ApplyAsync(Service().RejectRevisionsAsync), StringComparison.Ordinal);
    }

    private static async Task<string> ApplyAsync(
        Func<Stream, Stream, CancellationToken, Task> operation)
    {
        using var source = new MemoryStream(WithACommentAndTwoRevisions());
        using var destination = new MemoryStream();
        await operation(source, destination, CancellationToken.None);
        return DocToolkit.DocxEditor.ExtractText(destination.ToArray());
    }

    // ---- Guards ----------------------------------------------------------------------------

    [Fact]
    public void EveryByteArrayMember_RefusesNullAndEmpty()
    {
        var service = Service();
        var calls = new Action<byte[]>[]
        {
            b => service.Inspect(b),
            b => service.RemoveComments(b),
            b => service.AcceptRevisions(b),
            b => service.RejectRevisions(b),
        };

        foreach (var call in calls)
        {
            // A member that did nothing would return rather than throw, so this discriminates.
            Assert.Throws<ArgumentNullException>(() => call(null!));
            Assert.Throws<ArgumentException>(() => call([]));
        }
    }

    [Fact]
    public async Task EveryStreamMember_RefusesAnUnusableStream()
    {
        var service = Service();
        using var ok = new MemoryStream(WithACommentAndTwoRevisions());

        // The destinations are declared rather than passed inline so each is disposed. An
        // undisposed MemoryStream holds no OS handle and would be reclaimed either way - the
        // reason to bother is that cs/local-not-disposed once stood at 61 alerts here, all in
        // test code, and a signal nobody reads protects nothing.
        using var d1 = new MemoryStream();
        using var d2 = new MemoryStream();
        using var d3 = new MemoryStream();

        await Assert.ThrowsAsync<ArgumentNullException>(() => service.InspectAsync(null!));
        await Assert.ThrowsAsync<ArgumentNullException>(() => service.RemoveCommentsAsync(null!, d1));
        await Assert.ThrowsAsync<ArgumentNullException>(() => service.AcceptRevisionsAsync(null!, d2));
        await Assert.ThrowsAsync<ArgumentNullException>(() => service.RejectRevisionsAsync(null!, d3));
        await Assert.ThrowsAsync<ArgumentNullException>(() => service.RejectRevisionsAsync(ok, null!));
    }

    // ---- Fixture ---------------------------------------------------------------------------

    /// <summary>
    /// One comment, one inserted run and one deleted run — enough for every assertion above, in a
    /// single document so the tests share one fixture.
    ///
    /// Hand-assembled because the library cannot produce either state: <c>TrackChanges</c> records
    /// nothing for edits it makes, and <c>AddComment</c> needs a paragraph the placeholder-based API
    /// does not expose. A fixture built through <c>DocxEditor</c> would assert against empty sets
    /// forever while looking green.
    /// </summary>
    private static byte[] WithACommentAndTwoRevisions()
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();

            var paragraph = new Paragraph(
                new CommentRangeStart { Id = "1" },
                new Run(new Text("Kept. ") { Space = SpaceProcessingModeValues.Preserve }),
                new CommentRangeEnd { Id = "1" },
                new Run(new CommentReference { Id = "1" }));

            paragraph.Append(new InsertedRun(
                new Run(new Text("Inserted text. ") { Space = SpaceProcessingModeValues.Preserve }))
            {
                Id = "2",
                Author = "Reviewer Two",
                Date = new DateTimeValue(Reviewed),
            });

            paragraph.Append(new DeletedRun(new Run(new DeletedText("Removed text.")))
            {
                Id = "3",
                Author = "Reviewer Three",
                Date = new DateTimeValue(Reviewed),
            });

            main.Document = new Document(new Body(paragraph));

            var comments = main.AddNewPart<WordprocessingCommentsPart>();
            comments.Comments = new Comments(
                new Comment(new Paragraph(new Run(new Text("Please tighten this."))))
                {
                    Id = "1",
                    Author = "Reviewer One",
                    Initials = "R1",
                    Date = new DateTimeValue(Reviewed),
                });

            main.Document.Save();
        }

        return ms.ToArray();
    }
}
