using OfficeIMOCommentInfo = OfficeIMO.Word.WordCommentInfo;
using OfficeIMORevisionType = OfficeIMO.Word.WordReviewRevisionType;
using OfficeIMOWordDocument = OfficeIMO.Word.WordDocument;

namespace DocToolkit;

/// <summary>
/// Reads and resolves a document's review state — the comments and tracked changes a .docx carries
/// from having been through review.
/// </summary>
/// <remarks>
/// <b>Separate from <see cref="DocxEditor"/> deliberately.</b> That class is about a document's
/// content; this is about the record of people arguing over it. The two are read and acted on at
/// different times by different callers, and folding four more operations into an already-large
/// class would have made both harder to read.
///
/// <b>This class reads and resolves review state; it cannot create any.</b> There is no method
/// here to add a comment or to record an edit as a tracked change — see the A66 backlog row for
/// why authoring needs a targeting design of its own.
/// </remarks>
public static class DocxReview
{
    /// <summary>
    /// Reads the comments and tracked changes <paramref name="docx"/> carries.
    /// </summary>
    /// <remarks>
    /// One call rather than one per kind, because the underlying walk computes both together —
    /// separate reads could return counts that disagree because they came from different passes
    /// over the same document. A document nobody has reviewed reports empty rather than failing.
    /// </remarks>
    /// <param name="docx">The document to inspect.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The document could not be opened or read.</exception>
    public static DocxReviewReport Inspect(byte[] docx)
    {
        RequireContent(docx);

        using var source = new MemoryStream(docx, writable: false);
        return InspectCore(source);
    }

    /// <summary>
    /// Reads the comments and tracked changes the .docx in <paramref name="source"/> carries.
    /// <paramref name="source"/> is <b>read</b> to its end and is neither disposed, closed nor
    /// sought.
    /// </summary>
    /// <inheritdoc cref="Inspect(byte[])" path="/remarks"/>
    /// <param name="source">The document to inspect.</param>
    /// <param name="ct">Cancels before the document is read.</param>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be opened or read.</exception>
    public static async Task<DocxReviewReport> InspectAsync(Stream source, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, EmptySource, nameof(source), ReadFailure, ct)
            .ConfigureAwait(false);

        return InspectCore(docx);
    }

    /// <summary>
    /// A copy of <paramref name="docx"/> with every comment removed.
    /// </summary>
    /// <remarks>
    /// The text a comment was anchored to is left alone — only the comment and its anchor go. This
    /// is what "clear the review notes before sending it out" means, and it is not an edit to what
    /// the document says.
    /// </remarks>
    /// <param name="docx">The document to clean.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The document could not be read or written.</exception>
    public static byte[] RemoveComments(byte[] docx)
        => Apply(docx, static d => d.RemoveAllComments(), RemoveFailure);

    /// <summary>
    /// Writes a copy of the .docx in <paramref name="source"/> to <paramref name="destination"/>
    /// with every comment removed. <paramref name="source"/> is <b>read</b> to its end and
    /// <paramref name="destination"/> is <b>written</b>; neither is disposed, closed nor sought.
    /// </summary>
    /// <inheritdoc cref="RemoveComments(byte[])" path="/remarks"/>
    /// <param name="source">The document to clean.</param>
    /// <param name="destination">Receives the cleaned document.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> or <paramref name="destination"/> is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, or <paramref name="source"/> held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be read or written.</exception>
    public static Task RemoveCommentsAsync(Stream source, Stream destination, CancellationToken ct = default)
        => ApplyAsync(source, destination, static d => d.RemoveAllComments(), RemoveFailure, ct);

    /// <summary>
    /// A copy of <paramref name="docx"/> with every tracked change applied and its markup removed.
    /// </summary>
    /// <remarks>
    /// Insertions become ordinary text and deletions go, which is what a reviewer means by "accept
    /// all". The result carries no revisions at all, so <see cref="Inspect(byte[])"/> on it reports
    /// none.
    ///
    /// <b>This changes what the document says, and the result cannot be undone.</b> An accepted
    /// deletion is not recoverable from the bytes this returns; keep the original if that matters.
    /// </remarks>
    /// <param name="docx">The document to apply changes to.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The document could not be read or written.</exception>
    public static byte[] AcceptRevisions(byte[] docx)
        => Apply(docx, static d => d.AcceptRevisions(), AcceptFailure);

    /// <summary>
    /// Writes a copy of the .docx in <paramref name="source"/> to <paramref name="destination"/>
    /// with every tracked change applied. <paramref name="source"/> is <b>read</b> to its end and
    /// <paramref name="destination"/> is <b>written</b>; neither is disposed, closed nor sought.
    /// </summary>
    /// <inheritdoc cref="AcceptRevisions(byte[])" path="/remarks"/>
    /// <param name="source">The document to apply changes to.</param>
    /// <param name="destination">Receives the resulting document.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> or <paramref name="destination"/> is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, or <paramref name="source"/> held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be read or written.</exception>
    public static Task AcceptRevisionsAsync(Stream source, Stream destination, CancellationToken ct = default)
        => ApplyAsync(source, destination, static d => d.AcceptRevisions(), AcceptFailure, ct);

    /// <summary>
    /// A copy of <paramref name="docx"/> with every tracked change discarded and its markup
    /// removed.
    /// </summary>
    /// <remarks>
    /// Insertions go and deletions are put back, restoring what the document said before the
    /// review — the mirror of <see cref="AcceptRevisions(byte[])"/>. The result carries no
    /// revisions at all.
    ///
    /// <b>This changes what the document says, and the result cannot be undone.</b> Every
    /// insertion is gone from the bytes this returns; keep the original if that matters.
    /// </remarks>
    /// <param name="docx">The document to discard changes from.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The document could not be read or written.</exception>
    public static byte[] RejectRevisions(byte[] docx)
        => Apply(docx, static d => d.RejectRevisions(), RejectFailure);

    /// <summary>
    /// Writes a copy of the .docx in <paramref name="source"/> to <paramref name="destination"/>
    /// with every tracked change discarded. <paramref name="source"/> is <b>read</b> to its end and
    /// <paramref name="destination"/> is <b>written</b>; neither is disposed, closed nor sought.
    /// </summary>
    /// <inheritdoc cref="RejectRevisions(byte[])" path="/remarks"/>
    /// <param name="source">The document to discard changes from.</param>
    /// <param name="destination">Receives the resulting document.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> or <paramref name="destination"/> is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, or <paramref name="source"/> held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be read or written.</exception>
    public static Task RejectRevisionsAsync(Stream source, Stream destination, CancellationToken ct = default)
        => ApplyAsync(source, destination, static d => d.RejectRevisions(), RejectFailure, ct);

    private const string EmptySource = "DOCX content was empty.";
    private const string ReadFailure = "Failed to read the document's review state. See the inner exception for details.";
    private const string RemoveFailure = "Failed to remove the document's comments. See the inner exception for details.";
    private const string AcceptFailure = "Failed to accept the document's tracked changes. See the inner exception for details.";
    private const string RejectFailure = "Failed to reject the document's tracked changes. See the inner exception for details.";

    private static void RequireContent(byte[] docx)
    {
        ArgumentNullException.ThrowIfNull(docx);
        if (docx.Length == 0)
            throw new ArgumentException(EmptySource, nameof(docx));
    }

    private static byte[] Apply(byte[] docx, Action<OfficeIMOWordDocument> apply, string failure)
    {
        RequireContent(docx);

        using var source = new MemoryStream(docx, writable: false);
        using var result = ApplyCore(source, apply, failure);
        return result.ToArray();
    }

    private static async Task ApplyAsync(
        Stream source, Stream destination, Action<OfficeIMOWordDocument> apply, string failure,
        CancellationToken ct)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, EmptySource, nameof(source), failure, ct)
            .ConfigureAwait(false);

        using var result = ApplyCore(docx, apply, failure);
        await StreamPipeline.EmitAsync(result, destination, failure, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// The one implementation behind all six mutating overloads, so a <c>byte[]</c> call and a
    /// <c>Stream</c> call can never drift apart.
    /// </summary>
    /// <remarks>
    /// <b>The catch is unfiltered, and it disposes.</b> This method hands its buffer to its caller,
    /// so it owns that buffer until it returns — and a <i>filtered</i> catch that does not match
    /// never runs its body, which is exactly how a <see cref="MemoryStream"/> once escaped this
    /// repository with an exception. A method whose disposables are all under <c>using</c> may
    /// filter; this one may not.
    ///
    /// There is deliberately no second arm re-throwing <see cref="DocumentConversionException"/>
    /// unwrapped. Nothing inside this <c>try</c> can raise one — every call in it belongs to
    /// OfficeIMO — so such an arm would be a branch no test could ever reach, and unreachable
    /// defensive code is indistinguishable from a guard that stopped working.
    /// </remarks>
    private static MemoryStream ApplyCore(
        Stream source, Action<OfficeIMOWordDocument> apply, string failure)
    {
        var result = new MemoryStream();
        try
        {
            using var document = OfficeIMOWordDocument.Load(source);
            apply(document);
            document.Save(result);
            result.Position = 0;
            return result;
        }
        catch (Exception ex)
        {
            result.Dispose();
            throw new DocumentConversionException(failure, ex);
        }
    }

    private static DocxReviewReport InspectCore(Stream source)
    {
        try
        {
            using var document = OfficeIMOWordDocument.Load(source);
            var report = document.InspectReviewReport();

            // Mapped from CommentThreads, NEVER from the flat Comments list. That list carries each
            // reply as an entry of its own as well as under its parent, so mapping it into a DTO
            // that also nests Replies would report every reply twice.
            var comments = report.CommentThreads
                .Select(thread => new DocxComment(
                    Text(thread.Parent.Author),
                    Text(thread.Parent.Initials),
                    Text(thread.Parent.Text),
                    thread.Parent.DateTime,
                    thread.IsResolved,
                    thread.Replies.Select(Reply).ToArray()))
                .ToArray();

            var revisions = report.Revisions
                .Select(r => new DocxRevision(
                    Kind(r.RevisionType),
                    Text(r.Author),
                    Text(r.AffectedText),
                    r.DateTime,
                    r.IsInTable))
                .ToArray();

            return new DocxReviewReport(
                comments, revisions, report.CommentThreadCount, report.UnresolvedThreadCount);
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException(ReadFailure, ex);
        }
    }

    private static DocxComment Reply(OfficeIMOCommentInfo reply)
        => new(Text(reply.Author), Text(reply.Initials), Text(reply.Text), reply.DateTime,
               reply.IsResolved, []);

    /// <summary>
    /// Nine of the eleven revision types collapse onto <see cref="DocxRevisionKind.Other"/> — see
    /// that enum's remarks. Mapping rather than throwing means a document full of formatting
    /// revisions, or one carrying a kind a later Word introduces, still reads.
    /// </summary>
    private static DocxRevisionKind Kind(OfficeIMORevisionType type) => type switch
    {
        OfficeIMORevisionType.Insertion => DocxRevisionKind.Insertion,
        OfficeIMORevisionType.Deletion => DocxRevisionKind.Deletion,
        _ => DocxRevisionKind.Other,
    };

    /// <summary>
    /// A missing author, initials or text is reported as empty rather than null, so a caller
    /// formatting a report never has to null-check a string this returns.
    /// </summary>
    private static string Text(string? value) => value ?? string.Empty;
}
