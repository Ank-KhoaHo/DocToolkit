namespace DocToolkit;

/// <summary>What a document carries from having been through review.</summary>
/// <remarks>
/// One report rather than separate reads, because the underlying walk computes comments and
/// revisions together — two calls could return counts that disagree because they came from
/// different passes over the same document.
/// </remarks>
public sealed class DocxReviewReport
{
    internal DocxReviewReport(
        IReadOnlyList<DocxComment> comments,
        IReadOnlyList<DocxRevision> revisions,
        int commentThreadCount,
        int unresolvedThreadCount)
    {
        Comments = comments;
        Revisions = revisions;
        CommentThreadCount = commentThreadCount;
        UnresolvedThreadCount = unresolvedThreadCount;
    }

    /// <summary>
    /// The comments that start a thread, in document order. A reply is <b>not</b> listed here; it
    /// appears on its parent's <see cref="DocxComment.Replies"/>, so no comment is reported twice.
    /// </summary>
    public IReadOnlyList<DocxComment> Comments { get; }

    /// <summary>Every tracked change, in document order.</summary>
    public IReadOnlyList<DocxRevision> Revisions { get; }

    /// <summary>
    /// Comment threads, counting a comment and its replies as one. Equal to
    /// <see cref="Comments"/>.Count; both are kept because the counts below are what a caller
    /// usually wants and reading them should not require walking the list.
    /// </summary>
    public int CommentThreadCount { get; }

    /// <summary>
    /// Threads nobody has marked resolved — the number worth acting on, and the reason to prefer
    /// this over a bare comment count when deciding whether a document is finished with.
    /// </summary>
    public int UnresolvedThreadCount { get; }
}

/// <summary>A single comment, together with any replies to it.</summary>
public sealed class DocxComment
{
    internal DocxComment(
        string author, string initials, string text, DateTime? created, bool? isResolved,
        IReadOnlyList<DocxComment> replies)
    {
        Author = author;
        Initials = initials;
        Text = text;
        Created = created;
        IsResolved = isResolved;
        Replies = replies;
    }

    /// <summary>Who wrote it. Empty when the document records no author.</summary>
    public string Author { get; }

    /// <summary>Their initials, as Word stores them. Empty when the document records none.</summary>
    public string Initials { get; }

    /// <summary>The comment's text.</summary>
    public string Text { get; }

    /// <summary><see langword="null"/> when the document records no date, which is common.</summary>
    public DateTime? Created { get; }

    /// <summary>
    /// Whether the thread has been marked resolved. <see langword="null"/> when the document says
    /// nothing either way — which is not the same as unresolved, and is what a document written by
    /// a tool that does not track resolution looks like.
    /// </summary>
    public bool? IsResolved { get; }

    /// <summary>Replies to this comment, in order. Empty for a comment nobody answered.</summary>
    public IReadOnlyList<DocxComment> Replies { get; }
}

/// <summary>What kind of change a revision records.</summary>
/// <remarks>
/// <b><see cref="Other"/> is the ordinary case, not an exotic one.</b> Word records eleven kinds of
/// revision and this models the two that describe content changing; the remaining nine — the two
/// halves of a move, and formatting changes to runs, paragraphs, sections, tables, rows and cells —
/// all arrive as <see cref="Other"/>. A document edited with track-changes on while somebody
/// restyled it can therefore be entirely <see cref="Other"/>.
///
/// Reporting them coarsely beats refusing to read the document, and beats throwing on a kind added
/// by a later Word. Naming more of them later is additive.
/// </remarks>
public enum DocxRevisionKind
{
    /// <summary>Anything that is neither an insertion nor a deletion — see the remarks.</summary>
    Other = 0,

    /// <summary>Text somebody added.</summary>
    Insertion = 1,

    /// <summary>Text somebody removed.</summary>
    Deletion = 2,
}

/// <summary>A single tracked change.</summary>
public sealed class DocxRevision
{
    internal DocxRevision(
        DocxRevisionKind kind, string author, string affectedText, DateTime? created, bool isInTable)
    {
        Kind = kind;
        Author = author;
        AffectedText = affectedText;
        Created = created;
        IsInTable = isInTable;
    }

    /// <summary>What the change did.</summary>
    public DocxRevisionKind Kind { get; }

    /// <summary>Who made it. Empty when the document records no author.</summary>
    public string Author { get; }

    /// <summary>
    /// The text the change inserted or removed. Empty for a formatting revision, which changes how
    /// text looks rather than what it says.
    /// </summary>
    public string AffectedText { get; }

    /// <summary><see langword="null"/> when the document records no date.</summary>
    public DateTime? Created { get; }

    /// <summary>
    /// Worth knowing before applying a change in bulk: accepting or rejecting inside a table can
    /// add or remove rows and cells, so it moves more than text.
    /// </summary>
    public bool IsInTable { get; }
}
