namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IDocxReview"/>, delegating to <see cref="DocToolkit.DocxReview"/>.</summary>
internal sealed class DocxReviewService : IDocxReview
{
    public DocToolkit.DocxReviewReport Inspect(byte[] docx) => DocToolkit.DocxReview.Inspect(docx);

    public Task<DocToolkit.DocxReviewReport> InspectAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.DocxReview.InspectAsync(source, ct);

    public byte[] RemoveComments(byte[] docx) => DocToolkit.DocxReview.RemoveComments(docx);

    public Task RemoveCommentsAsync(Stream source, Stream destination, CancellationToken ct = default)
        => DocToolkit.DocxReview.RemoveCommentsAsync(source, destination, ct);

    public byte[] AcceptRevisions(byte[] docx) => DocToolkit.DocxReview.AcceptRevisions(docx);

    public Task AcceptRevisionsAsync(Stream source, Stream destination, CancellationToken ct = default)
        => DocToolkit.DocxReview.AcceptRevisionsAsync(source, destination, ct);

    public byte[] RejectRevisions(byte[] docx) => DocToolkit.DocxReview.RejectRevisions(docx);

    public Task RejectRevisionsAsync(Stream source, Stream destination, CancellationToken ct = default)
        => DocToolkit.DocxReview.RejectRevisionsAsync(source, destination, ct);
}
