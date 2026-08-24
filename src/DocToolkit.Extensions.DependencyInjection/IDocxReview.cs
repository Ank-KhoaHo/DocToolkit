namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Reads and resolves a Word document's review state — the comments and tracked changes it carries
/// from having been through review. Registered by
/// <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.
/// </summary>
/// <remarks>
/// <b>Separate from <see cref="IDocxEditor"/> deliberately.</b> That interface is about a document's
/// content; this is about the record of people arguing over it.
///
/// <b>Read and resolve only.</b> There is no member here to add a comment or to record an edit as a
/// tracked change — the core library cannot create review state, so neither can this.
/// </remarks>
public interface IDocxReview
{
    /// <summary>Reads the comments and tracked changes <paramref name="docx"/> carries.</summary>
    /// <param name="docx">The document to inspect.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be opened or read.</exception>
    DocToolkit.DocxReviewReport Inspect(byte[] docx);

    /// <inheritdoc cref="Inspect(byte[])"/>
    /// <param name="source">The document to inspect. Read to its end; never disposed or sought.</param>
    /// <param name="ct">Cancels before the document is read.</param>
    Task<DocToolkit.DocxReviewReport> InspectAsync(Stream source, CancellationToken ct = default);

    /// <summary>A copy of <paramref name="docx"/> with every comment removed.</summary>
    /// <param name="docx">The document to clean.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be read or written.</exception>
    byte[] RemoveComments(byte[] docx);

    /// <inheritdoc cref="RemoveComments(byte[])"/>
    /// <param name="source">The document to clean. Read to its end; never disposed or sought.</param>
    /// <param name="destination">Receives the cleaned document. Written; never disposed or sought.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    Task RemoveCommentsAsync(Stream source, Stream destination, CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="docx"/> with every tracked change applied: insertions kept,
    /// deletions dropped. <b>Cannot be undone from the result.</b>
    /// </summary>
    /// <param name="docx">The document to apply changes to.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be read or written.</exception>
    byte[] AcceptRevisions(byte[] docx);

    /// <inheritdoc cref="AcceptRevisions(byte[])"/>
    /// <param name="source">The document to apply changes to. Read to its end; never disposed or sought.</param>
    /// <param name="destination">Receives the result. Written; never disposed or sought.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    Task AcceptRevisionsAsync(Stream source, Stream destination, CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="docx"/> with every tracked change discarded: insertions dropped,
    /// deletions restored — the mirror of <see cref="AcceptRevisions(byte[])"/>. <b>Cannot be undone
    /// from the result.</b>
    /// </summary>
    /// <param name="docx">The document to discard changes from.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be read or written.</exception>
    byte[] RejectRevisions(byte[] docx);

    /// <inheritdoc cref="RejectRevisions(byte[])"/>
    /// <param name="source">The document to discard changes from. Read to its end; never disposed or sought.</param>
    /// <param name="destination">Receives the result. Written; never disposed or sought.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    Task RejectRevisionsAsync(Stream source, Stream destination, CancellationToken ct = default);
}
