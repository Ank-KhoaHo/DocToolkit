namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Operations on a PDF that already exists — page count, merge, page extraction and document
/// information. Registered by <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.
/// </summary>
/// <remarks>
/// The only injectable service here that <b>reads</b> a PDF; every other one writes or renders.
/// Nothing on it re-renders, so pages keep the text, fonts and images they arrived with.
/// </remarks>
public interface IPdfEditor
{
    /// <summary>The number of pages in <paramref name="pdf"/>.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="pdf"/> is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The bytes are not a readable PDF.</exception>
    int PageCount(byte[] pdf);

    /// <inheritdoc cref="PageCount(byte[])"/>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    Task<int> PageCountAsync(Stream source, CancellationToken ct = default);

    /// <summary>Joins <paramref name="pdfs"/> into one document, keeping the order given.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="pdfs"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="pdfs"/> is empty; merging nothing would produce a zero-page PDF, which
    /// several readers refuse to open.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">One of the inputs is not a readable PDF.</exception>
    byte[] Merge(IEnumerable<byte[]> pdfs);

    /// <inheritdoc cref="Merge(IEnumerable{byte[]})"/>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    Task MergeAsync(IEnumerable<Stream> sources, Stream destination, CancellationToken ct = default);

    /// <summary>
    /// A new document holding <paramref name="count"/> pages starting at <paramref name="firstPage"/>.
    /// </summary>
    /// <param name="pdf">The document to take pages out of. It is not modified.</param>
    /// <param name="firstPage">1-based, because that is how a reader numbers pages.</param>
    /// <param name="count">How many pages to take, starting at <paramref name="firstPage"/>.</param>
    /// <exception cref="ArgumentNullException"><paramref name="pdf"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The range is not entirely inside the document.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The bytes are not a readable PDF.</exception>
    byte[] ExtractPages(byte[] pdf, int firstPage, int count);

    /// <inheritdoc cref="ExtractPages(byte[], int, int)"/>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    Task ExtractPagesAsync(
        Stream source, int firstPage, int count, Stream destination, CancellationToken ct = default);

    /// <summary>The document information <paramref name="pdf"/> carries.</summary>
    /// <remarks>
    /// An absent entry reads back as <see langword="null"/> rather than an empty string, so "no
    /// title" stays distinguishable from "a title deliberately set to empty".
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="pdf"/> is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The bytes are not a readable PDF.</exception>
    DocToolkit.PdfMetadata ReadMetadata(byte[] pdf);

    /// <summary>A copy of <paramref name="pdf"/> carrying <paramref name="metadata"/>.</summary>
    /// <remarks>
    /// A <see langword="null"/> property leaves what the document already had in place, so stamping
    /// a title does not silently erase an author. Pass an empty string to clear one.
    /// </remarks>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The bytes are not a readable PDF.</exception>
    byte[] WithMetadata(byte[] pdf, DocToolkit.PdfMetadata metadata);
}
