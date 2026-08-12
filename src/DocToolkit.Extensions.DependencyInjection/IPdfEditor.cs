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

    /// <summary>
    /// A new document with <paramref name="count"/> pages removed, starting at
    /// <paramref name="firstPage"/>. The complement of <see cref="ExtractPages(byte[], int, int)"/>:
    /// that one keeps the range, this one keeps everything else.
    /// </summary>
    /// <param name="pdf">The document to take pages out of. It is not modified.</param>
    /// <param name="firstPage">1-based, because that is how a reader numbers pages.</param>
    /// <param name="count">How many pages to drop, starting at <paramref name="firstPage"/>.</param>
    /// <exception cref="ArgumentNullException"><paramref name="pdf"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The range is not entirely inside the document, or it covers every page - removing
    /// everything would leave a zero-page file.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The bytes are not a readable PDF.</exception>
    byte[] RemovePages(byte[] pdf, int firstPage, int count);

    /// <inheritdoc cref="RemovePages(byte[], int, int)"/>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    Task RemovePagesAsync(
        Stream source, int firstPage, int count, Stream destination, CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="pdf"/> with <paramref name="count"/> pages turned clockwise by
    /// <paramref name="degrees"/>, starting at <paramref name="firstPage"/>.
    /// </summary>
    /// <param name="pdf">The document to turn pages in. It is not modified.</param>
    /// <param name="firstPage">1-based, because that is how a reader numbers pages.</param>
    /// <param name="count">How many pages to turn, starting at <paramref name="firstPage"/>.</param>
    /// <param name="degrees">
    /// How far to turn, clockwise, as a multiple of 90. Negative turns anticlockwise.
    ///
    /// <b>This is relative, not absolute</b> — it adds to whatever rotation the page already
    /// carries, so calling it twice with 90 leaves the page at 180.
    /// </param>
    /// <exception cref="ArgumentNullException"><paramref name="pdf"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The range is not entirely inside the document, or <paramref name="degrees"/> is not a
    /// multiple of 90.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The bytes are not a readable PDF.</exception>
    byte[] RotatePages(byte[] pdf, int firstPage, int count, int degrees);

    /// <inheritdoc cref="RotatePages(byte[], int, int, int)"/>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    Task RotatePagesAsync(
        Stream source, int firstPage, int count, int degrees, Stream destination,
        CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="pdf"/> with its pages in the order given by
    /// <paramref name="order"/>, which holds 1-based page numbers.
    /// </summary>
    /// <param name="pdf">The document to reorder. It is not modified.</param>
    /// <param name="order">
    /// A <b>permutation of every page</b> — the same pages, in a different order. Not a subset, and
    /// no repeats.
    /// </param>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="order"/> is not a permutation of 1..<c>PageCount</c>.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The bytes are not a readable PDF.</exception>
    byte[] ReorderPages(byte[] pdf, IEnumerable<int> order);

    /// <inheritdoc cref="ReorderPages(byte[], IEnumerable{int})"/>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    Task ReorderPagesAsync(
        Stream source, IEnumerable<int> order, Stream destination, CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="target"/> with every page of <paramref name="source"/> inserted so
    /// that the first of them becomes page <paramref name="atPage"/>.
    /// </summary>
    /// <param name="target">The document to insert into. It is not modified.</param>
    /// <param name="source">The document whose pages are inserted. It is not modified.</param>
    /// <param name="atPage">
    /// 1-based position the first inserted page will occupy. <c>PageCount + 1</c> appends, which is
    /// deliberately allowed - it is the obvious way to say "after everything".
    /// </param>
    /// <exception cref="ArgumentNullException">Either document is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="atPage"/> is below 1 or more than one past the last page.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The bytes are not a readable PDF.</exception>
    byte[] InsertPages(byte[] target, byte[] source, int atPage);

    /// <inheritdoc cref="InsertPages(byte[], byte[], int)"/>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    Task InsertPagesAsync(
        Stream target, Stream source, int atPage, Stream destination, CancellationToken ct = default);
}
