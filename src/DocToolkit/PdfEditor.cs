using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DocToolkit;

/// <summary>
/// Operations on a PDF that already exists: how many pages it has, joining several into one,
/// taking a range of pages out, and reading or stamping its document information.
/// </summary>
/// <remarks>
/// This is the only part of the library that READS a PDF. Everything else here writes one —
/// <see cref="DocxToPdfConverter"/> and friends render into PDF and never look at the result,
/// which is why the test suite had to carry a hand-rolled parser before this existed.
///
/// Nothing here re-renders. Pages are moved between documents as they are, so text, fonts and
/// images arrive unchanged and the fidelity caveats that apply to the converters do not apply to
/// these operations.
/// </remarks>
public static class PdfEditor
{
    /// <summary>The number of pages in <paramref name="pdf"/>.</summary>
    /// <exception cref="DocumentConversionException">The bytes are not a readable PDF.</exception>
    public static int PageCount(byte[] pdf)
    {
        ArgumentNullException.ThrowIfNull(pdf);

        using var document = Open(pdf, PdfDocumentOpenMode.Import);
        return document.PageCount;
    }

    /// <inheritdoc cref="PageCount(byte[])"/>
    public static async Task<int> PageCountAsync(Stream source, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(source);

        return PageCount(await ReadAllAsync(source, ct).ConfigureAwait(false));
    }

    /// <inheritdoc cref="PageCount(byte[])"/>
    public static async Task<int> PageCountAsync(string path, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrEmpty(path);

        await using var source = File.OpenRead(path);
        return await PageCountAsync(source, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Joins <paramref name="pdfs"/> into one document, keeping the order given.
    /// </summary>
    /// <exception cref="ArgumentException">
    /// <paramref name="pdfs"/> is empty. A zero-page PDF is not a useful artefact and several
    /// readers refuse to open one, so this fails rather than returning something shaped like a
    /// document.
    /// </exception>
    public static byte[] Merge(IEnumerable<byte[]> pdfs)
    {
        ArgumentNullException.ThrowIfNull(pdfs);

        var sources = pdfs.ToArray();
        if (sources.Length == 0)
        {
            throw new ArgumentException(
                "At least one document is required; merging nothing would produce a zero-page PDF.",
                nameof(pdfs));
        }

        using var merged = new PdfDocument();

        foreach (var bytes in sources)
        {
            ArgumentNullException.ThrowIfNull(bytes, nameof(pdfs));

            using var input = Open(bytes, PdfDocumentOpenMode.Import);
            for (var page = 0; page < input.PageCount; page++)
            {
                merged.AddPage(input.Pages[page]);
            }
        }

        return Save(merged);
    }

    /// <inheritdoc cref="Merge(IEnumerable{byte[]})"/>
    public static async Task MergeAsync(
        IEnumerable<Stream> sources, Stream destination, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(sources);
        ArgumentNullException.ThrowIfNull(destination);

        var documents = new List<byte[]>();
        foreach (var source in sources)
        {
            documents.Add(await ReadAllAsync(source, ct).ConfigureAwait(false));
        }

        await destination.WriteAsync(Merge(documents), ct).ConfigureAwait(false);
    }

    /// <summary>
    /// A new document holding <paramref name="count"/> pages starting at <paramref name="firstPage"/>.
    /// </summary>
    /// <param name="pdf">The document to take pages out of. It is not modified.</param>
    /// <param name="firstPage">1-based, because that is how a reader numbers pages.</param>
    /// <param name="count">How many pages to take, starting at <paramref name="firstPage"/>.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The range is not entirely inside the document. Checked as a whole rather than per argument:
    /// a start inside the document and a count that runs off the end is the mistake worth catching,
    /// and neither argument is wrong on its own.
    /// </exception>
    public static byte[] ExtractPages(byte[] pdf, int firstPage, int count)
    {
        ArgumentNullException.ThrowIfNull(pdf);
        ArgumentOutOfRangeException.ThrowIfLessThan(firstPage, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(count, 1);

        using var input = Open(pdf, PdfDocumentOpenMode.Import);

        if (firstPage + count - 1 > input.PageCount)
        {
            throw new ArgumentOutOfRangeException(
                nameof(count), count,
                $"Pages {firstPage}-{firstPage + count - 1} were requested from a document with "
                + $"{input.PageCount} page(s).");
        }

        using var extracted = new PdfDocument();
        for (var offset = 0; offset < count; offset++)
        {
            extracted.AddPage(input.Pages[firstPage - 1 + offset]);
        }

        return Save(extracted);
    }

    /// <inheritdoc cref="ExtractPages(byte[], int, int)"/>
    public static async Task ExtractPagesAsync(
        Stream source, int firstPage, int count, Stream destination, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(destination);

        var pdf = await ReadAllAsync(source, ct).ConfigureAwait(false);
        await destination.WriteAsync(ExtractPages(pdf, firstPage, count), ct).ConfigureAwait(false);
    }

    /// <summary>The document information <paramref name="pdf"/> carries.</summary>
    public static PdfMetadata ReadMetadata(byte[] pdf)
    {
        ArgumentNullException.ThrowIfNull(pdf);

        using var document = Open(pdf, PdfDocumentOpenMode.Import);
        var info = document.Info;

        // Read through the underlying dictionary rather than the typed properties: those return
        // string.Empty for a key that is not there, which would erase the difference between "no
        // title" and "a title that is deliberately blank".
        return new PdfMetadata
        {
            Title = Entry(info, "/Title"),
            Author = Entry(info, "/Author"),
            Subject = Entry(info, "/Subject"),
            Keywords = Entry(info, "/Keywords"),
            Creator = Entry(info, "/Creator"),
        };
    }

    /// <summary>
    /// A copy of <paramref name="pdf"/> carrying <paramref name="metadata"/>.
    /// </summary>
    /// <remarks>
    /// A <see langword="null"/> property leaves what the document already had in place, so stamping
    /// a title does not silently erase an author. Pass an empty string to clear one.
    /// </remarks>
    public static byte[] WithMetadata(byte[] pdf, PdfMetadata metadata)
    {
        ArgumentNullException.ThrowIfNull(pdf);
        ArgumentNullException.ThrowIfNull(metadata);

        using var document = Open(pdf, PdfDocumentOpenMode.Modify);

        if (metadata.Title is not null) document.Info.Title = metadata.Title;
        if (metadata.Author is not null) document.Info.Author = metadata.Author;
        if (metadata.Subject is not null) document.Info.Subject = metadata.Subject;
        if (metadata.Keywords is not null) document.Info.Keywords = metadata.Keywords;
        if (metadata.Creator is not null) document.Info.Creator = metadata.Creator;

        return Save(document);
    }

    private static string? Entry(PdfDocumentInformation info, string key) =>
        info.Elements.ContainsKey(key) ? info.Elements.GetString(key) : null;

    /// <summary>
    /// Opens a document, translating every way PDFsharp can refuse into this library's one
    /// exception type — a caller should not have to catch a different type for PDFs than for
    /// everything else here.
    /// </summary>
    private static PdfDocument Open(byte[] pdf, PdfDocumentOpenMode mode)
    {
        try
        {
            return PdfReader.Open(new MemoryStream(pdf, writable: false), mode);
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read the PDF.", ex);
        }
    }

    private static byte[] Save(PdfDocument document)
    {
        try
        {
            using var buffer = new MemoryStream();
            document.Save(buffer, closeStream: false);
            return buffer.ToArray();
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to write the PDF.", ex);
        }
    }

    private static async Task<byte[]> ReadAllAsync(Stream source, CancellationToken ct)
    {
        ArgumentNullException.ThrowIfNull(source);

        if (source is MemoryStream ready)
        {
            return ready.ToArray();
        }

        using var buffer = new MemoryStream();
        await source.CopyToAsync(buffer, ct).ConfigureAwait(false);
        return buffer.ToArray();
    }
}
