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

        using var document = Open(pdf, PdfDocumentOpenMode.Import, nameof(pdf));
        return document.PageCount;
    }

    /// <inheritdoc cref="PageCount(byte[])"/>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes.
    /// </exception>
    public static async Task<int> PageCountAsync(Stream source, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));

        return PageCount(await ReadAsync(source, nameof(source), ct).ConfigureAwait(false));
    }

    /// <inheritdoc cref="PageCount(byte[])"/>
    public static async Task<int> PageCountAsync(string path, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        // `await using var source = ...` would leave the DISPOSAL await unconfigured, which is a
        // second await the declaration form gives no place to put ConfigureAwait on. Configuring
        // the disposable and scoping it with a block is the only way to reach it.
        var source = File.OpenRead(path);
        await using (source.ConfigureAwait(false))
        {
            return await PageCountAsync(source, ct).ConfigureAwait(false);
        }
    }

    /// <summary>
    /// Each page's text, in document order. <c>[0]</c> is page 1.
    /// </summary>
    /// <remarks>
    /// The index is a list position, not a page number — deliberately unlike this class's
    /// <c>firstPage</c> parameters, which are 1-based because that is how a reader numbers pages.
    /// <see cref="PresentationEditor.ExtractText(byte[])"/> returns per slide for the same reason.
    ///
    /// <b>A page with no text layer returns an empty string.</b> A scanned document is images, so
    /// this returns one empty string per page for one — that is what the file contains, not a
    /// failure, and OCR is out of scope. This is the commonest surprise in PDF text extraction, so
    /// it is stated here rather than left to be discovered.
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="pdf"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pdf"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">
    /// The bytes are not a readable PDF, or it requires a password to open. A PDF that is merely
    /// permission-restricted (an owner password with no user password) is not covered by that:
    /// PdfPig opens it with its default empty password and this returns its text like any other
    /// PDF - measured, not assumed.
    /// </exception>
    public static IReadOnlyList<string> ExtractText(byte[] pdf)
    {
        ArgumentNullException.ThrowIfNull(pdf);
        if (pdf.Length == 0)
            throw new ArgumentException("PDF content was empty.", nameof(pdf));

        return PdfTextExtractor.Pages(pdf);
    }

    /// <inheritdoc cref="ExtractText(byte[])"/>
    /// <remarks>
    /// <paramref name="source"/> is read to its end; it is not disposed, closed or sought, and does
    /// not have to be seekable.
    ///
    /// This <c>remarks</c> replaces the one on <see cref="ExtractText(byte[])"/> rather than adding
    /// to it, so its warning is restated rather than assumed to carry over: <b>a page with no text
    /// layer returns an empty string.</b> A scanned document is images, so this returns one empty
    /// string per page for one - that is what the file contains, not a failure.
    /// </remarks>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable.</exception>
    public static async Task<IReadOnlyList<string>> ExtractTextAsync(
        Stream source, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));

        return PdfTextExtractor.Pages(await ReadAsync(source, nameof(source), ct).ConfigureAwait(false));
    }

    /// <inheritdoc cref="ExtractText(byte[])"/>
    public static async Task<IReadOnlyList<string>> ExtractTextAsync(
        string path, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        return ExtractText(await File.ReadAllBytesAsync(path, ct).ConfigureAwait(false));
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

            using var input = Open(bytes, PdfDocumentOpenMode.Import, nameof(bytes));
            for (var page = 0; page < input.PageCount; page++)
            {
                merged.AddPage(input.Pages[page]);
            }
        }

        return Save(merged);
    }

    /// <inheritdoc cref="Merge(IEnumerable{byte[]})"/>
    /// <exception cref="ArgumentException">
    /// <paramref name="sources"/> is empty, or holds a stream that is not readable or held no
    /// bytes; or <paramref name="destination"/> is not writable.
    /// </exception>
    public static async Task MergeAsync(
        IEnumerable<Stream> sources, Stream destination, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(sources);
        StreamPipeline.RequireWritable(destination, nameof(destination));

        var documents = new List<byte[]>();
        foreach (var source in sources)
        {
            StreamPipeline.RequireReadable(source, nameof(sources));
            documents.Add(await ReadAsync(source, nameof(sources), ct).ConfigureAwait(false));
        }

        // Checked here rather than left to Merge, which would name its own `pdfs` parameter — a
        // parameter this caller never passed and cannot see.
        if (documents.Count == 0)
        {
            throw new ArgumentException(
                "At least one document is required; merging nothing would produce a zero-page PDF.",
                nameof(sources));
        }

        await WriteAsync(Merge(documents), destination, ct).ConfigureAwait(false);
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

        using var input = Open(pdf, PdfDocumentOpenMode.Import, nameof(pdf));

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
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or
    /// <paramref name="destination"/> is not writable.
    /// </exception>
    public static async Task ExtractPagesAsync(
        Stream source, int firstPage, int count, Stream destination, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));

        var pdf = await ReadAsync(source, nameof(source), ct).ConfigureAwait(false);
        await WriteAsync(ExtractPages(pdf, firstPage, count), destination, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// A new document with <paramref name="count"/> pages removed, starting at
    /// <paramref name="firstPage"/>. The complement of <see cref="ExtractPages(byte[], int, int)"/>:
    /// that one keeps the range, this one keeps everything else.
    /// </summary>
    /// <param name="pdf">The document to take pages out of. It is not modified.</param>
    /// <param name="firstPage">1-based, because that is how a reader numbers pages.</param>
    /// <param name="count">How many pages to drop, starting at <paramref name="firstPage"/>.</param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The range is not entirely inside the document, or it covers every page. Removing everything
    /// would leave a zero-page file, which is not a document any reader will open — refusing is more
    /// useful than returning it. The range is checked as a whole for the same reason as
    /// <see cref="ExtractPages(byte[], int, int)"/>: a start inside the document with a count that
    /// runs off the end is the mistake worth catching, and neither argument is wrong on its own.
    /// </exception>
    public static byte[] RemovePages(byte[] pdf, int firstPage, int count)
    {
        ArgumentNullException.ThrowIfNull(pdf);
        ArgumentOutOfRangeException.ThrowIfLessThan(firstPage, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(count, 1);

        using var input = Open(pdf, PdfDocumentOpenMode.Import, nameof(pdf));

        if (firstPage + count - 1 > input.PageCount)
        {
            throw new ArgumentOutOfRangeException(
                nameof(count), count,
                $"Pages {firstPage}-{firstPage + count - 1} were requested for removal from a "
                + $"document with {input.PageCount} page(s).");
        }

        if (count >= input.PageCount)
        {
            throw new ArgumentOutOfRangeException(
                nameof(count), count,
                $"Removing {count} of {input.PageCount} page(s) would leave nothing. A zero-page "
                + "PDF is not a document.");
        }

        using var kept = new PdfDocument();
        for (var page = 1; page <= input.PageCount; page++)
        {
            if (page >= firstPage && page < firstPage + count) continue;
            kept.AddPage(input.Pages[page - 1]);
        }

        return Save(kept);
    }

    /// <inheritdoc cref="RemovePages(byte[], int, int)"/>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or
    /// <paramref name="destination"/> is not writable.
    /// </exception>
    public static async Task RemovePagesAsync(
        Stream source, int firstPage, int count, Stream destination, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));

        var pdf = await ReadAsync(source, nameof(source), ct).ConfigureAwait(false);
        await WriteAsync(RemovePages(pdf, firstPage, count), destination, ct).ConfigureAwait(false);
    }

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
    /// carries, so calling it twice with 90 leaves the page at 180. That is the operation people
    /// actually want ("this scan came out sideways, turn it"), and an absolute setter would both
    /// leak PDF's own <c>/Rotate</c> model into the API and silently do nothing on a page that
    /// already held the value asked for.
    /// </param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// The range is not entirely inside the document, or <paramref name="degrees"/> is not a
    /// multiple of 90. The PDF specification requires <c>/Rotate</c> to be a quarter turn, so
    /// accepting 45 would write a file readers disagree about rather than fail.
    /// </exception>
    public static byte[] RotatePages(byte[] pdf, int firstPage, int count, int degrees)
    {
        ArgumentNullException.ThrowIfNull(pdf);
        ArgumentOutOfRangeException.ThrowIfLessThan(firstPage, 1);
        ArgumentOutOfRangeException.ThrowIfLessThan(count, 1);

        if (degrees % 90 != 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(degrees), degrees,
                "A PDF page rotation must be a multiple of 90 degrees.");
        }

        using var document = Open(pdf, PdfDocumentOpenMode.Modify, nameof(pdf));

        if (firstPage + count - 1 > document.PageCount)
        {
            throw new ArgumentOutOfRangeException(
                nameof(count), count,
                $"Pages {firstPage}-{firstPage + count - 1} were requested for rotation from a "
                + $"document with {document.PageCount} page(s).");
        }

        for (var offset = 0; offset < count; offset++)
        {
            var page = document.Pages[firstPage - 1 + offset];

            // Normalised into [0, 360) because PdfSharp rejects anything outside it, and because a
            // page that has been turned four times should read as upright rather than as 360.
            page.Rotate = ((page.Rotate + degrees) % 360 + 360) % 360;
        }

        return Save(document);
    }

    /// <inheritdoc cref="RotatePages(byte[], int, int, int)"/>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or
    /// <paramref name="destination"/> is not writable.
    /// </exception>
    public static async Task RotatePagesAsync(
        Stream source, int firstPage, int count, int degrees, Stream destination,
        CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));

        var pdf = await ReadAsync(source, nameof(source), ct).ConfigureAwait(false);
        await WriteAsync(RotatePages(pdf, firstPage, count, degrees), destination, ct)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// A copy of <paramref name="pdf"/> with its pages in the order given by
    /// <paramref name="order"/>, which holds 1-based page numbers.
    /// </summary>
    /// <param name="pdf">The document to reorder. It is not modified.</param>
    /// <param name="order">
    /// A <b>permutation of every page</b> — the same pages, in a different order. Not a subset, and
    /// no repeats. A "reorder" that quietly dropped a page would be the worst kind of bug here,
    /// because the result still looks like a document; taking a subset is what
    /// <see cref="ExtractPages(byte[], int, int)"/> is for.
    /// </param>
    /// <exception cref="ArgumentException">
    /// <paramref name="order"/> is not a permutation of 1..<c>PageCount</c>. Reported as one
    /// failure rather than per element, because a caller who passed the wrong list wants to know
    /// the list is wrong, not which entry was noticed first.
    /// </exception>
    public static byte[] ReorderPages(byte[] pdf, IEnumerable<int> order)
    {
        ArgumentNullException.ThrowIfNull(pdf);
        ArgumentNullException.ThrowIfNull(order);

        var wanted = order.ToArray();

        using var input = Open(pdf, PdfDocumentOpenMode.Import, nameof(pdf));

        var expected = Enumerable.Range(1, input.PageCount);
        if (!wanted.OrderBy(p => p).SequenceEqual(expected))
        {
            throw new ArgumentException(
                $"The order must be a permutation of pages 1-{input.PageCount}, each exactly once. "
                + $"Got [{string.Join(", ", wanted)}].",
                nameof(order));
        }

        using var reordered = new PdfDocument();
        foreach (var page in wanted)
        {
            reordered.AddPage(input.Pages[page - 1]);
        }

        return Save(reordered);
    }

    /// <inheritdoc cref="ReorderPages(byte[], IEnumerable{int})"/>
    /// <exception cref="ArgumentException">
    /// <paramref name="order"/> is not a permutation of every page, <paramref name="source"/> is
    /// not readable or held no bytes, or <paramref name="destination"/> is not writable.
    /// </exception>
    public static async Task ReorderPagesAsync(
        Stream source, IEnumerable<int> order, Stream destination, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(order);
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));

        var pdf = await ReadAsync(source, nameof(source), ct).ConfigureAwait(false);
        await WriteAsync(ReorderPages(pdf, order), destination, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// A copy of <paramref name="target"/> with every page of <paramref name="source"/> inserted so
    /// that the first of them becomes page <paramref name="atPage"/>.
    /// </summary>
    /// <param name="target">The document to insert into. It is not modified.</param>
    /// <param name="source">The document whose pages are inserted. It is not modified.</param>
    /// <param name="atPage">
    /// 1-based position the first inserted page will occupy. <c>1</c> puts them in front of
    /// everything; <c>PageCount + 1</c> appends, which is deliberately allowed — it is the obvious
    /// way to say "after everything", and rejecting it would leave appending expressible only
    /// through <see cref="Merge(IEnumerable{byte[]})"/>.
    /// </param>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="atPage"/> is below 1 or more than one past the last page.
    /// </exception>
    public static byte[] InsertPages(byte[] target, byte[] source, int atPage)
    {
        ArgumentNullException.ThrowIfNull(target);
        ArgumentNullException.ThrowIfNull(source);
        ArgumentOutOfRangeException.ThrowIfLessThan(atPage, 1);

        using var into = Open(target, PdfDocumentOpenMode.Import, nameof(target));
        using var from = Open(source, PdfDocumentOpenMode.Import, nameof(source));

        if (atPage > into.PageCount + 1)
        {
            throw new ArgumentOutOfRangeException(
                nameof(atPage), atPage,
                $"Cannot insert at page {atPage} of a document with {into.PageCount} page(s); "
                + $"{into.PageCount + 1} appends and is the highest position allowed.");
        }

        using var combined = new PdfDocument();

        for (var page = 0; page < atPage - 1; page++) combined.AddPage(into.Pages[page]);
        for (var page = 0; page < from.PageCount; page++) combined.AddPage(from.Pages[page]);
        for (var page = atPage - 1; page < into.PageCount; page++) combined.AddPage(into.Pages[page]);

        return Save(combined);
    }

    /// <inheritdoc cref="InsertPages(byte[], byte[], int)"/>
    /// <exception cref="ArgumentException">
    /// <paramref name="target"/> or <paramref name="source"/> is not readable or held no bytes, or
    /// <paramref name="destination"/> is not writable. The two sources are named separately, so a
    /// caller who mixed them up learns which one this is about.
    /// </exception>
    public static async Task InsertPagesAsync(
        Stream target, Stream source, int atPage, Stream destination,
        CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(target, nameof(target));
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));

        var into = await ReadAsync(target, nameof(target), ct).ConfigureAwait(false);
        var from = await ReadAsync(source, nameof(source), ct).ConfigureAwait(false);
        await WriteAsync(InsertPages(into, from, atPage), destination, ct).ConfigureAwait(false);
    }

    /// <summary>The document information <paramref name="pdf"/> carries.</summary>
    public static PdfMetadata ReadMetadata(byte[] pdf)
    {
        ArgumentNullException.ThrowIfNull(pdf);

        using var document = Open(pdf, PdfDocumentOpenMode.Import, nameof(pdf));
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

        using var document = Open(pdf, PdfDocumentOpenMode.Modify, nameof(pdf));

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
    private static PdfDocument Open(byte[] pdf, PdfDocumentOpenMode mode, string paramName)
    {
        // An empty array is an argument mistake, not a corrupt document, and this is the one
        // place every byte[] path passes through. Until 2026-08-15 only ExtractText checked it,
        // so PageCount(Array.Empty<byte>()) threw DocumentConversionException while
        // ExtractText(Array.Empty<byte>()) threw ArgumentException - the same input answered two
        // ways by one class. B17 fixed exactly this on the Stream path and left byte[] alone.
        //
        // paramName is threaded in rather than hard-coded to "pdf" because InsertPages takes
        // `target` and `source`, and an exception naming the wrong parameter is its own defect.
        if (pdf.Length == 0)
            throw new ArgumentException("PDF content was empty.", paramName);

        try
        {
            // PdfSharp reads the whole package during Open, so the stream is not needed
            // afterwards and disposing it here is safe - asserted by the PdfEditor suite, which
            // exercises every open mode. Held in a local rather than inlined so it CAN be disposed.
            using var source = new MemoryStream(pdf, writable: false);
            return PdfReader.Open(source, mode);
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

    /// <summary>
    /// The one place a <see cref="Stream"/> overload here reads its input, so the eight of them
    /// cannot disagree about what reading a source means.
    /// </summary>
    /// <remarks>
    /// This used to be a private <c>ReadAllAsync</c> with a <see cref="MemoryStream"/> fast-path
    /// that returned <c>ToArray()</c>. Three things were wrong with it and none was detectable,
    /// because no <see cref="PdfEditor"/> method was registered in <c>StreamOverloadTests</c>:
    /// <c>ToArray()</c> ignores <c>Position</c> and never advances the stream, so the fast path and
    /// the slow path disagreed about where reading starts; the token was not observed at all on
    /// that path; and an empty source came back as an empty array, to be reported later as a
    /// corrupt PDF rather than immediately as the empty argument it was.
    /// </remarks>
    private static async Task<byte[]> ReadAsync(Stream source, string paramName, CancellationToken ct)
    {
        ct.ThrowIfCancellationRequested();

        using var buffer = await StreamPipeline
            .DrainAsync(source, "PDF content was empty.", paramName, "Failed to read the PDF.", ct)
            .ConfigureAwait(false);

        return buffer.ToArray();
    }

    /// <summary>
    /// The one place a <see cref="Stream"/> overload here writes its result.
    /// </summary>
    /// <remarks>
    /// Going through <see cref="StreamPipeline.EmitAsync"/> rather than calling
    /// <c>destination.WriteAsync</c> directly is what makes a failure to write the caller's stream
    /// arrive as this library's own exception type, the way it already does everywhere else.
    /// </remarks>
    private static async Task WriteAsync(byte[] pdf, Stream destination, CancellationToken ct)
    {
        using var buffer = new MemoryStream(pdf, writable: false);

        await StreamPipeline
            .EmitAsync(buffer, destination, "Failed to write the PDF to the destination stream.", ct)
            .ConfigureAwait(false);
    }
}
