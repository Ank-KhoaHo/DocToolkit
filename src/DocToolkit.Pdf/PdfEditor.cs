using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;

namespace DocToolkit;

/// <summary>
/// Operations on a PDF that already exists: how many pages it has, joining several into one,
/// taking a range of pages out, and reading or stamping its document information.
/// </summary>
/// <remarks>
/// This is the only part of the library that READS a PDF. Everything else here writes one —
/// <c>DocxToPdfConverter</c> and friends render into PDF and never look at the result,
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

        // Reads the bytes rather than handing a FileStream to the Stream overload, so an empty
        // file is reported against `path` and not against that overload's `source`. Nothing is
        // lost by it: the Stream overload drains its source into an array before doing anything,
        // so this was never the streaming path it looked like.
        return PageCount(await FilePipeline.ReadAsync(path, nameof(path), ct).ConfigureAwait(false));
    }

    /// <summary>
    /// Each page's text, in document order. <c>[0]</c> is page 1.
    /// </summary>
    /// <remarks>
    /// The index is a list position, not a page number — deliberately unlike this class's
    /// <c>firstPage</c> parameters, which are 1-based because that is how a reader numbers pages.
    /// <c>PresentationEditor.ExtractText</c> returns per slide for the same reason.
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

        return ExtractText(await FilePipeline.ReadAsync(path, nameof(path), ct).ConfigureAwait(false));
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

    /// <summary>
    /// A copy of <paramref name="pdf"/> encrypted with the passwords and permissions in
    /// <paramref name="protection"/>.
    /// </summary>
    /// <remarks>
    /// <b>Set <see cref="PdfProtection.UserPassword"/> if the content must not be read.</b> An
    /// owner password alone leaves the document openable by anyone — the permissions are a request
    /// the reader is asked to honour, not a lock. <see cref="PdfProtection"/> explains the
    /// distinction; getting it wrong is the usual way a "protected" PDF turns out not to be.
    ///
    /// The result cannot be passed back into the other operations on this class: they refuse an
    /// encrypted document, by design. Use <see cref="Unprotect(byte[], string)"/> first.
    /// </remarks>
    /// <param name="pdf">The document to encrypt.</param>
    /// <param name="protection">The passwords and permissions to apply.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="pdf"/> or <paramref name="protection"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="pdf"/> is empty, or <paramref name="protection"/> sets neither password.
    /// </exception>
    /// <exception cref="DocumentConversionException">The document could not be read or written.</exception>
    public static byte[] Protect(byte[] pdf, PdfProtection protection)
    {
        ArgumentNullException.ThrowIfNull(pdf);
        ArgumentNullException.ThrowIfNull(protection);

        var user = protection.UserPassword;
        var owner = protection.OwnerPassword;

        // Checked here rather than left to PDFsharp, which throws a bare PdfSharpException reading
        // "At least a user or an owner password is required to encrypt the document." That is true
        // but says nothing about which one a caller actually wants, and the difference is the whole
        // point of the type.
        if (string.IsNullOrEmpty(user) && string.IsNullOrEmpty(owner))
            throw new ArgumentException(
                "Encrypting a PDF needs at least one password. Set UserPassword to stop the "
                + "document being opened at all, or OwnerPassword to leave it readable while asking "
                + "readers to honour the permissions.",
                nameof(protection));

        using var document = Open(pdf, PdfDocumentOpenMode.Modify, nameof(pdf));

        // Strength must be chosen BEFORE the passwords: SetEncryption resets the handler, so
        // calling it afterwards discards them and silently writes an unencrypted document.
        document.SecurityHandler.SetEncryption(protection.Strength == PdfEncryptionStrength.Aes256
            ? PdfDefaultEncryption.V5
            : PdfDefaultEncryption.V4UsingAES);

        var settings = document.SecuritySettings;
        if (!string.IsNullOrEmpty(user)) settings.UserPassword = user;
        if (!string.IsNullOrEmpty(owner)) settings.OwnerPassword = owner;

        settings.PermitPrint = protection.AllowPrinting;
        settings.PermitFullQualityPrint = protection.AllowHighQualityPrinting;
        settings.PermitExtractContent = protection.AllowCopying;
        settings.PermitModifyDocument = protection.AllowModification;
        settings.PermitAnnotations = protection.AllowAnnotations;
        settings.PermitFormsFill = protection.AllowFormFilling;
        settings.PermitAssembleDocument = protection.AllowAssembly;

        return Save(document);
    }

    /// <summary>
    /// A copy of <paramref name="pdf"/> with its encryption removed, so the rest of this class can
    /// work on it.
    /// </summary>
    /// <remarks>
    /// <b>This is the one method here that accepts an encrypted document</b>, and that is
    /// deliberate: threading a password through all nine other operations would put a security
    /// parameter on methods that have nothing to do with security. Unprotect once, then use the
    /// result normally.
    ///
    /// <b>Which password is needed is not always the obvious one.</b> If the document has an owner
    /// password, that is the one required here — even if you also know the user password. Removing
    /// protection is a <i>modification</i>, and the PDF format reserves modification for the owner.
    /// A document protected with only a user password takes that one. Measured 2026-08-16; the
    /// exception message names all three ways this can go wrong rather than guessing between them.
    ///
    /// <b>The output is not protected in any way</b> — that is what was asked for, but it means the
    /// bytes this returns are readable by anyone who obtains them.
    /// </remarks>
    /// <param name="pdf">The encrypted document.</param>
    /// <param name="password">The user or owner password.</param>
    /// <exception cref="ArgumentNullException"><paramref name="pdf"/> or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pdf"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">
    /// The password was wrong, or the document could not be read or written.
    /// </exception>
    public static byte[] Unprotect(byte[] pdf, string password)
    {
        ArgumentNullException.ThrowIfNull(pdf);
        ArgumentNullException.ThrowIfNull(password);

        using var document = Open(pdf, PdfDocumentOpenMode.Modify, nameof(pdf), password);
        document.SecurityHandler.SetEncryption(PdfDefaultEncryption.None);

        return Save(document);
    }

    /// <summary>
    /// Reads a PDF from <paramref name="source"/> and writes the encrypted copy to
    /// <paramref name="destination"/>.
    ///
    /// Neither stream is disposed, closed or sought.
    /// </summary>
    /// <inheritdoc cref="Protect(byte[], PdfProtection)" path="/remarks|/exception"/>
    /// <param name="source">The stream the PDF is read from.</param>
    /// <param name="destination">The stream the encrypted PDF is written to.</param>
    /// <param name="protection">The passwords and permissions to apply.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    public static async Task ProtectAsync(
        Stream source, Stream destination, PdfProtection protection, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ArgumentNullException.ThrowIfNull(protection);

        var pdf = await ReadAsync(source, nameof(source), ct).ConfigureAwait(false);
        await WriteAsync(Protect(pdf, protection), destination, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads an encrypted PDF from <paramref name="source"/> and writes the unprotected copy to
    /// <paramref name="destination"/>.
    ///
    /// Neither stream is disposed, closed or sought.
    /// </summary>
    /// <inheritdoc cref="Unprotect(byte[], string)" path="/remarks|/exception"/>
    /// <param name="source">The stream the encrypted PDF is read from.</param>
    /// <param name="destination">The stream the unprotected PDF is written to.</param>
    /// <param name="password">The user or owner password.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    public static async Task UnprotectAsync(
        Stream source, Stream destination, string password, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ArgumentNullException.ThrowIfNull(password);

        var pdf = await ReadAsync(source, nameof(source), ct).ConfigureAwait(false);
        await WriteAsync(Unprotect(pdf, password), destination, ct).ConfigureAwait(false);
    }

    private static string? Entry(PdfDocumentInformation info, string key) =>
        info.Elements.ContainsKey(key) ? info.Elements.GetString(key) : null;

    /// <summary>
    /// Opens a document, translating every way PDFsharp can refuse into this library's one
    /// exception type — a caller should not have to catch a different type for PDFs than for
    /// everything else here.
    /// </summary>
    // `password` is null for every caller but Unprotect, and that default is load-bearing: a null
    // password means the open REFUSES an encrypted PDF, which is what every other operation on
    // this class relies on to avoid quietly working on a document it could not really read.
    private static PdfDocument Open(
        byte[] pdf, PdfDocumentOpenMode mode, string paramName, string? password = null)
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
            return password is null
                ? PdfReader.Open(source, mode)
                : PdfReader.Open(source, password, mode);
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            // A caller who supplied a password already knows the document is encrypted, so
            // repeating "it may be password-protected" at them is noise — and it hides the one
            // thing that actually went wrong.
            // A permission-restricted PDF - an owner password, no user password - READS perfectly
            // and refuses only modification. Saying "failed to read" to somebody whose PageCount
            // and ExtractText calls just succeeded sends them to check bytes that are fine, and
            // names three causes when the real one is a fourth.
            //
            // Measured 2026-08-17 across 200 real PDFs from a .gov crawl: 11 of them (5.5%) are
            // exactly this, from 1 page to 1000. It is the single most common write failure on
            // real-world input, and the remedy - Unprotect with the owner password - shipped in
            // 0.28.0 and went unmentioned.
            //
            // Detected by open MODE rather than by matching the upstream message: a read mode that
            // failed is a genuinely unreadable document, while Modify/Import failing on a document
            // that reads is the restriction. Matching message text would break the moment PDFsharp
            // reworded it, which is the failure this repository has recorded elsewhere.
            var restricted = mode != PdfDocumentOpenMode.Import
                             && password is null
                             && CanOpenForReading(pdf);

            throw new DocumentConversionException(
                restricted
                    ? "This PDF is permission-restricted: it carries an owner password, so it can "
                      + "be read but not modified. Reading it works - PageCount and ExtractText are "
                      + "unaffected. To change it, call PdfEditor.Unprotect with the owner password "
                      + "first."
                : password is null
                    ? "Failed to read the PDF. This usually means the PDF is password-protected, "
                      + "truncated, or not actually a PDF — check the source bytes."
                    // Three candidates, named rather than guessed between. The third is the one
                    // that costs an afternoon: a correct USER password is refused here, because
                    // removing protection is a modification and the PDF specification reserves
                    // that for the owner password. Asserting any single cause would be wrong at
                    // least a third of the time - the same mistake as a timeout message that
                    // names a cause the timeout cannot distinguish.
                    : "Failed to read the PDF with the password supplied. Either the password is "
                      + "wrong, or the document is not encrypted and needs no password, or it was "
                      + "the user password when the document also has an owner password — removing "
                      + "protection needs the owner password, because that is what the PDF format "
                      + "requires to modify a document.",
                ex);
        }
    }

    /// <summary>
    /// Whether the document opens for READING without a password - which distinguishes a
    /// permission-restricted PDF from an unreadable one. Import is the read mode this class uses
    /// everywhere; PdfDocumentOpenMode.ReadOnly is obsolete in PDFsharp and not implemented.
    /// </summary>
    /// <remarks>
    /// Only ever called on a path that has already failed, so its cost is paid on the error branch
    /// and never on a successful call. It swallows everything deliberately: this is a question
    /// being asked to choose a message, and a failure to answer it just means the ordinary message
    /// is used.
    /// </remarks>
    private static bool CanOpenForReading(byte[] pdf)
    {
        try
        {
            using var probe = new MemoryStream(pdf, writable: false);
            using var _ = PdfReader.Open(probe, PdfDocumentOpenMode.Import);
            return true;
        }
        catch
        {
            return false;
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
            throw new DocumentConversionException("Failed to write the PDF. See the inner exception for details.", ex);
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
