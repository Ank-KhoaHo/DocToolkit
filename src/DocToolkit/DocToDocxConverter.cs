using OfficeIMO;
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc.Diagnostics;

namespace DocToolkit;

/// <summary>
/// Reads a Word 97-2003 binary document (.doc) — the format Word used before .docx — and converts
/// it to a .docx package, or reads its text directly.
/// </summary>
/// <remarks>
/// <b>Import only.</b> There is no .doc <i>writing</i> here and there will not be: the underlying
/// library reports native .doc saving as unsupported, so offering it would mean claiming something
/// that does not work.
///
/// <b>Converting refuses by default when the source holds content a .docx cannot carry.</b> A
/// legacy .doc keeps pictures, drawings and form fields in a binary stream that the import can see
/// but cannot project. Rather than quietly hand back a document missing them,
/// <see cref="Convert(byte[])"/> throws — see <see cref="LegacyDocOptions.AllowContentLoss"/> to
/// accept the loss on purpose, and <see cref="ConvertWithReport(byte[], LegacyDocOptions?)"/> to
/// record exactly what it was.
///
/// <b><see cref="ExtractText(byte[])"/> never refuses and takes no options</b>, because text is not
/// what that binary stream holds. Reading a .doc someone sent you is the common case and it does
/// not need a policy decision.
///
/// Measured 2026-08-16 against documents produced by Word itself: text, tables (every cell) and
/// character formatting such as bold survive conversion intact.
/// </remarks>
public static class DocToDocxConverter
{
    private const string FailureMessage =
        "Failed to convert the legacy .doc to DOCX. See the inner exception for details.";

    private const string ExtractFailureMessage =
        "Failed to read the legacy .doc. See the inner exception for details.";

    // Named on DocToolkit's own option, not on WordSaveOptions.LossPolicy. The upstream message
    // names types a consumer of this package cannot reach, which is the D18 failure shape: advice
    // that cannot be acted on is worse than none.
    private const string ContentLossMessage =
        "This .doc holds content a .docx cannot carry - pictures, drawings or form fields, kept in " +
        "the legacy binary stream. Text, tables and formatting would convert; those payloads would " +
        "not. Pass a LegacyDocOptions with AllowContentLoss set to true to accept that, or use " +
        "ConvertWithReport to see exactly what is dropped.";

    private const string NotALegacyDocMessage =
        "The bytes are not a Word 97-2003 binary document. A .docx is a different format that only " +
        "looks similar - if this file came from a modern Word, use DocxEditor or DocxToPdfConverter " +
        "instead. This converter reads the pre-2007 binary .doc format.";

    /// <summary>Converts the legacy .doc in <paramref name="doc"/> to a .docx package.</summary>
    /// <remarks>
    /// Throws when the source holds content the .docx cannot carry. Use
    /// <see cref="Convert(byte[], LegacyDocOptions?)"/> to accept that loss deliberately.
    /// </remarks>
    /// <param name="doc">The Word 97-2003 binary document to convert.</param>
    /// <exception cref="ArgumentNullException"><paramref name="doc"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="doc"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">
    /// The document could not be converted, or it holds content a .docx cannot carry and
    /// <see cref="LegacyDocOptions.AllowContentLoss"/> was not set.
    /// </exception>
    public static byte[] Convert(byte[] doc) => ConvertCore(doc, options: null).Value;

    /// <inheritdoc cref="Convert(byte[])" path="/summary|/remarks|/exception"/>
    /// <param name="doc">The Word 97-2003 binary document to convert.</param>
    /// <param name="options">
    /// How to treat content the .docx cannot carry. <see langword="null"/> means the default:
    /// refuse.
    /// </param>
    public static byte[] Convert(byte[] doc, LegacyDocOptions? options) => ConvertCore(doc, options).Value;

    /// <summary>
    /// Converts the legacy .doc in <paramref name="doc"/> and reports what the import could not
    /// carry across.
    /// </summary>
    /// <remarks>
    /// Returns exactly the bytes <see cref="Convert(byte[], LegacyDocOptions?)"/> returns for the
    /// same input and options — the conversion runs once, and the report is read off the same
    /// loaded document rather than from a second pass.
    ///
    /// The report is worth reading even on a document that converts without the opt-in: an import
    /// can be lossless and still have something to say, such as quick-save revision history that is
    /// readable but is not carried across as editable revisions.
    /// </remarks>
    /// <inheritdoc cref="Convert(byte[], LegacyDocOptions?)" path="/exception"/>
    /// <param name="doc">The Word 97-2003 binary document to convert.</param>
    /// <param name="options">
    /// How to treat content the .docx cannot carry. <see langword="null"/> means the default:
    /// refuse.
    /// </param>
    public static ConversionResult<byte[]> ConvertWithReport(byte[] doc, LegacyDocOptions? options = null) =>
        ConvertCore(doc, options);

    private static ConversionResult<byte[]> ConvertCore(byte[] doc, LegacyDocOptions? options)
    {
        ArgumentNullException.ThrowIfNull(doc);
        if (doc.Length == 0)
            throw new ArgumentException("DOC content was empty.", nameof(doc));

        using var word = Open(doc, nameof(doc));

        var warnings = Describe(word);
        var save = new WordSaveOptions
        {
            LossPolicy = options?.AllowContentLoss == true
                ? OfficeConversionLossPolicy.Allow
                : OfficeConversionLossPolicy.Block,
        };

        using var output = new MemoryStream();
        try
        {
            word.Save(output, save);
        }
        // NotSupportedException is how the loss block surfaces, and it is the ONLY thing that
        // arrives as one here - so translating it to a message naming AllowContentLoss cannot
        // swallow an unrelated failure. Anything else keeps the generic wrapper below.
        catch (NotSupportedException ex)
        {
            throw new DocumentConversionException(ContentLossMessage, ex);
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException(FailureMessage, ex);
        }

        return new ConversionResult<byte[]>(output.ToArray(), warnings);
    }

    /// <summary>
    /// Reads the text of the legacy .doc in <paramref name="doc"/>, including the contents of table
    /// cells.
    /// </summary>
    /// <remarks>
    /// <b>Takes no options and never refuses over content loss</b>, unlike the <c>Convert</c>
    /// overloads: what a .doc's binary stream holds is pictures, drawings and form fields, none of
    /// which are text, so there is nothing for a loss policy to decide.
    ///
    /// Blocks are separated the way <see cref="DocxEditor.ExtractText(byte[])"/> separates them, so
    /// adjacent paragraphs do not fuse into one word.
    /// </remarks>
    /// <param name="doc">The Word 97-2003 binary document to read.</param>
    /// <exception cref="ArgumentNullException"><paramref name="doc"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="doc"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The document could not be read.</exception>
    public static string ExtractText(byte[] doc)
    {
        ArgumentNullException.ThrowIfNull(doc);
        if (doc.Length == 0)
            throw new ArgumentException("DOC content was empty.", nameof(doc));

        using var word = Open(doc, nameof(doc));
        try
        {
            var blocks = word.Paragraphs
                .Select(p => p.Text)
                .Where(t => !string.IsNullOrEmpty(t))
                .ToList();

            // Table cells are NOT in Paragraphs, so reading only the list above would drop every
            // table silently. One line per row, cells tab-separated, matching DocxEditor.
            var rows = word.Tables
                .SelectMany(t => t.Rows)
                .Select(r => string.Join("\t", r.Cells
                    .Select(c => string.Concat(c.Paragraphs.Select(p => p.Text)))
                    .Where(cell => !string.IsNullOrEmpty(cell))))
                .Where(line => line.Length > 0);

            blocks.AddRange(rows);
            return string.Join("\n", blocks);
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException(ExtractFailureMessage, ex);
        }
    }

    /// <summary>
    /// Reads a legacy .doc from <paramref name="source"/> and writes the converted .docx package to
    /// <paramref name="destination"/>.
    ///
    /// Neither stream is disposed, closed or sought, so <paramref name="source"/> may be
    /// forward-only — an HTTP request body, for instance.
    /// </summary>
    /// <inheritdoc cref="Convert(byte[], LegacyDocOptions?)" path="/exception"/>
    /// <param name="source">The stream the .doc is read from.</param>
    /// <param name="destination">The stream the .docx package is written to.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or
    /// <paramref name="destination"/> is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    public static Task ConvertAsync(Stream source, Stream destination, CancellationToken ct = default) =>
        ConvertAsync(source, destination, options: null, ct);

    /// <inheritdoc cref="ConvertAsync(Stream, Stream, CancellationToken)" path="/summary|/exception"/>
    /// <param name="source">The stream the .doc is read from.</param>
    /// <param name="destination">The stream the .docx package is written to.</param>
    /// <param name="options">
    /// How to treat content the .docx cannot carry. <see langword="null"/> means the default:
    /// refuse.
    /// </param>
    /// <param name="ct">Cancels the read and the write.</param>
    public static async Task ConvertAsync(
        Stream source, Stream destination, LegacyDocOptions? options, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var input = await StreamPipeline
            .DrainAsync(source, "DOC content was empty.", nameof(source), FailureMessage, ct)
            .ConfigureAwait(false);

        var docx = Convert(input.ToArray(), options);
        using var scratch = new MemoryStream(docx, writable: false);
        await StreamPipeline.EmitAsync(scratch, destination, FailureMessage, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads a legacy .doc from <paramref name="source"/> and returns its text.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and is not disposed, closed or sought.
    /// </summary>
    /// <inheritdoc cref="ExtractText(byte[])" path="/exception"/>
    /// <param name="source">The stream the .doc is read from.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    public static async Task<string> ExtractTextAsync(Stream source, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var input = await StreamPipeline
            .DrainAsync(source, "DOC content was empty.", nameof(source), ExtractFailureMessage, ct)
            .ConfigureAwait(false);

        return ExtractText(input.ToArray());
    }

    /// <summary>
    /// Opens the legacy document, translating "this is not a .doc" into a message that says what to
    /// do about it. Shared so both entry points answer that mistake identically.
    /// </summary>
    private static WordDocument Open(byte[] doc, string paramName)
    {
        _ = paramName;

        // Not writable: nothing here edits the source, and the import reads it whole.
        using var input = new MemoryStream(doc, writable: false);
        try
        {
            return WordDocument.LoadLegacyDoc(input);
        }
        // The import throws InvalidDataException for a .docx and for arbitrary bytes alike -
        // verified 2026-08-16 with both as negative controls.
        catch (InvalidDataException ex)
        {
            throw new DocumentConversionException(NotALegacyDocMessage, ex);
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException(FailureMessage, ex);
        }
    }

    /// <summary>
    /// Maps the import's own diagnostics onto DocToolkit's warning type.
    /// </summary>
    /// <remarks>
    /// Two upstream lists are folded into one: <c>LegacyDocImportDiagnostics</c> (informational,
    /// such as quick-save history) and <c>LegacyDocCompoundFeatures</c> (the payloads that could
    /// not be projected). A caller wants one answer to "what did I lose", not two lists whose
    /// difference is an upstream implementation detail.
    /// </remarks>
    private static List<ConversionWarning> Describe(WordDocument word)
    {
        var warnings = new List<ConversionWarning>();

        foreach (var d in word.LegacyDocImportDiagnostics)
        {
            // Severity, not loss: a diagnostic describes the import, and only an Error means
            // something was actually lost. Info and Warning are reported as carrying no loss so a
            // caller filtering on Kind is not misled by ordinary chatter.
            var kind = d.Severity == LegacyDocDiagnosticSeverity.Error
                ? OfficeConversionLossKind.Failure
                : OfficeConversionLossKind.None;
            warnings.Add(ConversionDiagnostics.Warning(d.Code, d.Message, kind));
        }

        foreach (var f in word.LegacyDocCompoundFeatures)
        {
            // These are the payloads kept in the source but not projected, so Omission is exact.
            var message = $"{f.Description} ({f.EntryCount} entry/entries, {f.TotalBytes} bytes).";
            warnings.Add(ConversionDiagnostics.Warning(f.Code, message, OfficeConversionLossKind.Omission));
        }

        foreach (var f in word.LegacyDocUnsupportedFeatures)
            warnings.Add(ConversionDiagnostics.Warning(f.Code, f.Description, OfficeConversionLossKind.Omission));

        return warnings;
    }
}
