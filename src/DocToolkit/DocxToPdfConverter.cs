using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

namespace DocToolkit;

/// <summary>Renders a Word (.docx) package to PDF. Pure managed - no browser, no LibreOffice.</summary>
public static class DocxToPdfConverter
{
    /// <summary>Renders the .docx in <paramref name="docx"/> and returns PDF bytes.</summary>
    /// <example>
    /// <code source="../../tests/DocToolkit.Tests/DocumentationExamples.cs" region="DocxToPdf"/>
    /// </example>
    public static byte[] Convert(byte[] docx) => Convert(docx, null);

    /// <summary>
    /// Renders the .docx in <paramref name="docx"/>, using <paramref name="fonts"/> for characters
    /// the renderer cannot otherwise encode.
    /// </summary>
    /// <remarks>
    /// <b>Whether a document containing non-Latin text renders otherwise depends on the machine.</b>
    /// The renderer falls back to whatever fonts the host happens to have, and a Windows box offers
    /// ones that do not cover Cyrillic - so the same document converts on one machine and is refused
    /// on another. Supplying the font removes the machine from the answer.
    ///
    /// See <see cref="PdfFontOptions"/> for the one side effect worth knowing about.
    /// </remarks>
    /// <param name="docx">The document to render.</param>
    /// <param name="fonts">Fonts to fall back to, or <see langword="null"/> for none.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The document could not be rendered.</exception>
    public static byte[] Convert(byte[] docx, PdfFontOptions? fonts)
    {
        ArgumentNullException.ThrowIfNull(docx);
        if (docx.Length == 0)
            throw new ArgumentException("DOCX content was empty.", nameof(docx));

        try
        {
            // Word's default bullet is a Symbol-font glyph in the private-use area, which the PDF
            // renderer refuses to encode - so a document with a bulleted list could not be
            // rendered at all. Substituted before loading; see ListMarkerSubstitution for why the
            // replacements are measured rather than chosen, and why this is a deliberate trade.
            return Render(ListMarkerSubstitution.Apply(docx), fonts);
        }
        catch (DocumentConversionException) { throw; }
        catch (Exception ex) when (DocxPdfFailureDiagnosis.IsNegativeIndent(ex))
        {
            // A negative paragraph indent is refused at any magnitude and is legal in Word, so it
            // is clamped and retried - on the maintainer's decision, and see NegativeIndentClamp
            // for what that costs and buys.
            //
            // ON FAILURE ONLY, so the ordinary path never pays for it. The clamp has to open the
            // package to find out whether there is anything to clamp, and doing that on every
            // conversion would tax the 71 documents in 99 that never needed it. The retry costs a
            // second render, and only a document that already failed pays it - the same shape as
            // HtmlForPdf's repairs, for the same reason.
            var clamped = NegativeIndentClamp.Apply(ListMarkerSubstitution.Apply(docx));
            if (ReferenceEquals(clamped, docx)) throw new DocumentConversionException(
                DocxPdfFailureDiagnosis.Describe(ex) ?? FailureMessage, ex);

            try
            {
                return Render(clamped, fonts);
            }
            catch (Exception second)
            {
                throw new DocumentConversionException(
                    DocxPdfFailureDiagnosis.Describe(second) ?? FailureMessage, second);
            }
        }
        catch (Exception ex)
        {
            // A recognised cause gets named; everything else keeps the generic wrapper.
            // See DocxPdfFailureDiagnosis.
            throw new DocumentConversionException(DocxPdfFailureDiagnosis.Describe(ex) ?? FailureMessage, ex);
        }
    }

    /// <summary>Renders <paramref name="inputPath"/> to a PDF at <paramref name="outputPath"/>.</summary>
    public static void ConvertFile(string inputPath, string outputPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        File.WriteAllBytes(outputPath, Convert(File.ReadAllBytes(inputPath)));
    }

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and writes the rendered PDF to
    /// <paramref name="destination"/>.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed or closed, and <paramref name="destination"/> is never
    /// sought or read back, so both may be sockets, files or HTTP message bodies. Neither has to be
    /// seekable.
    ///
    /// The PDF is written straight through to <paramref name="destination"/> as the renderer
    /// produces it rather than being assembled into an array first, so a large document is
    /// delivered without ever existing in full in memory. The consequence of streaming is that a
    /// failure part-way through leaves whatever had already been produced on
    /// <paramref name="destination"/>.
    ///
    /// <b>No network access, and safe in an air-gapped environment</b>, as for
    /// <see cref="Convert(byte[])"/>.
    /// </summary>
    /// <param name="source">The stream the .docx package is read from.</param>
    /// <param name="destination">The stream the PDF is written to.</param>
    /// <param name="ct">Cancels the read, the render and the write.</param>
    /// <exception cref="ArgumentNullException">Either stream is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="destination"/>
    /// is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The DOCX could not be rendered.</exception>
    public static Task ConvertAsync(Stream source, Stream destination, CancellationToken ct = default)
        => ConvertAsync(source, destination, null, ct);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and writes the rendered PDF to
    /// <paramref name="destination"/>, using <paramref name="fonts"/> for characters the renderer
    /// cannot otherwise encode.
    /// </summary>
    /// <param name="source">The stream the .docx package is read from.</param>
    /// <param name="destination">The stream the PDF is written to.</param>
    /// <param name="fonts">Fonts to fall back to, or <see langword="null"/> for none.</param>
    /// <param name="ct">Cancels the read, the render and the write.</param>
    /// <exception cref="ArgumentNullException">Either stream is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="destination"/>
    /// is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be rendered.</exception>
    public static async Task ConvertAsync(
        Stream source, Stream destination, PdfFontOptions? fonts, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, "DOCX content was empty.", nameof(source), FailureMessage, ct)
            .ConfigureAwait(false);

        await RenderAsync(docx, destination, fonts, ct).ConfigureAwait(false);
    }

    /// <summary>Loads a package and renders it. The one place the renderer is actually called.</summary>
    private static byte[] Render(byte[] docx, PdfFontOptions? fonts)
    {
        // Copy into an expandable stream: OfficeIMO opens the package read/write.
        using var input = new MemoryStream();
        input.Write(docx, 0, docx.Length);
        input.Position = 0;

        using var word = WordDocument.Load(input);
        // NO ResourcePolicy, and that is measured rather than an oversight. See PdfRenderPolicy.
        var options = SaveOptions(fonts);
        return options is null ? word.ToPdf() : word.ToPdf(options);
    }

    /// <summary>
    /// The renderer's options, or <see langword="null"/> when the caller supplied no fonts.
    /// </summary>
    /// <remarks>
    /// <b>Null, not an empty options object, and the difference is measured.</b> Handing the Word
    /// renderer an options instance is not neutral - assigning a <c>ResourcePolicy</c> to one cost
    /// 14 of 99 real documents, for font reasons rather than resource ones (see
    /// <see cref="PdfRenderPolicy"/>). An empty instance was measured to behave like no instance,
    /// but "measured equivalent today" is exactly the assumption that failed last time, so the
    /// no-fonts path stays literally the call it always was.
    /// </remarks>
    private static WordPdfSaveOptions? SaveOptions(PdfFontOptions? fonts)
    {
        var fallbacks = fonts?.ToFallbackSet();
        return fallbacks is null
            ? null
            : new WordPdfSaveOptions { PdfOptions = new OfficeIMO.Pdf.PdfOptions { EmbeddedFontFallbacks = fallbacks } };
    }

    /// <summary>
    /// Renders an already-buffered .docx package onto <paramref name="destination"/>.
    ///
    /// Split out so <see cref="HtmlToPdfConverter"/> can hand over the package
    /// <see cref="HtmlToDocxConverter"/> has just built, rather than serialising it to an array and
    /// reading it back. <paramref name="docx"/> is a scratch buffer this library owns; it is read
    /// from its start and left to the caller to dispose.
    /// </summary>
    internal static async Task RenderAsync(
        MemoryStream docx, Stream destination, PdfFontOptions? fonts, CancellationToken ct)
    {
        docx.Position = 0;

        try
        {
            // OfficeIMO opens the package read/write, which is why this takes an expandable
            // MemoryStream rather than a read-only view over someone else's buffer.
            using var word = WordDocument.Load(docx);

            // Writes directly onto the caller's destination. OfficeIMO's writer emits the PDF in
            // pieces as it lays it out, so nothing here ever holds the whole rendered document.
            // NO ResourcePolicy - see PdfRenderPolicy for the measurement.
            var options = SaveOptions(fonts);
            if (options is null)
                await word.SaveAsPdfAsync(destination, cancellationToken: ct).ConfigureAwait(false);
            else
                await word.SaveAsPdfAsync(destination, options, ct).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            // A recognised cause gets named; everything else keeps the generic wrapper.
            // See DocxPdfFailureDiagnosis.
            throw new DocumentConversionException(DocxPdfFailureDiagnosis.Describe(ex) ?? FailureMessage, ex);
        }
    }

    private const string FailureMessage =
        "Failed to render DOCX to PDF. See the inner exception for details.";
}
