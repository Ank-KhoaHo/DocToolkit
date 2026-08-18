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
    public static byte[] Convert(byte[] docx)
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
            var renderable = ListMarkerSubstitution.Apply(docx);

            // Copy into an expandable stream: OfficeIMO opens the package read/write.
            using var input = new MemoryStream();
            input.Write(renderable, 0, renderable.Length);
            input.Position = 0;

            using var word = WordDocument.Load(input);
            return word.ToPdf(PdfRenderPolicy.ForDocument());
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
    public static async Task ConvertAsync(Stream source, Stream destination, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, "DOCX content was empty.", nameof(source), FailureMessage, ct)
            .ConfigureAwait(false);

        await RenderAsync(docx, destination, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Renders an already-buffered .docx package onto <paramref name="destination"/>.
    ///
    /// Split out so <see cref="HtmlToPdfConverter"/> can hand over the package
    /// <see cref="HtmlToDocxConverter"/> has just built, rather than serialising it to an array and
    /// reading it back. <paramref name="docx"/> is a scratch buffer this library owns; it is read
    /// from its start and left to the caller to dispose.
    /// </summary>
    internal static async Task RenderAsync(MemoryStream docx, Stream destination, CancellationToken ct)
    {
        docx.Position = 0;

        try
        {
            // OfficeIMO opens the package read/write, which is why this takes an expandable
            // MemoryStream rather than a read-only view over someone else's buffer.
            using var word = WordDocument.Load(docx);

            // Writes directly onto the caller's destination. OfficeIMO's writer emits the PDF in
            // pieces as it lays it out, so nothing here ever holds the whole rendered document.
            await word.SaveAsPdfAsync(destination, PdfRenderPolicy.ForDocument(), ct).ConfigureAwait(false);
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
