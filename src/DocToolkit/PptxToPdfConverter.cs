using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;

namespace DocToolkit;

/// <summary>
/// Renders a PowerPoint (.pptx) presentation to PDF, one page per slide. Pure managed — no browser,
/// no LibreOffice, no PowerPoint.
///
/// The same three members as <see cref="DocxToPdfConverter"/> and <see cref="XlsxToPdfConverter"/>,
/// deliberately: a consumer who has used one has used all three.
///
/// <b>No page-setup parameter.</b> A presentation's page size is its <c>p:sldSz</c> — slide
/// geometry, not paper. <see cref="PresentationEditor"/> writes 16:9, which renders as a 960 × 540
/// pt PDF page. Accepting a <see cref="PageSetup"/> here would invite a caller to letterbox their
/// own slides onto a paper size that means nothing to a deck.
///
/// <b>Fidelity is bounded</b>, and features the renderer cannot represent are dropped rather than
/// reported — see README's <i>Known limitations</i>.
/// </summary>
public static class PptxToPdfConverter
{
    private const string FailureMessage =
        "Failed to convert PPTX to PDF. See the inner exception for details.";

    /// <summary>
    /// Renders the .pptx in <paramref name="pptx"/> and returns PDF bytes, one page per slide.
    ///
    /// <b>No network access, and safe in an air-gapped environment.</b> The resource policy handed
    /// to the renderer refuses both remote resolution and local file access, set explicitly rather
    /// than inherited — see <see cref="PdfRenderPolicy"/>.
    /// </summary>
    /// <param name="pptx">The presentation to render.</param>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The presentation could not be rendered.</exception>
    /// <example>
    /// <code source="../../tests/DocToolkit.Tests/DocumentationExamples.cs" region="PptxToPdf"/>
    /// </example>
    public static byte[] Convert(byte[] pptx)
    {
        ArgumentNullException.ThrowIfNull(pptx);
        if (pptx.Length == 0)
            throw new ArgumentException("PPTX content was empty.", nameof(pptx));

        try
        {
            // Expandable copy: the loader opens the package read/write, exactly as
            // DocxToPdfConverter.Convert has to.
            using var input = new MemoryStream();
            input.Write(pptx, 0, pptx.Length);
            input.Position = 0;

            using var presentation = PowerPointPresentation.Load(input);
            return presentation.ToPdf(PdfRenderPolicy.ForPresentation());
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException(FailureMessage, ex);
        }
    }

    /// <summary>Renders <paramref name="inputPath"/> to a PDF at <paramref name="outputPath"/>.</summary>
    /// <param name="inputPath">The .pptx to read.</param>
    /// <param name="outputPath">Where to write the PDF. Overwritten if it exists.</param>
    /// <exception cref="ArgumentException">Either path is blank.</exception>
    /// <exception cref="DocumentConversionException">The presentation could not be rendered.</exception>
    public static void ConvertFile(string inputPath, string outputPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        File.WriteAllBytes(outputPath, Convert(FilePipeline.Read(inputPath, nameof(inputPath))));
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/> and writes the rendered PDF to
    /// <paramref name="destination"/>, one page per slide.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed or closed, and <paramref name="destination"/> is never
    /// sought or read back, so both may be sockets, files or HTTP message bodies.
    ///
    /// As with <see cref="DocxToPdfConverter.ConvertAsync(System.IO.Stream, System.IO.Stream, System.Threading.CancellationToken)"/>, the PDF is written straight through as
    /// the renderer produces it rather than assembled into an array first — so a failure part-way
    /// leaves whatever had already been produced on <paramref name="destination"/>.
    /// </summary>
    /// <param name="source">The stream the .pptx package is read from.</param>
    /// <param name="destination">The stream the PDF is written to.</param>
    /// <param name="ct">Cancels the read, the render and the write.</param>
    /// <exception cref="ArgumentNullException">Either stream is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="destination"/>
    /// is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The presentation could not be rendered.</exception>
    public static async Task ConvertAsync(
        Stream source, Stream destination, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var pptx = await StreamPipeline
            .DrainAsync(source, "PPTX content was empty.", nameof(source), FailureMessage, ct)
            .ConfigureAwait(false);

        try
        {
            using var presentation = PowerPointPresentation.Load(pptx);
            await presentation
                .SaveAsPdfAsync(destination, PdfRenderPolicy.ForPresentation(), ct)
                .ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException(FailureMessage, ex);
        }
    }
}
