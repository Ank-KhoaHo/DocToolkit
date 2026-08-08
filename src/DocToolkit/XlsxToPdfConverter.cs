using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;

namespace DocToolkit;

/// <summary>
/// Renders an Excel (.xlsx) workbook to PDF. Pure managed — no browser, no LibreOffice, no Excel.
///
/// The same three members as <see cref="DocxToPdfConverter"/> and <see cref="PptxToPdfConverter"/>,
/// deliberately: a consumer who has used one has used all three.
///
/// <b>No page-setup parameter</b>, for the same reason <see cref="DocxToPdfConverter"/> has none —
/// the workbook already carries its own print setup, per worksheet, and the renderer honours it.
/// Taking a <see cref="PageSetup"/> here would mean overriding something the caller's own document
/// already states.
///
/// <b>Fidelity is bounded</b>, and features the renderer cannot represent are dropped rather than
/// reported: this package has no warning channel, and adding one is a decision deferred until
/// somebody needs it rather than made by accident here. See README's <i>Known limitations</i>.
/// </summary>
public static class XlsxToPdfConverter
{
    private const string FailureMessage = "Failed to convert XLSX to PDF.";

    /// <summary>
    /// Renders the .xlsx in <paramref name="xlsx"/> and returns PDF bytes.
    ///
    /// <b>No network access, and safe in an air-gapped environment.</b> The resource policy handed
    /// to the renderer refuses both remote resolution and local file access, set explicitly rather
    /// than inherited — see <see cref="PdfRenderPolicy"/>.
    /// </summary>
    /// <param name="xlsx">The workbook to render.</param>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be rendered.</exception>
    public static byte[] Convert(byte[] xlsx)
    {
        ArgumentNullException.ThrowIfNull(xlsx);
        if (xlsx.Length == 0)
            throw new ArgumentException("XLSX content was empty.", nameof(xlsx));

        try
        {
            // Expandable copy: the loader opens the package read/write, exactly as
            // DocxToPdfConverter.Convert has to.
            using var input = new MemoryStream();
            input.Write(xlsx, 0, xlsx.Length);
            input.Position = 0;

            using var workbook = ExcelDocument.Load(input);
            return workbook.ToPdf(PdfRenderPolicy.ForWorkbook());
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException(FailureMessage, ex);
        }
    }

    /// <summary>Renders <paramref name="inputPath"/> to a PDF at <paramref name="outputPath"/>.</summary>
    /// <param name="inputPath">The .xlsx to read.</param>
    /// <param name="outputPath">Where to write the PDF. Overwritten if it exists.</param>
    /// <exception cref="ArgumentException">Either path is blank.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be rendered.</exception>
    public static void ConvertFile(string inputPath, string outputPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        File.WriteAllBytes(outputPath, Convert(File.ReadAllBytes(inputPath)));
    }

    /// <summary>
    /// Reads an .xlsx from <paramref name="source"/> and writes the rendered PDF to
    /// <paramref name="destination"/>.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed or closed, and <paramref name="destination"/> is never
    /// sought or read back, so both may be sockets, files or HTTP message bodies.
    ///
    /// As with <see cref="DocxToPdfConverter.ConvertAsync"/>, the PDF is written straight through as
    /// the renderer produces it rather than assembled into an array first — so a failure part-way
    /// leaves whatever had already been produced on <paramref name="destination"/>.
    /// </summary>
    /// <param name="source">The stream the .xlsx package is read from.</param>
    /// <param name="destination">The stream the PDF is written to.</param>
    /// <param name="ct">Cancels the read, the render and the write.</param>
    /// <exception cref="ArgumentNullException">Either stream is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="destination"/>
    /// is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The workbook could not be rendered.</exception>
    public static async Task ConvertAsync(
        Stream source, Stream destination, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var xlsx = await StreamPipeline
            .DrainAsync(source, "XLSX content was empty.", nameof(source), FailureMessage, ct)
            .ConfigureAwait(false);

        try
        {
            using var workbook = ExcelDocument.Load(xlsx);
            await workbook
                .SaveAsPdfAsync(destination, PdfRenderPolicy.ForWorkbook(), ct)
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
