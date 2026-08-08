using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace DocToolkit;

/// <summary>
/// Converts a Word (.docx) package to HTML, keeping the structure
/// <see cref="DocxEditor.ExtractText(byte[])"/> throws away: headings stay headings, tables stay
/// tables. Use <c>ExtractText</c> when flat text is what you want.
///
/// <b>The result is a complete HTML document</b> — <c>&lt;html&gt;&lt;head&gt;…&lt;body&gt;</c> —
/// not a fragment, and there is no option to change that. Embedding the output in a larger page
/// means extracting the body with an HTML parser you already trust; this package will not do it by
/// string surgery on the renderer's output.
///
/// <b>The result is self-contained.</b> Images in the source document are embedded as
/// <c>data:</c> URIs, so nothing in the output points at a file that does not exist.
/// </summary>
public static class DocxToHtmlConverter
{
    private const string FailureMessage = "Failed to convert DOCX to HTML.";

    /// <summary>Converts the .docx in <paramref name="docx"/> to a complete HTML document.</summary>
    /// <param name="docx">The document to convert.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The document could not be converted.</exception>
    public static string Convert(byte[] docx)
    {
        ArgumentNullException.ThrowIfNull(docx);
        if (docx.Length == 0)
            throw new ArgumentException("DOCX content was empty.", nameof(docx));

        try
        {
            // Expandable copy: the loader opens the package read/write, as DocxToPdfConverter does.
            using var input = new MemoryStream();
            input.Write(docx, 0, docx.Length);
            input.Position = 0;

            using var word = WordDocument.Load(input);
            return word.ToHtml(TextExportOptions.ForHtml());
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException(FailureMessage, ex);
        }
    }

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and returns it as a complete HTML document.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and is not disposed, closed or sought, so
    /// it may be forward-only — an HTTP request body, for instance.
    /// </summary>
    /// <param name="source">The stream the .docx package is read from.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be converted.</exception>
    public static async Task<string> ConvertAsync(Stream source, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, "DOCX content was empty.", nameof(source), FailureMessage, ct)
            .ConfigureAwait(false);

        return Convert(docx.ToArray());
    }
}
