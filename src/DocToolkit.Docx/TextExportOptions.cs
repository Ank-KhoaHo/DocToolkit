using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;

namespace DocToolkit;

/// <summary>
/// The one place the DOCX → text-format exporters are configured.
///
/// <b><see cref="WordToHtmlOptions.EmbedImagesAsBase64"/> is set explicitly even though it already
/// defaults to true.</b> That default is what makes the output self-contained — an image embedded in
/// the source document comes back as a <c>data:</c> URI rather than a reference to a file that does
/// not exist anywhere. Inheriting it would make that property a decision the upstream author could
/// revisit in a patch release, silently turning self-contained output into output with dangling
/// references. Same reasoning as <c>PdfRenderPolicy</c>.
///
/// Single site so the HTML and Markdown paths cannot drift, as with
/// <c>SectionPropertiesFactory</c> and <c>WorkbookEditor.SetCellValue</c>.
/// </summary>
internal static class TextExportOptions
{
    public static WordToHtmlOptions ForHtml() => new() { EmbedImagesAsBase64 = true };

    // Markdown expresses the same choice differently: an enum with a File alternative that would
    // write images beside the output and reference them by path. Base64 is the self-contained one.
    public static WordToMarkdownOptions ForMarkdown() =>
        new() { ImageExportMode = ImageExportMode.Base64 };

}
