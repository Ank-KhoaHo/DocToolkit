using OfficeIMO.Word.Markdown;

namespace DocToolkit;

/// <summary>
/// The one place Markdown → DOCX conversion is configured.
///
/// <b>Every option here that can touch the network or the disk is set explicitly, even where it
/// already carries the value we want.</b> That is the same rule <see cref="TextExportOptions"/> and
/// <c>PdfRenderPolicy</c> follow, and the argument is stronger here: these decide whether this
/// library performs I/O at all, and an inherited default is a decision the upstream author may
/// revisit in a patch release. Measured 2026-08-13 — the defaults happened to be right; that is
/// exactly the situation in which they are easiest to lose.
///
/// Single site so there is no second place the offline decision can be made, as with
/// <c>SectionPropertiesFactory</c> and <c>WorkbookEditor.SetCellValue</c>.
/// </summary>
internal static class MarkdownImportOptions
{
    public static MarkdownToWordOptions ForImport()
    {
        var options = NewOptions();

        // Get-only, so it is emptied in place rather than replaced - the same shape, and the same
        // reasoning, as DocToolkitOptions.RemoteImage: mutating cannot lose a restrictive default
        // that nobody deliberately changed, whereas assigning a fresh collection can.
        options.AllowedImageDirectories.Clear();

        return options;
    }

    private static MarkdownToWordOptions NewOptions() => new()
    {
        // No resolver means nothing can fetch a remote image. This IS the offline guarantee for
        // this converter - not a rendering-policy hint, but the absence of any mechanism to
        // perform I/O. Same reasoning as HtmlToDocxConverter always handing HtmlConverter an
        // explicit IWebRequest rather than trusting an ImageProcessingMode flag.
        RemoteImageResolver = null,

        // A remote image becomes a hyperlink rather than a fetch, so the content survives in a
        // form the reader can follow deliberately.
        FallbackRemoteImagesToHyperlinks = true,

        // Reading a path out of untrusted document content is a file-disclosure primitive: the
        // document decides which file gets embedded into the output. Refused outright; the
        // allow-list is emptied separately, in ForImport, because it is get-only.
        AllowLocalImages = false,

        // data: URIs carry their own bytes, so they cost no I/O. Bounded anyway - a document
        // should not be able to make us materialise an arbitrary amount of memory.
        AllowDataUriImages = true,
        MaxDataUriImageBytes = 32 * 1024 * 1024,

        // Nothing to resolve a relative reference against. A BaseUri would turn "![](a.png)" into
        // a fetchable absolute URL, which is the hole the two settings above exist to close.
        BaseUri = null,
    };
}
