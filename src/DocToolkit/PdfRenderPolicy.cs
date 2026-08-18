using OfficeIMO.Excel.Pdf;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;

namespace DocToolkit;

/// <summary>
/// The one place the PDF renderers' resource policy is decided.
///
/// <b>Both flags are set explicitly even though the upstream defaults already match.</b> That is not
/// belt-and-braces: it is the same reasoning <see cref="HtmlToDocxConverter"/> follows by always
/// handing <c>HtmlConverter</c> an explicit <c>IWebRequest</c> rather than relying on
/// <c>ImageProcessingMode</c>. A default is a policy the upstream author may revisit in a patch
/// release, and this package's offline guarantee must not be a property of somebody else's default.
///
/// Single site for the same reason as <c>SectionPropertiesFactory</c> and
/// <c>WorkbookEditor.SetCellValue</c>: two construction sites is how the XLSX and PPTX paths get to
/// disagree about what "offline" means.
/// </summary>
internal static class PdfRenderPolicy
{
    /// <summary>
    /// <c>AllowLocalFileAccess</c> is refused alongside the network flag because a document that can
    /// read the host's disk is the same class of disclosure the <c>file://</c> reach was closed for
    /// on the HTML path — it just arrives through a different door.
    /// </summary>
    private static PdfResourcePolicy Policy() => new()
    {
        AllowRemoteResourceResolution = false,
        AllowLocalFileAccess = false,
    };

    /// <summary>
    /// The Word path's options.
    /// </summary>
    /// <remarks>
    /// <b>Added 2026-08-18, and its absence is the more interesting half.</b> The XLSX and PPTX
    /// paths stated this policy from the start while <c>DocxToPdfConverter</c> called a bare
    /// <c>ToPdf()</c> - so the one path a Word document actually takes was the one inheriting the
    /// guarantee from a dependency default, which is precisely what the class comment above rejects.
    ///
    /// Nothing was leaking: <c>AirGapGuardTests</c> covers this path, including a DOCX carrying
    /// external references, and it passed throughout. But what stood between the guarantee and a
    /// dependency changing its mind was a behavioural test whose timing half is the one that flakes
    /// on macOS - a real guard, and a poor last line for something that can instead be said in a
    /// line of code.
    /// </remarks>
    public static WordPdfSaveOptions ForDocument() => new() { ResourcePolicy = Policy() };

    public static ExcelPdfSaveOptions ForWorkbook() => new() { ResourcePolicy = Policy() };

    public static PowerPointPdfSaveOptions ForPresentation() => new() { ResourcePolicy = Policy() };

    /// <summary>
    /// The two flags, for <c>XlsxPptxToPdfTests</c> to assert directly.
    ///
    /// A socket count of zero is also what a document referencing nothing produces, so the
    /// behavioural guard and this one say different things: this fails the moment somebody
    /// constructs the options without the flags, whatever the document happens to contain.
    /// </summary>
    internal static (bool AllowRemote, bool AllowLocal) DescribeForTests()
    {
        var policy = Policy();
        return (policy.AllowRemoteResourceResolution, policy.AllowLocalFileAccess);
    }
}
