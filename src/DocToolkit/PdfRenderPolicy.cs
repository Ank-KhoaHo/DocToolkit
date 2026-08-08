using OfficeIMO.Excel.Pdf;
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
