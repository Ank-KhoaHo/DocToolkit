using System.Xml.Linq;

namespace DocToolkit;

/// <summary>
/// Turns a percentage table or cell width under 2% into automatic so the document can be rendered.
/// </summary>
/// <remarks>
/// <b>HTML-to-DOCX writers turn <c>width="50%"</c> into <c>w:w="50" w:type="pct"</c>, and OOXML
/// counts <c>pct</c> in fiftieths of a percent</b> - so the width is 1%. OfficeIMO 3.5 resolves
/// that to no usable width and refuses the table; Word shows a sliver.
///
/// <b>Below 100 fiftieths (2%) is read as "the writer meant automatic".</b> No real layout asks for
/// a table narrower than 2% of the page, so the cutoff separates the misunderstood unit from
/// deliberate widths. A width of 2% or more is left exactly as it was, even if it is also wrong.
/// Only <c>w:tblW</c> and <c>w:tcW</c> are touched, and only integer values - the
/// <c>"50%"</c> spelling is a different, valid form.
///
/// <b>It cannot change a document that renders today</b>: the refused shape is the only one edited.
/// </remarks>
internal static class TinyPercentWidthRepair
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>The repaired package, or the same array when no width is a tiny percentage.</summary>
    internal static byte[] Apply(byte[] docx) =>
        DocxPartRewrite.Apply(docx, name => name == "word/document.xml", MarkAuto);

    private static bool MarkAuto(XDocument doc)
    {
        var changed = false;
        foreach (var width in doc.Descendants().Where(e => e.Name == W + "tblW" || e.Name == W + "tcW"))
        {
            if ((string?)width.Attribute(W + "type") != "pct"
                || !int.TryParse((string?)width.Attribute(W + "w"), out var value)
                || value >= 100) continue;

            width.SetAttributeValue(W + "type", "auto");
            width.SetAttributeValue(W + "w", "0");
            changed = true;
        }

        return changed;
    }
}
