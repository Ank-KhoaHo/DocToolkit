using System.Xml.Linq;

namespace DocToolkit;

/// <summary>
/// Turns a percentage table or cell width under 2% into automatic so the document can be rendered.
/// </summary>
/// <remarks>
/// <b>OOXML counts <c>pct</c> in fiftieths of a percent, so <c>w:w="50" w:type="pct"</c> is a
/// genuine 1% width</b> - and HTML-to-DOCX writers emit exactly that for <c>width="1%"</c>
/// (html-000079: <c>&lt;td width="1%"&gt;</c> cells and a <c>width="1%"</c> table). Word and
/// OfficeIMO 3.2.6 rendered a sliver. OfficeIMO 3.5 resolves it, and a percentage of it again when
/// tables nest, to a width of about zero and refuses the document.
///
/// <b>The refusal is ArgumentException "Table preferred width must be a positive finite value".</b>
/// It was bisected from html-000079 to a 1% table with cell spacing nested in a cell; neither the
/// 1% width nor the spacing alone is refused, and a top-level table is not. Other shapes with a tiny
/// percentage may fail the same way, which is why the repair keys on the width and not on the
/// nesting.
///
/// <b>Below 100 fiftieths (2%) is read as "the writer meant automatic".</b> No real layout asks for
/// a table narrower than 2% of the page, so the cutoff separates the misunderstood unit from
/// deliberate widths. A width of 2% or more is left exactly as it was, even if it is also wrong.
/// Only <c>w:tblW</c> and <c>w:tcW</c> are touched, and only integer values - the
/// <c>"50%"</c> spelling is a different, valid form. Auto sizes the table to its content, which is
/// what the writer asked for by leaving the cell otherwise unsized.
///
/// <b>It does change some documents that render today.</b> A top-level 1% table is not refused, and
/// is edited all the same: it becomes a table sized to its content instead of a 1% sliver. That is
/// what the writer meant, and it is why the width is repaired rather than only the refused nesting.
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
