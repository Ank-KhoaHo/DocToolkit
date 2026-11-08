using System.Xml.Linq;

namespace DocToolkit;

/// <summary>Zeroes the left and right cell padding of every table, so a column has room for text.</summary>
/// <remarks>
/// <b>Runs only after 3.5 refuses with "cell padding must leave a positive text width", and only
/// when <see cref="EmptyCellMinWidth"/> did not already fix it.</b> 3.5 refuses a table whose
/// resolved column is no wider than its left and right padding together. Word draws such a table,
/// squeezing the text against the cell edge. Measured: a 50-column automatic table with the usual
/// 108-twip side margins is refused, and the same table with them at zero renders.
///
/// <b>It is the blunt step, which is why it comes second.</b> Every <c>w:left</c>/<c>w:right</c>
/// (and <c>w:start</c>/<c>w:end</c>) inside a <c>w:tblCellMar</c> or <c>w:tcMar</c> is set to zero,
/// in the document and in the table styles, so every table in the document loses its side padding,
/// not only the one that was refused. Top and bottom padding are not touched. A margin that is
/// absent is not added: the default padding of a table with no explicit margin is not rewritten.
/// </remarks>
internal static class CellPaddingZero
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>The repaired package, or the same array when every side margin is already zero.</summary>
    internal static byte[] Apply(byte[] docx) =>
        DocxPartRewrite.Apply(docx, name => name is "word/document.xml" or "word/styles.xml", ZeroSides);

    private static bool ZeroSides(XDocument doc)
    {
        var changed = false;
        foreach (var margins in doc.Descendants().Where(e => e.Name == W + "tblCellMar" || e.Name == W + "tcMar"))
        foreach (var side in margins.Elements().Where(s => s.Name.LocalName is "left" or "right" or "start" or "end"))
        {
            if ((string?)side.Attribute(W + "w") == "0" && (string?)side.Attribute(W + "type") == "dxa") continue;

            side.SetAttributeValue(W + "w", "0");
            side.SetAttributeValue(W + "type", "dxa");
            changed = true;
        }

        return changed;
    }
}
