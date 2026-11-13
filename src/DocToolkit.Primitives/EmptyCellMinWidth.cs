using System.Globalization;
using System.Xml.Linq;

namespace DocToolkit;

/// <summary>Gives a table cell with nothing in it a minimum width, so its column cannot size to zero.</summary>
/// <remarks>
/// <b>Runs only after 3.5 refuses with "cell padding must leave a positive text width".</b> An
/// automatic column whose cells are all empty and unsized resolves to a width of zero, and then no
/// padding - not even none - leaves a positive text width. Measured on html-000062: a 9-column table
/// whose 9th column was empty in every row was refused at zero padding too, and 144 twips (a tenth
/// of an inch) fixed it. HTML-to-DOCX writers emit exactly such tables for a trailing spacer cell.
///
/// <b>It gives the width and leaves the padding alone</b>, which is why it runs before
/// <see cref="CellPaddingZero"/>: when this is enough, the document keeps its cell margins.
///
/// <b>Empty means no text and no drawing.</b> A cell whose <c>w:tcW</c> is a positive width in
/// twips, or a percentage, is already sized and is left as it was. A missing <c>w:tcW</c>, a zero
/// width, or any other type (<c>auto</c>, <c>nil</c>) is set to 144 twips.
/// </remarks>
internal static class EmptyCellMinWidth
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string MinimumWidth = "144";

    /// <summary>The repaired package, or the same array when no empty cell needed a width.</summary>
    internal static byte[] Apply(byte[] docx) =>
        DocxPartRewrite.Apply(docx, name => name == "word/document.xml", GiveWidth);

    private static bool GiveWidth(XDocument doc)
    {
        var changed = false;
        foreach (var cell in doc.Descendants(W + "tc")
                     .Where(cell => string.IsNullOrWhiteSpace(string.Concat(cell.Descendants(W + "t").Select(t => t.Value)))
                                    && !cell.Descendants(W + "drawing").Any()))
        {
            var properties = cell.Element(W + "tcPr");
            var width = properties?.Element(W + "tcW");
            var type = (string?)width?.Attribute(W + "type");

            if (type == "pct") continue;
            if (type == "dxa"
                && int.TryParse((string?)width!.Attribute(W + "w"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var twips)
                && twips > 0) continue;

            if (properties is null)
            {
                properties = new XElement(W + "tcPr");
                cell.AddFirst(properties);
            }

            if (width is null)
            {
                width = new XElement(W + "tcW");
                // cnfStyle is the only element the schema puts before tcW.
                if (properties.Element(W + "cnfStyle") is { } conditional) conditional.AddAfterSelf(width);
                else properties.AddFirst(width);
            }

            width.SetAttributeValue(W + "w", MinimumWidth);
            width.SetAttributeValue(W + "type", "dxa");
            changed = true;
        }

        return changed;
    }
}
