using System.Globalization;
using System.Xml.Linq;

namespace DocToolkit;

/// <summary>
/// Scales a fixed-width table that is wider than the text area down to the text area, so OfficeIMO
/// 3.5 does not draw it off the page.
/// </summary>
/// <remarks>
/// <b>This is not a refusal: 3.5 renders the table, and runs it off the right edge.</b> OfficeIMO 3.2.6
/// squeezed a table wider than the page into it; 3.5 honours the width. Measured on raw 3.5.3 on an
/// A4 page with 1440 twip side margins (text area 9026 twips, 595 pt page): a two-column table of
/// 6000 + 6000 twips puts the right-aligned text of its second column at x = 624 pt, outside the page.
///
/// <b>Measured, and the conditions are what keep this from touching a table that fits today.</b>
/// Two shapes are drawn off the page, and only they are scaled:
/// <list type="bullet">
/// <item>a <c>dxa</c> table width larger than the text area (autofit or fixed layout, whatever the grid);</item>
/// <item>with no <c>dxa</c> width (none, or automatic): a fixed layout whose grid and whose cell
/// widths are both wider than the text area.</item>
/// </list>
/// These render inside the page as they are, so they are left alone: a percentage width (also with a
/// fixed layout), an autofit table with no width whose grid is wide, a <c>dxa</c> table width that
/// fits over a wider grid, and a fixed layout whose cell widths are narrower than its grid.
///
/// <b>Each table is measured against its own section.</b> A section's page size and margins sit in
/// the <c>w:pPr/w:sectPr</c> of its last paragraph; only the final section's <c>w:sectPr</c> is a direct
/// child of <c>w:body</c>. A table is measured against the next paragraph-level section after it, or
/// the body's own when none follows. Widths are read as <c>long</c>, so a grid whose columns sum past
/// <c>int.MaxValue</c> scales instead of throwing.
///
/// <b>Only top-level body tables are scaled.</b> A table nested in a cell is sized by that cell, and
/// scaling it as well would shrink it twice. A wide nested table is therefore not repaired.
///
/// <b>What it costs.</b> Every column of a repaired table is narrower by the same factor, so long
/// words wrap sooner and a very wide table gets small columns. The factor truncates, so the scaled
/// widths never add up to more than the text area. Cell margins, indents and percentage cell widths
/// are not scaled.
/// </remarks>
internal static class WideTableScale
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>The package with wide tables scaled, or the same array when none is wider than the page.</summary>
    internal static byte[] Apply(byte[] docx) =>
        DocxPartRewrite.Apply(docx, name => name == "word/document.xml", ScaleTables);

    private static bool ScaleTables(XDocument doc)
    {
        var body = doc.Root?.Element(W + "body");
        if (body is null) return false;

        // Each top-level table is measured against ITS OWN section. In Word a section's w:sectPr
        // sits in the pPr of that section's last paragraph; only the final section's is a direct
        // child of w:body. So the section of a table is the next paragraph-level w:sectPr after it,
        // or the body's own when none follows. Walking the body once, backwards, keeps that a single
        // pass: `text` is the text width of the section that the elements still to be visited close.
        var bodySection = body.Elements(W + "sectPr").LastOrDefault();
        long? text = bodySection is null ? null : TextWidth(bodySection);

        var changed = false;
        foreach (var element in body.Elements().Reverse())
        {
            if (element.Name == W + "tbl")
            {
                if (text is > 0) changed |= Scale(element, text.Value);
            }
            else if (element.Name == W + "p" && element.Element(W + "pPr")?.Element(W + "sectPr") is { } ending)
            {
                // This paragraph ends a section: everything before it (back to the previous such
                // paragraph) belongs to that section.
                text = TextWidth(ending);
            }
        }

        return changed;
    }

    /// <summary>The page width less both side margins, or 0 when the section has no page size.</summary>
    private static long TextWidth(XElement section)
    {
        var margins = section.Element(W + "pgMar");
        return Twips(section.Element(W + "pgSz"), "w") - Twips(margins, "left") - Twips(margins, "right");
    }

    private static bool Scale(XElement table, long text)
    {
        var properties = table.Element(W + "tblPr");
        var tableWidth = properties?.Element(W + "tblW");
        var type = (string?)tableWidth?.Attribute(W + "type");
        if (type == "pct") return false;

        var grid = table.Element(W + "tblGrid")?.Elements(W + "gridCol").ToList() ?? [];
        var gridWidth = grid.Sum(g => Twips(g, "w"));  // long: Twips is long, so the sum cannot overflow

        long width;
        if (type == "dxa")
        {
            width = Twips(tableWidth, "w");
        }
        else
        {
            var fixedLayout = (string?)properties?.Element(W + "tblLayout")?.Attribute(W + "type") == "fixed";
            width = fixedLayout && gridWidth > text ? Math.Min(gridWidth, WidestRow(table, gridWidth)) : 0;
        }

        if (width <= text) return false;

        var factor = (double)text / width;  // 0 < factor < 1: width > text > 0
        var toScale = grid
            .Concat(table.Elements(W + "tr").Elements(W + "tc").Elements(W + "tcPr").Elements(W + "tcW")
                .Where(e => (string?)e.Attribute(W + "type") == "dxa"));
        if (type == "dxa" && tableWidth is not null) toScale = toScale.Append(tableWidth);

        foreach (var element in toScale)
            element.SetAttributeValue(W + "w", ((long)(Twips(element, "w") * factor)).ToString(CultureInfo.InvariantCulture));
        return true;
    }

    /// <summary>
    /// The widest row, by its cell widths. A row in which any cell has no <c>dxa</c> width is taken
    /// as <paramref name="gridWidth"/>: the renderer then sizes it from the grid.
    /// </summary>
    private static long WidestRow(XElement table, long gridWidth)
    {
        long widest = 0;
        foreach (var row in table.Elements(W + "tr"))
        {
            var cells = row.Elements(W + "tc").ToList();
            var widths = cells
                .Select(c => c.Element(W + "tcPr")?.Element(W + "tcW"))
                .Select(w => (string?)w?.Attribute(W + "type") == "dxa" ? Twips(w, "w") : -1)
                .ToList();
            widest = Math.Max(widest, widths.Contains(-1) ? gridWidth : widths.Sum());
        }

        return widest == 0 ? gridWidth : widest;
    }

    /// <summary>
    /// A twip attribute as a <see langword="long"/>, clamped to the range of an <see langword="int"/>
    /// so that summing many of them cannot overflow either. A value that is not a number is 0.
    /// </summary>
    private static long Twips(XElement? element, string attribute) =>
        long.TryParse((string?)element?.Attribute(W + attribute), NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? Math.Clamp(value, int.MinValue, int.MaxValue)
            : 0;
}
