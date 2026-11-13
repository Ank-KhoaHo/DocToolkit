using System.Globalization;
using System.Xml.Linq;

namespace DocToolkit;

/// <summary>Pulls in a paragraph indent that is wider than the narrowest column it could sit in.</summary>
/// <remarks>
/// <b>Runs only after 3.5 refuses with "indents must leave a positive text width".</b> Word accepts a
/// paragraph indented further than its column is wide (it overflows or wraps one word per line); 3.5
/// refuses it. Measured on doc-000111: a section with two explicit columns of 4,752 twips and a
/// paragraph with <c>left=6480 right=-576 hanging=6480</c>, refused from the multi-column renderer.
///
/// <b>The column is the narrowest one in the document.</b> Which paragraph sits in which column is
/// not known here, so the limit is the narrowest of every section's explicit <c>w:col</c> widths, or
/// else the text width less the gaps, divided by the column count (the gap defaults to 720). A left
/// (or <c>start</c>) indent above <c>column - 720 - max(right, 0)</c> is set to that value, floored
/// at zero, and a hanging indent is capped at it. Half an inch is left for text, which is what makes
/// the line measurable again. Both spellings are read, since newer Word and Strict write
/// <c>start</c> and <c>end</c>. The document, the styles and the numbering levels are all rewritten,
/// as an indent can come from any of them.
///
/// <b>It changes the layout of a paragraph that was always going to be wider than its column</b>, and
/// only in a document 3.5 refused, so nothing that renders today is touched. A column of 720 twips or
/// less (or none found) is left alone: there is no sensible indent to fit it to.
/// </remarks>
internal static class IndentFitToColumn
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const int TextRoom = 720;

    /// <summary>The repaired package, or the same array when no indent is too wide.</summary>
    internal static byte[] Apply(byte[] docx)
    {
        // The limit comes from the main part, so it is read first and handed to the other parts.
        var column = NarrowestColumn(docx);
        if (column <= TextRoom) return docx;

        return DocxPartRewrite.Apply(
            docx,
            name => name is "word/document.xml" or "word/styles.xml" or "word/numbering.xml",
            doc => Fit(doc, column));
    }

    private static int Number(XElement? element, string attribute, int fallback = 0) =>
        int.TryParse((string?)element?.Attribute(W + attribute), NumberStyles.Integer, CultureInfo.InvariantCulture, out var value)
            ? value
            : fallback;

    private static int NarrowestColumn(byte[] docx)
    {
        var narrowest = int.MaxValue;

        // An edit that never reports a change still receives each parsed part, which reads the main
        // part with the same safe parsing and without copying the package.
        DocxPartRewrite.Apply(docx, name => name == "word/document.xml", doc =>
        {
            foreach (var section in doc.Descendants(W + "sectPr"))
            {
                // long, because the values are untrusted: Math.Abs(int.MinValue) throws, and the
                // sums below can overflow an int.
                var margins = section.Element(W + "pgMar");
                var text = (long)Number(section.Element(W + "pgSz"), "w")
                    - Math.Abs((long)Number(margins, "left")) - Math.Abs((long)Number(margins, "right"));
                if (text <= 0) continue;

                var columns = section.Element(W + "cols");
                var explicitWidths = columns?.Elements(W + "col").Select(c => Number(c, "w")).Where(w => w > 0).ToList();
                var count = Math.Max(1, Number(columns, "num", 1));
                var width = explicitWidths is { Count: > 0 }
                    ? explicitWidths.Min()
                    : (text - (count - 1L) * Number(columns, "space", 720)) / count;

                narrowest = (int)Math.Clamp(Math.Min(narrowest, width), int.MinValue, int.MaxValue);
            }

            return false;
        });

        return narrowest == int.MaxValue ? 0 : narrowest;
    }

    private static bool Fit(XDocument doc, int column)
    {
        var changed = false;
        foreach (var indent in doc.Descendants(W + "ind"))
        {
            var right = Math.Max(Number(indent, "right"), Number(indent, "end"));
            var room = (int)Math.Max(0L, column - (long)TextRoom - Math.Max(right, 0));

            var fitted = false;
            foreach (var name in new[] { "left", "start" }
                         .Where(name => indent.Attribute(W + name) is not null && Number(indent, name) > room))
            {
                indent.SetAttributeValue(W + name, room.ToString(CultureInfo.InvariantCulture));
                fitted = true;
            }

            if (!fitted) continue;

            if (Number(indent, "hanging") > room)
                indent.SetAttributeValue(W + "hanging", room.ToString(CultureInfo.InvariantCulture));
            changed = true;
        }

        return changed;
    }
}
