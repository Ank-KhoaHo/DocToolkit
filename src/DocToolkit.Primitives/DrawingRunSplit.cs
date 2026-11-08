using System.Xml.Linq;

namespace DocToolkit;

/// <summary>
/// Gives every <c>w:drawing</c> in a run its own run, so the Markdown export keeps every picture.
/// </summary>
/// <remarks>
/// <b>OfficeIMO 3.5's DOCX to Markdown export emits only the first <c>w:drawing</c> of a run.</b>
/// A run holding two pictures lost the second (measured on a real document, doc-000130; 3.2.6 kept
/// both). Word writes such a run when pictures are pasted side by side inside one run.
///
/// <b>Markdown only.</b> The PDF path renders both drawings of a run, so it is never split there
/// and its output cannot change.
///
/// <b>Each new run copies the original's attributes and <c>rPr</c>.</b> Children other than the
/// drawings (text, breaks) stay with the drawing run that precedes them, so their order is kept; a
/// leading non-drawing child stays in the first run.
///
/// <b>Cost:</b> one parse of <c>word/document.xml</c>. A document where every run holds at most one
/// drawing is returned by reference. Only the main part is read: that is where the export reads
/// pictures from.
/// </remarks>
internal static class DrawingRunSplit
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>The repaired package, or the same array when no run holds two drawings.</summary>
    internal static byte[] Apply(byte[] docx) =>
        DocxPartRewrite.Apply(docx, name => name == "word/document.xml", Split);

    private static bool Split(XDocument doc)
    {
        var changed = false;
        foreach (var run in doc.Descendants(W + "r").Where(r => r.Elements(W + "drawing").Count() > 1).ToList())
        {
            var rPr = run.Element(W + "rPr");
            var pieces = new List<XElement>();
            XElement? current = null;
            foreach (var child in run.Elements().Where(e => e.Name != W + "rPr"))
            {
                if (current is null || child.Name == W + "drawing")
                {
                    current = new XElement(W + "r", run.Attributes(), rPr is null ? null : new XElement(rPr));
                    pieces.Add(current);
                }

                current.Add(new XElement(child));
            }

            run.ReplaceWith(pieces);
            changed = true;
        }

        return changed;
    }
}
