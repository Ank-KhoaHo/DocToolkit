using System.Xml.Linq;

namespace DocToolkit;

/// <summary>
/// Moves runs with no visible text out of a hyperlink, splitting the link around them, so the
/// document can be rendered.
/// </summary>
/// <remarks>
/// <b>OfficeIMO 3.5 turns each hyperlink into link segments, and refuses a segment with no text.</b>
/// A run holding only a break, a tab or spaces inside a hyperlink becomes such a segment. Word
/// writes this when a line break sits inside a link. 3.2.6 rendered it.
///
/// <b>The break is kept, in the same place.</b> The blank run is moved to sit between two copies of
/// the link (same attributes, the original runs split across them), so the line still breaks where
/// it did; only the link no longer spans the break. If every run in a link is blank the link is
/// unwrapped, since it had no text to click on. A drawing-only run has no <c>w:t</c> either, so it
/// counts as blank and a link around only a picture is unwrapped too - that is what the corpus
/// measurement used, and it has not been separated out.
///
/// <b>It cannot change a document that renders today.</b> A link with a blank run is the refused
/// shape, so a document without one is returned by reference.
/// </remarks>
internal static class HyperlinkBlankRunSplit
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>The repaired package, or the same array when no hyperlink holds a blank run.</summary>
    internal static byte[] Apply(byte[] docx) =>
        DocxPartRewrite.Apply(docx, name => name == "word/document.xml", Split);

    private static bool Split(XDocument doc)
    {
        var changed = false;
        foreach (var link in doc.Descendants(W + "hyperlink").ToList())
        {
            var children = link.Elements().ToList();
            if (!children.Any(IsBlankRun)) continue;

            if (children.All(IsBlankRun))
            {
                link.ReplaceWith(children.Select(r => new XElement(r)));
                changed = true;
                continue;
            }

            var pieces = new List<XElement>();
            XElement? current = null;
            foreach (var child in children)
            {
                if (IsBlankRun(child))
                {
                    current = null;
                    pieces.Add(new XElement(child));
                    continue;
                }

                if (current is null)
                {
                    current = new XElement(link.Name, link.Attributes());
                    pieces.Add(current);
                }

                current.Add(new XElement(child));
            }

            link.ReplaceWith(pieces);
            changed = true;
        }

        return changed;
    }

    private static bool IsBlankRun(XElement e) =>
        e.Name == W + "r" && string.IsNullOrWhiteSpace(string.Concat(e.Elements(W + "t").Select(t => t.Value)));
}
