using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace DocToolkit;

/// <summary>
/// Moves runs with no visible text out of a hyperlink that sits in a table cell, splitting the link
/// around them, so the document can be rendered.
/// </summary>
/// <remarks>
/// <b>OfficeIMO 3.5 refuses a hyperlink in a table cell that holds a run with no text</b>
/// (ArgumentException "Parameter 'text' cannot be empty or whitespace."). A break, a tab or spaces
/// inside the link is the usual shape - Word writes it when a line break sits inside a link. 3.2.6
/// rendered it. Measured on 3.5.3, the same link in a body paragraph renders, and a cell link is
/// refused in the body, headers, footers and footnotes alike (endnotes share the code path and are
/// covered on that basis, not measured).
///
/// <b>Only links with a <c>w:tc</c> ancestor are touched,</b> so a link that renders today is never
/// rewritten.
///
/// <b>A run is blank only when nothing in it can be seen or means something.</b> Its content is a
/// break, a tab, a carriage return, a hyphen or whitespace-only text - nothing else. A run with a
/// picture, a footnote reference, a symbol or a field character is NOT blank: a link around only a
/// picture renders on 3.5.3 (measured), and unwrapping it would silently lose the link. A link that
/// contains a field is left alone altogether, because splitting it between the field's begin and
/// end would break the field; that link stays refused rather than corrupted.
///
/// <b>The break is kept, in the same place.</b> The blank run sits between two copies of the link
/// carrying the same <c>r:id</c> / <c>w:anchor</c>, so the line still breaks where it did; only the
/// link no longer spans the break. If every child is blank the link is unwrapped, since it had no
/// text to click on.
///
/// <b>It cannot change a document that renders today</b>: a cell link with a blank run is the
/// refused shape, so a document without one is returned by reference.
/// </remarks>
internal static partial class HyperlinkBlankRunSplit
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static readonly HashSet<XName> InvisibleContent =
    [
        W + "rPr", W + "br", W + "tab", W + "cr", W + "noBreakHyphen", W + "softHyphen",
    ];

    [GeneratedRegex(@"^word/(document|footnotes|endnotes|header\d*|footer\d*)\.xml$")]
    private static partial Regex Parts();

    /// <summary>The repaired package, or the same array when no cell link holds a blank run.</summary>
    internal static byte[] Apply(byte[] docx) =>
        DocxPartRewrite.Apply(docx, name => Parts().IsMatch(name), Split);

    private static bool Split(XDocument doc)
    {
        var changed = false;
        foreach (var link in doc.Descendants(W + "hyperlink").ToList())
        {
            if (!link.Ancestors(W + "tc").Any()) continue;
            if (link.Descendants().Any(e => e.Name == W + "fldChar" || e.Name == W + "instrText" || e.Name == W + "fldSimple")) continue;

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
        e.Name == W + "r"
        && e.Elements().All(c => InvisibleContent.Contains(c.Name)
                                 || (c.Name == W + "t" && string.IsNullOrWhiteSpace(c.Value)));
}
