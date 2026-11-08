using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace DocToolkit;

/// <summary>
/// Moves runs with no visible text out of a hyperlink, splitting the link around them, so the
/// document can be rendered. Runs only after 3.5 refuses ("cannot be empty or whitespace"), from
/// <see cref="DocxPdfRefusalLadder"/>.
/// </summary>
/// <remarks>
/// <b>OfficeIMO 3.5 refuses some hyperlinks that hold a run with no text</b>
/// (ArgumentException "Parameter 'text' cannot be empty or whitespace."). A break, a tab or spaces
/// inside the link is the usual shape - Word writes it when a line break sits inside a link. 3.2.6
/// rendered it. The first version assumed only links in a table cell were refused; the corpus
/// (html-000214, html-000438) showed body-paragraph links are refused too, and a link with a break
/// next to text can render, so whether a link is refused is not a function of where it sits.
///
/// <b>Every link in the document, headers, footers, footnotes and endnotes is considered,</b> and
/// that is safe only because this runs after the refusal: a document that renders never reaches it.
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
/// <b>A document with no link holding a blank run is returned by reference</b>, which the ladder
/// reads as "this step has nothing to do".
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

    /// <summary>The repaired package, or the same array when no link holds a blank run.</summary>
    internal static byte[] Apply(byte[] docx) =>
        DocxPartRewrite.Apply(docx, name => Parts().IsMatch(name), Split);

    private static bool Split(XDocument doc)
    {
        var changed = false;
        foreach (var link in doc.Descendants(W + "hyperlink").ToList())
        {
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
