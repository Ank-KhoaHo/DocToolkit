using System.Xml.Linq;

namespace DocToolkit;

/// <summary>Follows every list marker with a space instead of a tab.</summary>
/// <remarks>
/// <b>Runs only after 3.5 refuses with "Inline element width exceeds the available paragraph
/// width".</b> A marker followed by a tab runs to the next tab stop, and a paragraph style can clear
/// the numbering level's own stop and set another one far to the right. Measured on doc-000006: a
/// numbered paragraph in a 4,698-twip table cell, whose style clears stop 1080 and sets a centre stop
/// at 4680, so the tab ran past the 4,482-twip text area and 3.5 refused the marker. Word lets it
/// overflow.
///
/// <b>Every level of every list is changed, not the one that was refused.</b> The refusal does not
/// say which paragraph it is, so a level cannot be picked. A document that reaches this step has
/// already failed, and a space after the marker is a small change next to losing the document. An
/// absent <c>w:suff</c> means a tab, so one is added (before <c>w:lvlText</c>, where the schema puts
/// it); a <c>space</c> or <c>nothing</c> suffix is left as it was.
/// </remarks>
internal static class ListMarkerSuffixSpace
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>The repaired package, or the same array when no level followed its marker with a tab.</summary>
    internal static byte[] Apply(byte[] docx) =>
        DocxPartRewrite.Apply(docx, name => name == "word/numbering.xml", SuffixSpace);

    private static bool SuffixSpace(XDocument doc)
    {
        var changed = false;
        foreach (var level in doc.Descendants(W + "lvl"))
        {
            var suffix = level.Element(W + "suff");
            if (suffix is null)
            {
                suffix = new XElement(W + "suff", new XAttribute(W + "val", "space"));
                if (level.Element(W + "lvlText") is { } text) text.AddBeforeSelf(suffix);
                else level.Add(suffix);
                changed = true;
            }
            else if ((string?)suffix.Attribute(W + "val") is null or "tab")
            {
                suffix.SetAttributeValue(W + "val", "space");
                changed = true;
            }
        }

        return changed;
    }
}
