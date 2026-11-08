using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace DocToolkit;

/// <summary>
/// Removes C0 and C1 control characters from document text so the document can be rendered.
/// </summary>
/// <remarks>
/// <b>No font has a glyph for a control character, and OfficeIMO 3.5 refuses the run when fonts
/// are supplied.</b> The usual source is a mis-decoded page (U+009D where a curly quote was) pasted
/// into Word. The character is invisible in Word, so removing it loses nothing a reader could see.
///
/// <b>What goes:</b> everything below U+0020 except tab, line feed and carriage return, and
/// U+007F to U+009F. Only <c>w:t</c> text is touched, in the body, headers, footers, footnotes and
/// endnotes - the text the renderer lays out. Field instructions, attributes and properties are
/// left alone.
///
/// <b>Unlike the table repairs it is not gated on a refusal</b>, so it also edits documents that
/// render today. The edit is invisible by construction (the characters have no glyph), and a
/// document without one comes back by reference.
/// </remarks>
internal static partial class ControlCharacterStrip
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [GeneratedRegex(@"^word/(document|footnotes|endnotes|header\d*|footer\d*)\.xml$")]
    private static partial Regex TextParts();

    /// <summary>The repaired package, or the same array when no text holds a control character.</summary>
    internal static byte[] Apply(byte[] docx) =>
        DocxPartRewrite.Apply(docx, name => TextParts().IsMatch(name), Strip);

    private static bool Strip(XDocument doc)
    {
        var changed = false;
        foreach (var t in doc.Descendants(W + "t"))
        {
            var value = t.Value;
            if (!value.Any(IsControl)) continue;

            t.Value = new string(value.Where(c => !IsControl(c)).ToArray());
            changed = true;
        }

        return changed;
    }

    private static bool IsControl(char c) =>
        c is (< ' ' and not '\t' and not '\n' and not '\r') or (>= '\u007F' and <= '\u009F');
}
