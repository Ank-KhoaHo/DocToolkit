using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace DocToolkit;

/// <summary>
/// Replaces the Symbol-font bullet glyphs Word writes into <c>numbering.xml</c> with characters the
/// PDF renderer can actually encode.
/// </summary>
/// <remarks>
/// <b>Without this, a Word document containing a bulleted list cannot be rendered to PDF at all.</b>
/// Word's default bullet is not <c>U+2022</c>; it is <c>U+F0B7</c>, a Symbol-font glyph in the
/// Unicode private-use area, stored as <c>&lt;w:lvlText w:val="."/&gt;</c>. The PDF renderer refuses
/// to encode it — and the diagnostic it raises names the source as <c>PdfListMarker</c>, so it is
/// rejecting a marker it generated itself.
///
/// <b>Measured 2026-08-16, and this is not an edge case:</b> a three-line document made in Word with
/// a default bulleted list fails. Any <c>.doc</c> or <c>.docx</c> containing a table or a bulleted
/// list is affected, which is most real documents.
///
/// <b>The replacements are measured, not chosen for looks.</b> That distinction cost a wrong first
/// attempt: <c>U+25AA BLACK SMALL SQUARE</c> is the visually correct stand-in for Word's square
/// sub-bullet and is <i>itself</i> unencodable, failing identically. Verified encodable:
/// <c>U+2022</c>, <c>U+00B7</c>, <c>U+2013</c>, <c>o</c>, <c>-</c>. Verified NOT encodable:
/// <c>U+25CF</c>, <c>U+25AA</c>, <c>U+25A0</c>, <c>U+25E6</c>, <c>U+2043</c>. <b>Anything added to
/// the map below must be measured the same way</b>, because a wrong entry turns one failure into a
/// different one.
///
/// <b>This substitutes a glyph, and that is a deliberate trade.</b> The bullet changes from a Symbol
/// <c>·</c> to a Unicode <c>•</c> — visually near-identical, and the alternative is not a faithful
/// conversion but no conversion at all. It applies only to list markers in <c>numbering.xml</c>;
/// document text is never touched.
/// </remarks>
internal static class ListMarkerSubstitution
{
    /// <summary>
    /// The private-use range that Symbol and Wingdings glyphs are mapped into. A character here is
    /// never real text — it is a font-specific glyph index — so replacing one cannot corrupt
    /// content the way replacing an ordinary character would.
    /// </summary>
    private const char PrivateUseStart = '';
    private const char PrivateUseEnd = '';

    /// <summary>Measured replacements for the glyphs Word actually emits.</summary>
    private static readonly Dictionary<char, char> Known = new()
    {
        [''] = '•',   // Symbol bullet, Word's level-1 default -> BULLET
        [''] = '·',   // Symbol square, a common sub-bullet    -> MIDDLE DOT
        [''] = '•',   // Wingdings check, used as a marker     -> BULLET
        [''] = '·',   // Wingdings filled square               -> MIDDLE DOT
    };

    /// <summary>The fallback for any other private-use glyph: a bullet is what a marker means.</summary>
    private const char Fallback = '•';

    /// <summary>
    /// Returns <paramref name="docx"/> with its list markers made renderable, or the same array
    /// when there is nothing to change.
    /// </summary>
    /// <remarks>
    /// Returning the input unchanged when no marker needs substituting matters: the overwhelming
    /// majority of documents are unaffected, and they should not pay a repackaging cost — nor risk
    /// a rewrite altering a package that was already fine.
    /// </remarks>
    internal static byte[] Apply(byte[] docx)
    {
        // Cheap pre-check on the raw bytes before opening the package at all. numbering.xml is
        // deflated inside the ZIP, so this cannot detect the glyph directly - what it detects is
        // whether the package has a numbering part worth opening.
        if (!LooksLikeItCouldHaveLists(docx)) return docx;

        byte[]? rewritten = null;
        var buffer = new MemoryStream();
        buffer.Write(docx, 0, docx.Length);
        buffer.Position = 0;

        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry("word/numbering.xml");
            if (entry is not null)
            {
                string xml;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                    xml = reader.ReadToEnd();

                var replaced = Substitute(ArialForSymbolBullets(xml));
                if (!ReferenceEquals(replaced, xml))
                {
                    // Rewritten rather than edited in place: a ZipArchiveEntry cannot be resized
                    // downwards safely, and the replacement can change the byte length.
                    entry.Delete();
                    var fresh = zip.CreateEntry("word/numbering.xml", CompressionLevel.Optimal);
                    using var writer = new StreamWriter(fresh.Open(), new UTF8Encoding(false));
                    writer.Write(replaced);
                    rewritten = Array.Empty<byte>();   // marker: the archive was modified
                }
            }
        }

        return rewritten is null ? docx : buffer.ToArray();
    }

    /// <summary>
    /// Gives every level whose marker is about to be substituted an explicit Arial font in place of
    /// its Symbol or Wingdings one, or returns the same string when no level needs it.
    /// </summary>
    /// <remarks>
    /// <b>Substituting the character is not enough on OfficeIMO 3.5.3.</b> The level keeps naming
    /// Symbol or Wingdings, those faces have no glyph for U+2022 or U+00B7, and the marker then fails
    /// to encode - the same refusal the substitution exists to avoid, one step later.
    ///
    /// <b>Explicit Arial, not "remove the font".</b> Measured on doc-000966: with the <c>rFonts</c>
    /// removed the marker inherits the font of the paragraph mark, and that can be Wingdings too.
    /// All four slots (<c>ascii</c>, <c>hAnsi</c>, <c>cs</c>, <c>eastAsia</c>) are set so no slot
    /// falls back, and a theme font attribute is dropped because it would override the name.
    ///
    /// <b>It costs the marker its face.</b> The bullet is drawn in Arial instead of Symbol; the shape
    /// is the same. Only levels whose own text holds a private-use glyph are touched, so a level that
    /// already carries a real bullet keeps its font. A part this parser rejects is left to the
    /// character substitution alone.
    /// </remarks>
    private static string ArialForSymbolBullets(string xml)
    {
        if (!ContainsPrivateUse(xml)) return xml;

        XDocument doc;
        try
        {
            using var reader = XmlReader.Create(new StringReader(xml), new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
            });
            doc = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        catch (XmlException)
        {
            return xml;
        }

        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var changed = false;
        foreach (var level in doc.Descendants(w + "lvl"))
        {
            var marker = (string?)level.Element(w + "lvlText")?.Attribute(w + "val");
            if (marker is null || !ContainsPrivateUse(marker)) continue;

            var fonts = level.Element(w + "rPr")?.Element(w + "rFonts");
            if (fonts is null || !fonts.Attributes().Any(a => IsSymbolFace(a.Value))) continue;

            fonts.ReplaceAttributes(
                new XAttribute(w + "ascii", "Arial"), new XAttribute(w + "hAnsi", "Arial"),
                new XAttribute(w + "cs", "Arial"), new XAttribute(w + "eastAsia", "Arial"));
            changed = true;
        }

        if (!changed) return xml;
        var declaration = doc.Declaration is null ? "" : doc.Declaration + "\n";
        return declaration + doc.ToString(SaveOptions.DisableFormatting);

        static bool IsSymbolFace(string name) =>
            name.StartsWith("Symbol", StringComparison.OrdinalIgnoreCase)
            || name.StartsWith("Wingdings", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ContainsPrivateUse(string text)
    {
        foreach (var c in text)
            if (c >= PrivateUseStart && c <= PrivateUseEnd) return true;
        return false;
    }

    /// <summary>Replaces every private-use character, or returns the same string when there are none.</summary>
    private static string Substitute(string xml)
    {
        var first = -1;
        for (var i = 0; i < xml.Length; i++)
        {
            if (xml[i] >= PrivateUseStart && xml[i] <= PrivateUseEnd) { first = i; break; }
        }

        if (first < 0) return xml;

        var sb = new StringBuilder(xml.Length);
        sb.Append(xml, 0, first);
        for (var i = first; i < xml.Length; i++)
        {
            var c = xml[i];
            sb.Append(c >= PrivateUseStart && c <= PrivateUseEnd
                ? Known.TryGetValue(c, out var mapped) ? mapped : Fallback
                : c);
        }
        return sb.ToString();
    }

    /// <summary>
    /// Whether the package is worth opening — i.e. whether it contains a numbering part at all.
    /// </summary>
    /// <remarks>
    /// Matches the entry NAME in the ZIP's central directory, which is stored uncompressed. A
    /// document with no list has no numbering.xml and is returned untouched without the archive
    /// ever being opened.
    /// </remarks>
    private static bool LooksLikeItCouldHaveLists(byte[] docx)
    {
        ReadOnlySpan<byte> needle = "word/numbering.xml"u8;
        return docx.AsSpan().IndexOf(needle) >= 0;
    }
}
