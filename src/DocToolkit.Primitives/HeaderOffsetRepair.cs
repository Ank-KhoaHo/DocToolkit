using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace DocToolkit;

/// <summary>
/// Raises every top margin of under 360 twips to 360 when the document has a header with content, so
/// OfficeIMO 3.5 will lay the page out. Runs only after 3.5 refuses ("header offset must not exceed
/// the top margin"), from <see cref="DocxPdfRefusalLadder"/>.
/// </summary>
/// <remarks>
/// <b>OfficeIMO 3.5 floors its own header offset at 18 pt (360 twips), and refuses the page when the
/// top margin is smaller.</b> The refusal is ArgumentException "PDF header offset must not exceed the
/// top margin when header content is enabled", and the header distance does not matter: a top margin
/// of 359 is refused with a header distance of 0, 360 renders. Word and OfficeIMO 3.2.6 accepted such
/// pages (doc-000111).
///
/// <b>The rule is deliberately wide, because it runs only on a refusal.</b> The first version
/// repaired just the last section, and doc-000111 showed why that is not enough: its sections
/// inherit the header of the first one (a section with no <c>headerReference</c> of its own uses the
/// previous section's), so a middle section with a top margin of 288 is refused too. Any header part
/// with content - text that is not whitespace, a field or a picture - means every section's top
/// margin under 360 is raised. Nothing is repaired in a document that renders: this is not an
/// up-front repair, so a refusal-free document is never edited. An empty or space-only header has no
/// content and leaves the document alone, as does a document with no header part at all.
///
/// <b>What it costs.</b> The body starts up to 0.25 inch lower than the document asked for.
///
/// <b>The footer is not repaired, and that is a measurement, not an oversight.</b> A footer with a
/// bottom margin of 360 or less is refused too ("PDF footer offset must not exceed the bottom
/// margin") but 400 renders, so the floor is not 360, and a footer distance of 0 fails again with "Running PDF content
/// cannot create another page" at any bottom margin. It has a different, larger floor that depends on
/// more than the margin, so the same fix does not apply and no repair was shipped without proof.
///
/// <b>Deliberately not done:</b> clamping the header distance down to the top margin, which the spike
/// did. Measured: a header distance of 720 over a top margin of 576, or even 8000 over 1440, renders
/// on raw 3.5.3, so editing it would change documents that render today. A negative margin is refused
/// by 3.5 for a different reason ("Parameter 'top' must be a finite non-negative number") at any
/// magnitude, so raising it to -360 would not help and it is left alone.
/// </remarks>
internal static class HeaderOffsetRepair
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>The smallest top margin 3.5 accepts under a header, in twips.</summary>
    private const int FloorTwips = 360;

    /// <summary>The repaired package, or the same array when no margin needs raising.</summary>
    internal static byte[] Apply(byte[] docx)
    {
        // Entry names sit uncompressed in the central directory: a package with no header part
        // cannot be refused for one, and is never opened.
        if (docx.AsSpan().IndexOf("word/header"u8) < 0) return docx;

        var withContent = PartsWithContent(docx);
        if (withContent.Count == 0) return docx;

        return DocxPartRewrite.Apply(docx, name => name == "word/document.xml", doc => RaiseMargins(doc, withContent));
    }

    /// <summary>The relationship ids of the header parts that hold content.</summary>
    private static HashSet<string> PartsWithContent(byte[] docx)
    {
        HashSet<string> ids = [];
        try
        {
            using var source = new MemoryStream(docx, writable: false);
            using var zip = new ZipArchive(source, ZipArchiveMode.Read);
            var relationships = zip.GetEntry("word/_rels/document.xml.rels");
            if (relationships is null) return ids;

            var rels = Load(relationships);
            foreach (var rel in rels.Descendants().Where(e => e.Name.LocalName == "Relationship"))
            {
                var type = (string?)rel.Attribute("Type") ?? "";
                if (!type.EndsWith("/header", StringComparison.Ordinal)) continue;

                var target = ((string?)rel.Attribute("Target") ?? "").TrimStart('/');
                var entry = zip.GetEntry(target.StartsWith("word/", StringComparison.Ordinal) ? target : "word/" + target);
                if (entry is null || (string?)rel.Attribute("Id") is not { } id) continue;

                if (HasContent(Load(entry))) ids.Add(id);
            }
        }
        catch (Exception ex) when (ex is XmlException or InvalidDataException)
        {
            // A package this cannot read is left to the renderer's own error.
            return [];
        }

        return ids;

        static XDocument Load(ZipArchiveEntry entry)
        {
            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            return XDocument.Load(reader);
        }
    }

    /// <summary>
    /// Text that is not whitespace, a field or a picture. An empty paragraph or a run of spaces is
    /// what 3.5 treats as "no header content" (measured); text, a page-number field and a picture
    /// each count as content (measured), and so does the legacy page-number element <c>w:pgNum</c>
    /// that a document converted from .doc carries (doc-000111). Other constructs are not counted, so
    /// a header made only of one is left as it is.
    /// </summary>
    private static bool HasContent(XDocument part) =>
        part.Descendants().Any(e =>
            (e.Name == W + "t" && !string.IsNullOrWhiteSpace(e.Value))
            || e.Name == W + "fldChar" || e.Name == W + "drawing"
            || e.Name == W + "pgNum");

    private static bool RaiseMargins(XDocument doc, HashSet<string> withContent)
    {
        var sections = doc.Descendants(W + "sectPr").ToList();
        // Sections inherit headers, so a header reference anywhere puts a header on later sections.
        if (!sections.Any(s => HasContentHeader(s, withContent))) return false;

        var changed = false;
        // Non-short-circuit |=: Raise must run on every section, not stop at the first change.
        foreach (var margins in sections.Select(s => s.Element(W + "pgMar")).OfType<XElement>())
            changed |= Raise(margins);

        return changed;
    }

    private static bool HasContentHeader(XElement section, HashSet<string> withContent) =>
        section.Elements(W + "headerReference").Any(e =>
            (string?)e.Attribute(R + "id") is { } id && withContent.Contains(id));

    private static bool Raise(XElement margins)
    {
        if (!int.TryParse((string?)margins.Attribute(W + "top"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var twips)
            || twips < 0 || twips >= FloorTwips) return false;

        margins.SetAttributeValue(W + "top", FloorTwips.ToString(CultureInfo.InvariantCulture));
        return true;
    }
}
