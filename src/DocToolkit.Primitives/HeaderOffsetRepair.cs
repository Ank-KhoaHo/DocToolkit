using System.Globalization;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace DocToolkit;

/// <summary>
/// Raises a top margin of under 360 twips to 360 when the document has a header with content, so
/// OfficeIMO 3.5 will lay the page out.
/// </summary>
/// <remarks>
/// <b>OfficeIMO 3.5 floors its own header offset at 18 pt (360 twips), and refuses the page when the
/// top margin is smaller.</b> The refusal is ArgumentException "PDF header offset must not exceed the
/// top margin when header content is enabled", and the header distance does not matter: a top margin
/// of 359 is refused with a header distance of 0, 360 renders. Word and OfficeIMO 3.2.6 accepted such
/// pages (doc-000111).
///
/// <b>Measured on raw 3.5.3, and the conditions are what keep this from touching documents that
/// render today.</b> The refusal needs all of: the LAST body section (the one that sets the page),
/// a <c>default</c> header reference, and a part with content - text that is not whitespace, a
/// field or a picture. An empty header, a space-only header, a <c>first</c> or <c>even</c> header and
/// a header on an earlier section all render as they are, so none of them is edited.
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
    /// each count as content (measured). Other constructs are not counted, so a header made only of
    /// one is left as it is.
    /// </summary>
    private static bool HasContent(XDocument part) =>
        part.Descendants().Any(e =>
            (e.Name == W + "t" && !string.IsNullOrWhiteSpace(e.Value))
            || e.Name == W + "fldChar" || e.Name == W + "drawing");

    private static bool RaiseMargins(XDocument doc, HashSet<string> withContent)
    {
        // The last body-level sectPr is the section that sets the page. Earlier ones sit in a
        // paragraph's properties and render with any margin (measured).
        var section = doc.Root?.Element(W + "body")?.Elements(W + "sectPr").LastOrDefault();
        var margins = section?.Element(W + "pgMar");
        if (section is null || margins is null) return false;

        return HasDefaultHeader(section, withContent) && Raise(margins);
    }

    private static bool HasDefaultHeader(XElement section, HashSet<string> withContent) =>
        section.Elements(W + "headerReference").Any(e =>
            ((string?)e.Attribute(W + "type") ?? "default") == "default"
            && (string?)e.Attribute(R + "id") is { } id && withContent.Contains(id));

    private static bool Raise(XElement margins)
    {
        if (!int.TryParse((string?)margins.Attribute(W + "top"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var twips)
            || twips < 0 || twips >= FloorTwips) return false;

        margins.SetAttributeValue(W + "top", FloorTwips.ToString(CultureInfo.InvariantCulture));
        return true;
    }
}
