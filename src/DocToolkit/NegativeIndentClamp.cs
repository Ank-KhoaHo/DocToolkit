using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

namespace DocToolkit;

/// <summary>
/// Clamps a negative paragraph indent to zero so the document can be rendered.
/// </summary>
/// <remarks>
/// <b>The PDF renderer refuses any negative left or right paragraph indent, at any magnitude</b> —
/// measured, a <c>-7</c> twentieths value (0.35pt) fails exactly as <c>-720</c> (36pt) does. Ordinary
/// hanging indents are unaffected: <c>w:hanging</c> and a negative <c>w:firstLine</c> both convert.
///
/// <b>Unlike the HTML repairs, this one takes a liberty, and it was a maintainer decision rather
/// than a measurement.</b> A negative indent is legal in Word, which honours it — it is how content
/// is deliberately set outside the margin, in a letterhead or a pull-quote. Clamping pulls that
/// content back inside the margin. No browser or reference renderer says this is right; the argument
/// is simply that a document that renders slightly differently beats one that does not render.
///
/// <b>The payoff is small and was known before this was built: 2 documents out of 99.</b> The other
/// six with negative indents fail again immediately on something underneath — a glyph no available
/// font covers, usually. The indent was the first error reported, not the only one, which is why
/// clamping "imperceptible" values recovers nothing at all.
///
/// <b>It cannot change a document that renders today.</b> Every document carrying a negative left or
/// right indent is refused, so nothing that works is touched — and the package is returned by
/// reference when there is nothing to clamp.
/// </remarks>
internal static class NegativeIndentClamp
{
    /// <summary>
    /// Negative <c>w:left</c>, <c>w:right</c>, <c>w:start</c> and <c>w:end</c> on a <c>w:ind</c>.
    /// </summary>
    /// <remarks>
    /// <b>Scoped to <c>w:ind</c> deliberately.</b> Those attribute names appear on other elements —
    /// table indents, cell margins, drawing anchors — and only the paragraph indent is refused. A
    /// looser pattern would rewrite parts of the document that were never the problem, which is the
    /// difference between a repair and damage.
    ///
    /// <c>w:start</c> and <c>w:end</c> are the newer names for <c>w:left</c> and <c>w:right</c>; both
    /// spellings appear in documents produced from legacy formats.
    /// </remarks>
    private static readonly Regex Indent = new(
        @"<w:ind\b[^>]*?/>|<w:ind\b[^>]*?>", RegexOptions.Compiled);

    private static readonly Regex NegativeAttribute = new(
        @"\b(w:(?:left|right|start|end))=""-\d+""", RegexOptions.Compiled);

    /// <summary>
    /// Returns <paramref name="docx"/> with negative paragraph indents clamped to zero, or the same
    /// array when there are none.
    /// </summary>
    internal static byte[] Apply(byte[] docx)
    {
        // Cheap reject on the raw bytes. The parts are deflated, so this cannot see the attribute
        // itself - what it avoids is opening a package that has no chance of containing one.
        if (docx.AsSpan().IndexOf("word/document.xml"u8) < 0) return docx;

        var buffer = new MemoryStream();
        buffer.Write(docx, 0, docx.Length);
        buffer.Position = 0;

        var clamped = false;

        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var entry in zip.Entries
                         .Where(e => e.FullName.EndsWith(".xml", StringComparison.Ordinal))
                         .ToList())
            {
                string xml;
                using (var reader = new StreamReader(entry.Open(), Encoding.UTF8))
                    xml = reader.ReadToEnd();

                var rewritten = Indent.Replace(xml, m => NegativeAttribute.Replace(m.Value, n => $"{n.Groups[1].Value}=\"0\""));
                if (ReferenceEquals(rewritten, xml) || rewritten == xml) continue;

                // Rewritten rather than edited in place, for the same reason as
                // ListMarkerSubstitution: an entry cannot be resized downwards safely.
                var name = entry.FullName;
                entry.Delete();
                var fresh = zip.CreateEntry(name, CompressionLevel.Optimal);
                using var writer = new StreamWriter(fresh.Open(), new UTF8Encoding(false));
                writer.Write(rewritten);
                clamped = true;
            }
        }

        return clamped ? buffer.ToArray() : docx;
    }
}
