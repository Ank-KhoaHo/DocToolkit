using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace DocToolkit;

/// <summary>
/// The shared package-rewrite step behind the DOCX repairs for constructs OfficeIMO 3.5 refuses.
/// </summary>
/// <remarks>
/// <b>It exists so each repair is only the XML edit.</b> Opening the package, choosing the parts,
/// keeping whitespace, recreating entries and deciding whether anything changed are the same for
/// every repair, and the "same array when nothing changed" rule is the one a repair must not get
/// wrong - it is what makes an up-front repair unable to touch a document that renders today.
///
/// <b>Nothing is copied until an edit reports a change.</b> The parts are parsed straight from the
/// caller's bytes in read mode, so a document with nothing to repair - almost all of them - costs
/// one parse of the matching parts and no second copy of the package. Only when an edit says it
/// changed something is the package copied and the changed parts written back.
///
/// <b>Whitespace is preserved</b> (<see cref="LoadOptions.PreserveWhitespace"/>, no formatting on
/// save): a <c>w:t</c> holding only a space would otherwise change meaning. Changed entries are
/// deleted and recreated rather than overwritten, because an entry cannot be resized downwards
/// safely.
/// </remarks>
internal static class DocxPartRewrite
{
    /// <summary>
    /// Runs <paramref name="edit"/> on every XML part <paramref name="parts"/> accepts and returns
    /// the package with the edited parts written back - or <paramref name="docx"/> itself, the same
    /// array, when no edit reported a change.
    /// </summary>
    /// <param name="docx">The package. Never modified.</param>
    /// <param name="parts">Takes a part's name inside the package (<c>word/document.xml</c>).</param>
    /// <param name="edit">Edits the part in place and returns whether it changed anything.</param>
    /// <param name="requireWordMain">False skips the DOCX-only early reject, for a repair that works on any OPC package.</param>
    /// <remarks>
    /// <b>Returning the same array is load-bearing, not a nicety.</b> Callers and tests use
    /// reference equality to tell "rendering this again would render what just failed" from "this
    /// is a new document", so a repair that rebuilt an identical package would still read as a
    /// change.
    /// </remarks>
    internal static byte[] Apply(byte[] docx, Func<string, bool> parts, Func<XDocument, bool> edit, bool requireWordMain = true)
    {
        // Cheap reject on the raw bytes: a package with no main part is not a DOCX this can help.
        // A repair that also serves XLSX and PPTX passes requireWordMain: false and rejects itself.
        if (requireWordMain && docx.AsSpan().IndexOf("word/document.xml"u8) < 0) return docx;

        List<(string Name, XDocument Doc)> changed = [];

        using (var source = new MemoryStream(docx, writable: false))
        using (var read = new ZipArchive(source, ZipArchiveMode.Read))
        {
            foreach (var entry in read.Entries.Where(e => parts(e.FullName)).ToList())
            {
                XDocument doc;
                try
                {
                    // The package is untrusted input and this is a public library, and the parse
                    // happens here BEFORE OfficeIMO sees the document. A DOCTYPE is never legal in
                    // WordprocessingML, so it is refused outright (which also rules out entity
                    // expansion) and no resolver is set, so nothing is fetched.
                    using var stream = entry.Open();
                    using var reader = XmlReader.Create(stream, new XmlReaderSettings
                    {
                        DtdProcessing = DtdProcessing.Prohibit,
                        XmlResolver = null,
                    });
                    doc = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
                }
                catch (Exception ex) when (ex is XmlException or InvalidDataException)
                {
                    // A part this parser rejects (a C0 control character is not legal XML at all, and
                    // a DOCTYPE is refused) is left alone: the repair is an attempt to help, and the renderer's own
                    // error is the better answer than a parse failure from here. The same goes for a
                    // part whose compressed data is corrupt (InvalidDataException), which would
                    // otherwise escape from this up-front repair before the renderer is reached.
                    continue;
                }

                if (edit(doc)) changed.Add((entry.FullName, doc));
            }
        }

        if (changed.Count == 0) return docx;

        var buffer = new MemoryStream();
        buffer.Write(docx, 0, docx.Length);
        buffer.Position = 0;

        using (var zip = new ZipArchive(buffer, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var (name, doc) in changed)
            {
                zip.GetEntry(name)!.Delete();
                using var stream = zip.CreateEntry(name, CompressionLevel.Optimal).Open();
                using var writer = XmlWriter.Create(stream, new XmlWriterSettings
                {
                    Encoding = new UTF8Encoding(false),
                    Indent = false,
                });
                doc.Save(writer);
            }
        }

        return buffer.ToArray();
    }
}
