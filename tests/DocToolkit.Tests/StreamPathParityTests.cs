namespace DocToolkit.Tests;

/// <summary>
/// A <c>Stream</c> overload must answer the same as its <c>byte[]</c> sibling.
/// </summary>
/// <remarks>
/// <b>They did not, and it was measured rather than suspected.</b> Until 2026-08-20 the
/// <c>Stream</c>-destination PDF overloads wrote straight onto the caller's stream as OfficeIMO
/// laid the document out. That reads like a virtue and cost correctness: a repair that retries a
/// failed render cannot un-write bytes already sent, so those overloads could not go through
/// <see cref="HtmlForPdf"/> at all and applied <b>no repairs</b>.
///
/// The two measurements that established it:
///
/// <list type="bullet">
/// <item>a page whose internal links use <c>&lt;a name&gt;</c> - 27 of 181 real .gov pages by
/// <see cref="HtmlAnchorRepair"/>'s own count - converted through <c>ConvertAsync(html)</c> and was
/// refused through <c>ConvertAsync(html, destination)</c>;</item>
/// <item><b>4 of 99 real Word documents</b> converted through <c>DocxToPdfConverter.Convert</c> and
/// were refused through its stream overload, every one a negative indent the clamp recovers.</item>
/// </list>
///
/// <b>The fix buffers, on the maintainer's decision, and that is the trade this file pins.</b> One
/// PDF of memory buys identical behaviour everywhere and a failure that leaves the destination
/// untouched instead of carrying a truncated PDF. The <c>Stream</c> overloads were never a memory
/// optimisation - <c>DrainAsync</c> buffers the source regardless.
///
/// <b>Parity is asserted on the SAME input through BOTH paths</b>, rather than each path being
/// tested against its own expectation. Two separately-correct-looking tests are exactly how the
/// divergence survived.
/// </remarks>
public class StreamPathParityTests
{
    /// <summary>A page whose only fault is the one <see cref="HtmlAnchorRepair"/> fixes.</summary>
    private const string NeedsAnchorRepair =
        "<html><body><p><a href=\"#part2\">Jump</a></p>"
        + "<p><a name=\"part2\"></a>Part two.</p></body></html>";

    private static async Task<byte[]> ViaStream(Func<Stream, Task> write)
    {
        using var destination = new MemoryStream();
        await write(destination);
        return destination.ToArray();
    }

    // ---- HTML to PDF --------------------------------------------------------------------------------

    [Fact]
    public async Task HtmlToPdf_StreamAgreesWithBytes_OnAPageNeedingARepair()
    {
        var viaBytes = await HtmlToPdfConverter.ConvertAsync(NeedsAnchorRepair);
        var viaStream = await ViaStream(d => HtmlToPdfConverter.ConvertAsync(NeedsAnchorRepair, d));

        // Both must succeed. Before the fix the second threw
        // "PDF bookmark link target 'part2' was not found".
        Assert.True(PdfProbe.IsPdf(viaBytes));
        Assert.True(PdfProbe.IsPdf(viaStream));
        Assert.Equal(PdfProbe.MediaBoxes(viaBytes).Count, PdfProbe.MediaBoxes(viaStream).Count);
    }

    [Fact]
    public async Task HtmlToPdf_StreamCarriesTheSameText()
    {
        // Not just "both produced a PDF": a stream path that quietly produced a blank page would
        // satisfy the assertion above.
        const string html = "<h1>PARITY-TITLE</h1><p>PARITY-BODY</p>";

        var viaStream = await ViaStream(d => HtmlToPdfConverter.ConvertAsync(html, d));
        var text = PdfProbe.ExtractText(viaStream);

        Assert.Contains("PARITY-TITLE", text, StringComparison.Ordinal);
        Assert.Contains("PARITY-BODY", text, StringComparison.Ordinal);
    }

    // ---- DOCX to PDF --------------------------------------------------------------------------------

    /// <summary>A .docx carrying the negative indent the renderer refuses and the clamp repairs.</summary>
    private static byte[] NegativeIndentDocument()
    {
        var docx = DocxEditor.Create([DocxBlock.Paragraph("Text long enough to wrap onto a second line.")]);

        using var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        ms.Position = 0;

        using (var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry("word/document.xml")!;
            string xml;
            using (var reader = new StreamReader(entry.Open(), System.Text.Encoding.UTF8)) xml = reader.ReadToEnd();

            var paragraph = xml.IndexOf("<w:p ", StringComparison.Ordinal);
            if (paragraph < 0) paragraph = xml.IndexOf("<w:p>", StringComparison.Ordinal);
            var afterOpen = xml.IndexOf('>', paragraph) + 1;
            xml = xml[..afterOpen] + "<w:pPr><w:ind w:right=\"-720\"/></w:pPr>" + xml[afterOpen..];

            entry.Delete();
            var fresh = zip.CreateEntry("word/document.xml", System.IO.Compression.CompressionLevel.Optimal);
            using var writer = new StreamWriter(fresh.Open(), new System.Text.UTF8Encoding(false));
            writer.Write(xml);
        }

        return ms.ToArray();
    }

    [Fact]
    public async Task DocxToPdf_StreamAgreesWithBytes_OnADocumentNeedingTheClamp()
    {
        var docx = NegativeIndentDocument();

        var viaBytes = DocxToPdfConverter.Convert(docx);
        var viaStream = await ViaStream(async d =>
        {
            using var source = new MemoryStream(docx);
            await DocxToPdfConverter.ConvertAsync(source, d);
        });

        Assert.True(PdfProbe.IsPdf(viaBytes));
        Assert.True(PdfProbe.IsPdf(viaStream));
        Assert.Contains("second line", PdfProbe.ExtractText(viaStream), StringComparison.OrdinalIgnoreCase);
    }

    // ---- what buffering bought ----------------------------------------------------------------------

    [Fact]
    public async Task AFailedConversionLeavesTheDestinationUntouched()
    {
        // The property the old straight-through write could not offer, and its doc comments said so:
        // "a failure part-way leaves whatever had already been produced on destination". A caller
        // streaming to an HTTP response body got a truncated PDF on the wire.
        using var destination = new MemoryStream();
        destination.Write("SENTINEL"u8);
        var before = destination.ToArray();

        await Assert.ThrowsAsync<DocumentConversionException>(
            () => HtmlToPdfConverter.ConvertAsync("<p>a\u0010b</p>", destination));

        Assert.Equal(before, destination.ToArray());
    }

    [Fact]
    public async Task TheDestinationIsNotDisposedOrSought()
    {
        // Buffering must not have quietly acquired ownership of the caller's stream.
        using var destination = new MemoryStream();
        destination.Write("PREFIX"u8);

        await HtmlToPdfConverter.ConvertAsync("<p>ok</p>", destination);

        Assert.True(destination.CanWrite, "the destination must still be usable");
        Assert.StartsWith("PREFIX", System.Text.Encoding.ASCII.GetString(destination.ToArray()[..6]), StringComparison.Ordinal);
    }

    // ---- XLSX: the guard, not the repairs ------------------------------------------------------------

    [Fact]
    public async Task XlsxToPdf_StreamRefusesALegacyWorkbookLikeTheBytePath()
    {
        // Fixed differently and deliberately: this path has no repair or retry, so only the
        // compound-file guard was missing. Its absence made a documented denial of service - a
        // legacy .xls that "did not finish in ten minutes" - reachable through the overload an
        // upload endpoint uses.
        var legacy = new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

        var fromBytes = Assert.Throws<DocumentConversionException>(() => XlsxToPdfConverter.Convert(legacy));
        var fromStream = await Assert.ThrowsAsync<DocumentConversionException>(async () =>
        {
            using var source = new MemoryStream(legacy);
            using var destination = new MemoryStream();
            await XlsxToPdfConverter.ConvertAsync(source, destination);
        });

        Assert.Contains("not an .xlsx package", fromBytes.Message, StringComparison.Ordinal);
        Assert.Equal(fromBytes.Message, fromStream.Message);
    }
}
