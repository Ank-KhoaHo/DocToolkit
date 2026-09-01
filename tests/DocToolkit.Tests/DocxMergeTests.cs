using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// A107: <see cref="DocxEditor.Merge(IEnumerable{byte[]})"/>, the DOCX counterpart of
/// <c>PdfEditor.Merge</c>.
/// </summary>
/// <remarks>
/// Two of these tests pin **documented limitations** rather than desirable behaviour, and that is
/// deliberate. Both were measured before any code was written (see the A107 spec), and a
/// limitation stated only in prose is a claim nothing verifies — the failure this repository keeps
/// correcting. If OfficeIMO's behaviour changes, these fail and the documentation gets fixed with
/// them, instead of quietly becoming wrong.
/// </remarks>
public class DocxMergeTests
{
    private static byte[] Doc(string text) =>
        DocxEditor.Create(new[] { DocxBlock.Paragraph(text) });

    private static Body BodyOf(byte[] docx)
    {
        using var ms = new MemoryStream(docx, writable: false);
        using var doc = WordprocessingDocument.Open(ms, false);
        return (Body)doc.MainDocumentPart!.Document!.Body!.CloneNode(true);
    }

    /// <summary>A document whose Heading1 carries an explicit colour, so a collision is visible.</summary>
    private static byte[] DocWithHeadingColour(string text, string hex)
    {
        var docx = DocxEditor.Create(new[] { DocxBlock.Heading(text, 1) });

        using var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        ms.Position = 0;

        using (var w = WordprocessingDocument.Open(ms, true))
        {
            var part = w.MainDocumentPart!.StyleDefinitionsPart
                ?? w.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
            part.Styles ??= new Styles();
            part.Styles.Elements<Style>().FirstOrDefault(s => s.StyleId?.Value == "Heading1")?.Remove();
            part.Styles.AppendChild(new Style(
                new StyleName { Val = "heading 1" },
                new StyleRunProperties(new Color { Val = hex }))
            {
                Type = StyleValues.Paragraph,
                StyleId = "Heading1",
            });
            part.Styles.Save();
        }

        return ms.ToArray();
    }

    [Fact]
    public void Merge_JoinsDocumentsInOrder()
    {
        var merged = DocxEditor.Merge(new[] { Doc("FIRST"), Doc("SECOND"), Doc("THIRD") });

        // The ORDER, not just the presence of all three: a merge that reversed or reordered them
        // would still contain every marker.
        var text = DocxEditor.ExtractText(merged);
        Assert.Equal(
            new[] { "FIRST", "SECOND", "THIRD" },
            new[] { "FIRST", "SECOND", "THIRD" }.OrderBy(m => text.IndexOf(m, StringComparison.Ordinal)).ToArray());
        Assert.Contains("FIRST", text, StringComparison.Ordinal);
        Assert.Contains("THIRD", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Merge_OfASingleDocument_ReturnsItsContent()
    {
        var merged = DocxEditor.Merge(new[] { Doc("ONLY") });

        Assert.Contains("ONLY", DocxEditor.ExtractText(merged), StringComparison.Ordinal);
    }

    [Fact]
    public void Merge_DoesNotModifyItsInputs()
    {
        var first = Doc("FIRST");
        var second = Doc("SECOND");
        var firstBefore = (byte[])first.Clone();
        var secondBefore = (byte[])second.Clone();

        DocxEditor.Merge(new[] { first, second });

        Assert.Equal(firstBefore, first);
        Assert.Equal(secondBefore, second);
    }

    [Fact]
    public void Merge_KeepsTheSectionPropertiesLastInTheBody()
    {
        // The corruption shape CLAUDE.md records: a body whose last child is not a sectPr is
        // schema-valid, reads back fine, and makes Word offer to repair the file. Nothing else
        // here would notice.
        var body = BodyOf(DocxEditor.Merge(new[] { Doc("FIRST"), Doc("SECOND") }));

        Assert.IsType<SectionProperties>(body.ChildElements[^1]);
    }

    [Fact]
    public void Merge_GivesEachDocumentItsOwnSection_SoPageSetupSurvives()
    {
        var a4 = DocxEditor.Create(new[] { DocxBlock.Paragraph("A") }, PageSetup.A4);
        var letter = DocxEditor.Create(new[] { DocxBlock.Paragraph("L") }, PageSetup.Letter);

        var body = BodyOf(DocxEditor.Merge(new[] { a4, letter }));
        var sizes = body.Descendants<PageSize>()
            .Select(p => (W: p.Width?.Value, H: p.Height?.Value))
            .ToList();

        // Two sections with DIFFERENT dimensions - a merge that flattened page setup would leave
        // one, or two identical ones.
        Assert.Equal(2, sizes.Count);
        Assert.NotEqual(sizes[0], sizes[1]);
    }

    [Fact]
    public void Merge_WhenTwoDocumentsDefineTheSameStyleDifferently_TheFirstWins()
    {
        // A DOCUMENTED LIMITATION, pinned so it cannot drift out of the documentation silently.
        // Measured before implementing: the appended content adopts the target's definition, with
        // no error and no lost text - only a changed appearance. This is what Aspose's
        // ImportFormatMode exists to decide, and this package does not expose that choice.
        var red = DocWithHeadingColour("FIRST-HEADING", "FF0000");
        var blue = DocWithHeadingColour("SECOND-HEADING", "0000FF");

        var merged = DocxEditor.Merge(new[] { red, blue });

        using var ms = new MemoryStream(merged, writable: false);
        using var doc = WordprocessingDocument.Open(ms, false);
        var headings = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!
            .Elements<Style>()
            .Where(s => s.StyleId?.Value == "Heading1")
            .Select(s => s.StyleRunProperties?.GetFirstChild<Color>()?.Val?.Value)
            .ToList();

        Assert.Single(headings);
        Assert.Equal("FF0000", headings[0]);

        // And the text is all still there - the loss is formatting only, which is exactly what
        // makes it easy to miss.
        var text = DocxEditor.ExtractText(merged);
        Assert.Contains("FIRST-HEADING", text, StringComparison.Ordinal);
        Assert.Contains("SECOND-HEADING", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Merge_NullOrEmptyInput_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => DocxEditor.Merge(null!));

        Assert.Equal("docx", Assert.Throws<ArgumentException>(
            () => DocxEditor.Merge(Array.Empty<byte[]>())).ParamName);

        Assert.Equal("docx", Assert.Throws<ArgumentNullException>(
            () => DocxEditor.Merge(new byte[]?[] { Doc("A"), null }!)).ParamName);

        Assert.Equal("docx", Assert.Throws<ArgumentException>(
            () => DocxEditor.Merge(new[] { Doc("A"), Array.Empty<byte>() })).ParamName);
    }

    [Fact]
    public async Task MergeAsync_MatchesTheByteArrayOverload()
    {
        using var first = new MemoryStream(Doc("FIRST"), writable: false);
        using var second = new MemoryStream(Doc("SECOND"), writable: false);
        using var destination = new MemoryStream();

        await DocxEditor.MergeAsync(new Stream[] { first, second }, destination);

        var text = DocxEditor.ExtractText(destination.ToArray());
        Assert.Contains("FIRST", text, StringComparison.Ordinal);
        Assert.Contains("SECOND", text, StringComparison.Ordinal);
    }

    [Fact]
    public async Task MergeAsync_NoSources_ThrowsAgainstItsOwnParameter()
    {
        using var destination = new MemoryStream();

        // `sources`, never `docx` - the caller of this overload never passed a `docx`. Same rule
        // CLAUDE.md records for the file-path overloads.
        var ex = await Assert.ThrowsAsync<ArgumentException>(
            () => DocxEditor.MergeAsync(Array.Empty<Stream>(), destination));
        Assert.Equal("sources", ex.ParamName);
    }
}
