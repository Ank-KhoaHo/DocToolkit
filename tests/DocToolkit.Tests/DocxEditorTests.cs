using DocumentFormat.OpenXml.Wordprocessing;
using DocToolkit;
using Xunit;
using static DocToolkit.Tests.DocxFixtures;

namespace DocToolkit.Tests;

public class DocxEditorTests
{
    private static Dictionary<string, string> Map(string key, string value) => new() { [key] = value };

    [Fact]
    public async Task ReplaceText_SubstitutesPlaceholders()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<p>Dear {{name}}, your balance is {{balance}}.</p>");

        var edited = DocxEditor.ReplaceText(docx, new Dictionary<string, string>
        {
            ["{{name}}"] = "Contoso Ltd",
            ["{{balance}}"] = "4,250.00",
        });

        var text = DocxEditor.ExtractText(edited);
        Assert.Contains("Contoso Ltd", text);
        Assert.Contains("4,250.00", text);
        Assert.DoesNotContain("{{name}}", text);
        Assert.DoesNotContain("{{balance}}", text);
    }

    [Fact]
    public async Task ReplaceText_LeavesTheDocumentOpenable()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync("<p>Hello {{who}}</p>");

        var edited = DocxEditor.ReplaceText(docx,
            new Dictionary<string, string> { ["{{who}}"] = "world" });

        // Still a valid package, and still renders.
        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, edited.Take(4).ToArray());
        Assert.Contains("world", PdfProbe.ExtractText(DocxToPdfConverter.Convert(edited)));
    }

    [Fact]
    public async Task ExtractText_ReturnsDocumentText()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync("<h1>Title</h1><p>Body copy.</p>");
        var text = DocxEditor.ExtractText(docx);

        Assert.Contains("Title", text);
        Assert.Contains("Body copy.", text);
    }

    // ---------------------------------------------------------------------------------------
    // Split runs. HtmlToOpenXml emits one w:t per paragraph, so the fixtures above can never
    // reach the multi-run path - the whole reason DocxEditor does not just call Replace per run.
    // These build the runs by hand instead.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void ReplaceText_SubstitutesPlaceholderSplitAcrossThreeRuns()
    {
        // "Dear {{name}}, hi" laid out as three w:t runs with the placeholder straddling both
        // boundaries - exactly what Word produces after an edit inside a word.
        var docx = Build(P(R("Dear {{na"), R("m"), R("e}}, hi")));

        var edited = DocxEditor.ReplaceText(docx, Map("{{name}}", "Contoso Ltd"));

        var paragraph = ReadBody(edited, b => b.Descendants<Paragraph>().First());
        Assert.Equal("Dear Contoso Ltd, hi", OwnText(paragraph));
    }

    [Fact]
    public void ReplaceText_KeepsTheReplacementInTheRunThatOwnsTheMatchStart()
    {
        // Documented contract: the value lands in the run holding the first matched character,
        // and every run the match spans keeps whatever text the match did not cover.
        var docx = Build(P(R("{{na"), R("me}} tail")));

        var edited = DocxEditor.ReplaceText(docx, Map("{{name}}", "VALUE"));

        var runs = ReadBody(edited, b => OwnRunTexts(b.Descendants<Paragraph>().First()));
        Assert.Equal(new[] { "VALUE", " tail" }, runs);
    }

    [Fact]
    public void ReplaceText_DoesNotImposeTheFirstRunsFormattingOnTheParagraph()
    {
        // Blocker 3: collapsing every run onto run 0 turned the whole paragraph bold.
        var docx = Build(P(R("Bold ", bold: true), R("plain {{x}} tail")));

        var edited = DocxEditor.ReplaceText(docx, Map("{{x}}", "VALUE"));

        var runs = ReadBody(edited, b => b.Descendants<Paragraph>().First()
            .Elements<Run>()
            .Select(r => (Text: string.Concat(r.Elements<Text>().Select(t => t.Text)),
                          Bold: r.RunProperties?.Bold is not null))
            .ToArray());

        Assert.Equal(2, runs.Length);
        Assert.Equal(("Bold ", true), runs[0]);
        Assert.Equal(("plain VALUE tail", false), runs[1]);
    }

    [Fact]
    public void ReplaceText_LeavesHyperlinksIntact()
    {
        // w:hyperlink is a sibling of w:r under w:p, so the old merge swallowed its runs and left
        // a dead, invisible link behind.
        var docx = Build(P(
            R("See {{doc}} at "),
            new Hyperlink(R("example site")) { Anchor = "top" },
            R(" now")));

        var edited = DocxEditor.ReplaceText(docx, Map("{{doc}}", "the spec"));

        var (linkText, paragraphText) = ReadBody(edited, b =>
        {
            var paragraph = b.Descendants<Paragraph>().First();
            return (paragraph.Descendants<Hyperlink>().Single().InnerText, OwnText(paragraph));
        });

        Assert.Equal("example site", linkText);
        Assert.Equal("See the spec at example site now", paragraphText);
    }

    // ---------------------------------------------------------------------------------------
    // Text boxes (Blocker 1). w:txbxContent nests a whole w:p inside a run of the outer w:p, so
    // paragraph.Descendants<Text>() reaches into it and the merge hoisted the text box's content
    // into the outer paragraph.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void ReplaceText_LeavesAnUnrelatedTextBoxByteIdentical()
    {
        var docx = Build(P(R("Outer {{a}} end"), TextBoxRun("Confidential sidebar")));
        var before = ReadBody(docx, b => b.Descendants<TextBoxContent>().Single().OuterXml);

        var edited = DocxEditor.ReplaceText(docx, Map("{{a}}", "X"));

        var after = ReadBody(edited, b => b.Descendants<TextBoxContent>().Single().OuterXml);
        Assert.Equal(before, after);

        var outer = ReadBody(edited, b => OwnText(b.Descendants<Paragraph>().First()));
        Assert.Equal("Outer X end", outer);
    }

    [Fact]
    public void ReplaceText_SubstitutesInsideATextBox()
    {
        var docx = Build(P(R("Outer text"), TextBoxRun("Sidebar {{b}}")));

        var edited = DocxEditor.ReplaceText(docx, Map("{{b}}", "Y"));

        Assert.Equal("Sidebar Y", ReadBody(edited, b => b.Descendants<TextBoxContent>().Single().InnerText));
        Assert.Equal("Outer text", ReadBody(edited, b => OwnText(b.Descendants<Paragraph>().First())));
        Assert.Empty(Validate(edited));
    }

    // ---------------------------------------------------------------------------------------
    // Headers and footers (Blocker 4).
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void ReplaceText_SubstitutesInHeadersAndFooters()
    {
        var docx = Build("Header {{h}}", "Footer {{f}}", P(R("Body {{b}}")));

        var edited = DocxEditor.ReplaceText(docx, new Dictionary<string, string>
        {
            ["{{h}}"] = "H-VALUE",
            ["{{f}}"] = "F-VALUE",
            ["{{b}}"] = "B-VALUE",
        });

        var (header, footer, body) = Read(edited, main => (
            main.HeaderParts.Single().Header!.InnerText,
            main.FooterParts.Single().Footer!.InnerText,
            main.Document!.Body!.InnerText));

        Assert.Equal("Header H-VALUE", header);
        Assert.Equal("Footer F-VALUE", footer);
        Assert.Contains("Body B-VALUE", body);
    }

    [Fact]
    public void ExtractText_OmitsHeadersAndFootersByDefaultAndIncludesThemOnRequest()
    {
        var docx = Build("Page header", "Page footer", P(R("Body copy.")));

        var bodyOnly = DocxEditor.ExtractText(docx);
        Assert.Contains("Body copy.", bodyOnly);
        Assert.DoesNotContain("Page header", bodyOnly);
        Assert.DoesNotContain("Page footer", bodyOnly);

        var everything = DocxEditor.ExtractText(docx, includeHeadersAndFooters: true);
        Assert.Contains("Body copy.", everything);
        Assert.Contains("Page header", everything);
        Assert.Contains("Page footer", everything);
    }

    // ---------------------------------------------------------------------------------------
    // Block boundaries (A26). ExtractTextCore returned Body.InnerText, which concatenates every
    // descendant text node with NO separator, so "Title" and "Body" came back as one token.
    //
    // Nineteen tests in this file and 845 elsewhere missed it, and the reason is worth keeping:
    // every assertion touching ExtractText was structurally unable to see it. Assert.Contains is
    // substring-based, so "First." and "Second." are both found in "First.Second."; the ordering
    // check compares two IndexOf results, which also holds; and the exact-equality assertions
    // compare ExtractText(a) against ExtractText(b) - the method against ITSELF, which is true
    // whatever it does. The tests below assert the literal string, which is the only shape that
    // can fail.
    //
    // Separator choice is not arbitrary: '\n' is already what this same method puts between the
    // body and each header/footer part, so the body path was the odd one out. '\t' between cells
    // matches what Word's own "save as plain text" writes.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void ExtractText_SeparatesBlocks_SoAdjacentWordsDoNotFuse()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Heading("Title", 1),
            DocxBlock.Paragraph("Body text."),
        });

        var text = DocxEditor.ExtractText(docx);

        Assert.DoesNotContain("TitleBody", text, StringComparison.Ordinal);
        Assert.Equal("Title\nBody text.", text);
    }

    [Fact]
    public void ExtractText_SeparatesTableCells_SoAdjacentCellsDoNotFuse()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Table(
                new[] { "Region", "Q1" },
                new[] { new object?[] { "EMEA", 1200 } }),
        });

        var text = DocxEditor.ExtractText(docx);

        Assert.DoesNotContain("RegionQ1", text, StringComparison.Ordinal);
        Assert.Equal("Region\tQ1\nEMEA\t1200", text);
    }

    // ---------------------------------------------------------------------------------------
    // Error handling (I-6).
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void ExtractText_WrapsCorruptInputInDocumentConversionException()
    {
        Assert.Throws<DocumentConversionException>(
            () => DocxEditor.ExtractText(new byte[] { 1, 2, 3, 4, 5 }));
    }

    [Fact]
    public void ExtractText_RejectsEmptyInput()
    {
        Assert.Throws<ArgumentException>(() => DocxEditor.ExtractText(Array.Empty<byte>()));
    }

    // ---------------------------------------------------------------------------------------
    // File-path overloads.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public async Task ReplaceTextAsync_FromFileToFile_MatchesTheByteArrayOverload()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("Dear {{customer}}")));
        var replacements = new Dictionary<string, string> { ["{{customer}}"] = "Contoso Ltd" };

        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, docx);

        await DocxEditor.ReplaceTextAsync(input.Path, output.Path, replacements);

        // B16: the substitution itself, asserted against literals. The parity line below cannot
        // see a ReplaceText that stopped substituting - both sides would carry "{{customer}}" and
        // still be equal.
        var written = DocxEditor.ExtractText(await File.ReadAllBytesAsync(output.Path));
        Assert.Contains("Dear Contoso Ltd", written, StringComparison.Ordinal);
        Assert.DoesNotContain("{{customer}}", written, StringComparison.Ordinal);

        Assert.Equal(
            DocxEditor.ExtractText(DocxEditor.ReplaceText(docx, replacements)),
            DocxEditor.ExtractText(await File.ReadAllBytesAsync(output.Path)));
    }

    [Fact]
    public async Task ExtractTextAsync_FromFile_MatchesTheByteArrayOverload()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("Hello from disk")));

        using var input = new TempFile();
        await File.WriteAllBytesAsync(input.Path, docx);

        Assert.Equal(DocxEditor.ExtractText(docx), await DocxEditor.ExtractTextAsync(input.Path));
    }

    [Fact]
    public async Task FillRowsAsync_FromFileToFile_ExpandsTheTemplateRow()
    {
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.Row(DocxFixtures.R("Description")),
            DocxFixtures.Row(DocxFixtures.R("{{item.Desc}}"))));
        var records = new[]
        {
            new Dictionary<string, string> { ["Desc"] = "Widget" },
            new Dictionary<string, string> { ["Desc"] = "Gadget" },
        };

        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, docx);

        await DocxEditor.FillRowsAsync(input.Path, output.Path, "item", records);

        var text = DocxEditor.ExtractText(await File.ReadAllBytesAsync(output.Path));
        Assert.Contains("Widget", text);
        Assert.Contains("Gadget", text);
        Assert.DoesNotContain("{{item.", text);
    }

    [Fact]
    public async Task ReplaceImageAsync_FromFileToFile_ReplacesThePlaceholder()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("Logo: {{logo}}")));

        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, docx);

        await DocxEditor.ReplaceImageAsync(
            input.Path, output.Path, "{{logo}}", ImageFixtures.Png(), widthPoints: 96);

        Assert.DoesNotContain(
            "{{logo}}", DocxEditor.ExtractText(await File.ReadAllBytesAsync(output.Path)));
    }

    [Fact]
    public async Task FilePathOverloads_RejectBlankPathsBeforeTouchingTheDisk()
    {
        var replacements = new Dictionary<string, string>();

        await Assert.ThrowsAsync<ArgumentNullException>(
            () => DocxEditor.ReplaceTextAsync(null!, "out.docx", replacements));
        await Assert.ThrowsAsync<ArgumentException>(
            () => DocxEditor.ReplaceTextAsync(" ", "out.docx", replacements));
        await Assert.ThrowsAsync<ArgumentException>(
            () => DocxEditor.ReplaceTextAsync("in.docx", " ", replacements));
    }

    [Fact]
    public async Task FilePathOverloads_LetFileNotFoundThrough_RatherThanWrappingIt()
    {
        // A wrong path and a broken document are different problems. Wrapping this in
        // DocumentConversionException would make the caller unwrap it to tell them apart.
        var missing = Path.Join(Path.GetTempPath(), $"doctoolkit-missing-{Guid.NewGuid():N}.docx");

        await Assert.ThrowsAsync<FileNotFoundException>(
            () => DocxEditor.ExtractTextAsync(missing));
    }

    [Fact]
    public void ReplaceText_PreservesAnEmbeddedObject()
    {
        // A real OLE package embed, authored through OfficeIMO's own API - not hand-built markup.
        // AddEmbeddedObject needs real files on disk for both the embedded content and the icon.
        using var embeddedFile = new TempFile();
        using var iconFile = new TempFile();
        File.WriteAllBytes(embeddedFile.Path,
            WorkbookEditor.Create("Data", new object[][] { new object[] { "X" } }));
        File.WriteAllBytes(iconFile.Path, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));

        byte[] docx;
        using (var word = OfficeIMO.Word.WordDocument.Create())
        {
            word.AddParagraph("Hello {{name}}, see the attachment below.");
            word.AddEmbeddedObject(embeddedFile.Path, iconFile.Path, null, null);
            using var ms = new MemoryStream();
            word.Save(ms);
            docx = ms.ToArray();
        }

        var edited = DocxEditor.ReplaceText(docx, Map("{{name}}", "Alice"));

        // GetEmbeddedPayloads is unchanged in kind/content-type/length - not merely "still
        // present", since a defect that replaced the payload with something else, or duplicated
        // it, would still pass a bare Assert.Single.
        using var before = new MemoryStream(docx, writable: false);
        using var beforeDoc = OfficeIMO.Word.WordDocument.Load(before);
        var beforePayload = Assert.Single(beforeDoc.GetEmbeddedPayloads(false));

        using var after = new MemoryStream(edited, writable: false);
        using var afterDoc = OfficeIMO.Word.WordDocument.Load(after);
        var afterPayload = Assert.Single(afterDoc.GetEmbeddedPayloads(false));

        Assert.Equal(beforePayload.Kind, afterPayload.Kind);
        Assert.Equal(beforePayload.ContentType, afterPayload.ContentType);
        Assert.Equal(beforePayload.Length, afterPayload.Length);
    }

    [Fact]
    public void InspectSignatures_ReportsAnUnsignedDocumentCleanly()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("unsigned")));

        var info = DocxEditor.InspectSignatures(docx);

        Assert.False(info.HasSignatures);
        Assert.Equal(0, info.SignatureCount);
    }

    [Fact]
    public async Task InspectSignaturesAsync_MatchesTheByteArrayOverload()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("unsigned")));

        var expected = DocxEditor.InspectSignatures(docx);
        using var source = new MemoryStream(docx, writable: false);
        var actual = await DocxEditor.InspectSignaturesAsync(source);

        Assert.Equal(expected.HasSignatures, actual.HasSignatures);
        Assert.Equal(expected.SignatureCount, actual.SignatureCount);
    }

    [Fact]
    public void ValidateSignatures_ReportsAnUnsignedDocumentCleanly()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("unsigned")));

        var report = DocxEditor.ValidateSignatures(docx);

        Assert.False(report.HasSignatures);
        Assert.False(report.IsCryptographicallyValid);
        Assert.Empty(report.Signatures);
    }

    [Fact]
    public async Task ValidateSignaturesAsync_MatchesTheByteArrayOverload()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("unsigned")));

        var expected = DocxEditor.ValidateSignatures(docx);
        using var source = new MemoryStream(docx, writable: false);
        var actual = await DocxEditor.ValidateSignaturesAsync(source);

        Assert.Equal(expected.HasSignatures, actual.HasSignatures);
        Assert.Equal(expected.IsCryptographicallyValid, actual.IsCryptographicallyValid);
    }

    [Fact]
    public void InspectSignatures_NullDocx_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => DocxEditor.InspectSignatures(null!));
    }

    [Fact]
    public void InspectSignatures_EmptyDocx_Throws()
    {
        var ex = Assert.Throws<ArgumentException>(() => DocxEditor.InspectSignatures(Array.Empty<byte>()));
        Assert.Equal("docx", ex.ParamName);
    }
}
