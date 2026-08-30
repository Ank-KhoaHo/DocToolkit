using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocToolkit.Tests;

public class DocxEditorCreateTests
{
    /// <summary>
    /// Asserts the package is schema-valid, not merely readable. Extracted text tells you what a
    /// document says, never whether Word will open it.
    /// </summary>
    private static void AssertValid(byte[] docx)
    {
        var errors = DocxFixtures.Validate(docx);
        Assert.True(errors.Count == 0,
            "expected a schema-valid package, got:\n" +
            string.Join("\n", errors.Take(3).Select(e => "  " + e.Description)));
    }

    [Fact]
    public void Create_WithNoBlocks_ProducesAValidEmptyDocument()
    {
        // Deliberately valid: this is the "blank DOCX to then edit" case, and a body with no
        // content is schema-valid.
        var docx = DocxEditor.Create(Array.Empty<DocxBlock>());

        AssertValid(docx);
        Assert.Equal("", DocxEditor.ExtractText(docx).Trim());
    }

    [Fact]
    public void Create_WritesParagraphsInOrder()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Paragraph("First."),
            DocxBlock.Paragraph("Second."),
        });

        AssertValid(docx);
        var text = DocxEditor.ExtractText(docx);
        Assert.Contains("First.", text);
        Assert.Contains("Second.", text);
        Assert.True(text.IndexOf("First.", StringComparison.Ordinal)
                  < text.IndexOf("Second.", StringComparison.Ordinal),
            "blocks must appear in the order they were given");
    }

    [Fact]
    public void Create_RejectsNullBlocks()
        => Assert.Throws<ArgumentNullException>(() => DocxEditor.Create(null!));

    [Fact]
    public void Create_RejectsANullBlockInTheSequence()
        => Assert.Throws<ArgumentException>(
            () => DocxEditor.Create(new DocxBlock?[] { DocxBlock.Paragraph("ok"), null }!));

    /// <summary>
    /// The output must be a real document, not merely one that opens: the existing editing API has
    /// to work against it.
    /// </summary>
    [Fact]
    public void Create_ProducesADocumentTheEditingApiCanEdit()
    {
        var docx = DocxEditor.Create(new[] { DocxBlock.Paragraph("Dear {{customer}}, hello.") });

        var edited = DocxEditor.ReplaceText(docx, new Dictionary<string, string>
        {
            ["{{customer}}"] = "Acme",
        });

        AssertValid(edited);
        Assert.Contains("Dear Acme, hello.", DocxEditor.ExtractText(edited));
    }

    [Fact]
    public void Create_WritesHeadingsAsRealWordHeadingStyles()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Heading("Title", 1),
            DocxBlock.Heading("Section", 2),
            DocxBlock.Paragraph("Body."),
        });

        AssertValid(docx);

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);

        var styleIds = doc.MainDocumentPart!.Document!.Body!
            .Descendants<ParagraphStyleId>().Select(s => s.Val?.Value).ToList();
        Assert.Contains("Heading1", styleIds);
        Assert.Contains("Heading2", styleIds);

        // The reference alone is not enough. A document that references Heading1 without DEFINING
        // it renders as plain text - no navigation pane entry, no table of contents - and no schema
        // check would ever flag it. This is the assertion that makes the feature real.
        var definitions = doc.MainDocumentPart.StyleDefinitionsPart;
        Assert.NotNull(definitions);
        var defined = definitions!.Styles!.Descendants<Style>()
            .Select(s => s.StyleId?.Value).ToList();
        Assert.Contains("Heading1", defined);
        Assert.Contains("Heading2", defined);

        // Normal is asserted because every heading style is basedOn it, and a DANGLING basedOn
        // degrades exactly as silently as a missing style: measured, a heading style whose
        // basedOn names an undefined Normal produces ZERO OpenXmlValidator errors. Without this
        // line, dropping the Normal definition while leaving basedOn in place passes every test.
        Assert.Contains("Normal", defined);

        // outlineLvl is what actually drives the navigation pane and TOC depth.
        var heading1 = definitions.Styles.Descendants<Style>()
            .Single(s => s.StyleId?.Value == "Heading1");
        Assert.Equal(0, heading1.StyleParagraphProperties!.OutlineLevel!.Val!.Value);

        // w:pStyle references a PARAGRAPH style. A style with the right id and outline level but
        // w:type="character" satisfies every other assertion here and validates clean - measured,
        // zero OpenXmlValidator errors - while being the wrong kind of style entirely.
        //
        // These four assertions - the style is defined, basedOn resolves, outlineLvl is right, and
        // the type is Paragraph - are deliberately the whole set. They are what make a heading
        // FUNCTION as one. Bold and font size are cosmetic: getting them wrong produces an ugly
        // heading, not a non-heading, so they are not pinned here and a future restyle should not
        // have to edit this test.
        Assert.Equal(StyleValues.Paragraph, heading1.Type!.Value);
    }

    [Fact]
    public void Create_DoesNotDefineHeadingStylesItDoesNotUse()
    {
        var docx = DocxEditor.Create(new[] { DocxBlock.Heading("Only one", 1) });

        AssertValid(docx);

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var defined = doc.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Descendants<Style>()
            .Select(s => s.StyleId?.Value).ToList();

        Assert.Contains("Heading1", defined);
        Assert.DoesNotContain("Heading4", defined);
    }

    [Fact]
    public void Create_WritesATableWithAGridAndAHeaderRow()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Table(
                new[] { "Region", "Total" },
                new[] { new object?[] { "North", 1200 }, new object?[] { "South", 980 } }),
        });

        AssertValid(docx);

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var table = Assert.Single(doc.MainDocumentPart!.Document!.Body!.Elements<Table>());

        // A w:tbl carrying no w:tblGrid makes the validator reject the first w:tr as an unexpected
        // child. An earlier fixture in this repo built exactly that and every text assertion passed.
        var grid = Assert.Single(table.Elements<TableGrid>());
        Assert.Equal(2, grid.Elements<GridColumn>().Count());

        Assert.Equal(3, table.Elements<TableRow>().Count()); // header + two data rows

        var text = DocxEditor.ExtractText(docx);
        Assert.Contains("Region", text);
        Assert.Contains("North", text);
        Assert.Contains("1200", text);
    }

    /// <summary>
    /// The same call must produce identical CONTENT regardless of the machine's regional settings.
    /// Deliberately not identical bytes: an OOXML package is a ZIP whose entries carry timestamps,
    /// so a byte comparison would be a flaky test wearing the costume of a strict one.
    /// Under a comma-decimal culture an uninvariant format writes "1234,5" instead of "1234.5".
    /// </summary>
    [Fact]
    public void Create_FormatsCellValuesInvariantly()
    {
        var blocks = new[]
        {
            DocxBlock.Table(new[] { "Value" }, new[] { new object?[] { 1234.5d } }),
        };

        var original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("de-DE");
            var text = DocxEditor.ExtractText(DocxEditor.Create(blocks));
            Assert.Contains("1234.5", text);
            Assert.DoesNotContain("1234,5", text);
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    /// <summary>
    /// A created table must be a real template the repeating-row API can expand, not merely a table
    /// that renders. FillRows walks each table's DIRECT child rows and each row's own cells' direct
    /// child paragraphs, so a table built with an unexpected nesting shape would be found by nothing
    /// and silently expand zero rows.
    /// </summary>
    [Fact]
    public void Create_ProducesATableFillRowsCanExpand()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Table(new[] { "Description" }, new[] { new object?[] { "{{item.Desc}}" } }),
        });

        var filled = DocxEditor.FillRows(docx, "item", new[]
        {
            new Dictionary<string, string>(StringComparer.Ordinal) { ["Desc"] = "Widget" },
            new Dictionary<string, string>(StringComparer.Ordinal) { ["Desc"] = "Gadget" },
        });

        AssertValid(filled);
        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("Widget", text);
        Assert.Contains("Gadget", text);
        Assert.DoesNotContain("{{item.Desc}}", text);
    }

    /// <summary>
    /// Every row must carry exactly as many cells as the grid has columns. A row with a different
    /// count renders as a malformed table in Word while staying schema-valid, so BuildRow's padding
    /// is load-bearing rather than cosmetic — and it was exercised by no other test here. (The
    /// over-long direction is rejected at the factory now; see Table_RejectsARowLongerThanItsHeader.)
    /// </summary>
    [Fact]
    public void Create_PadsARowShorterThanItsHeader()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Table(new[] { "A", "B", "C" }, new[] { new object?[] { "only one" } }),
        });

        AssertValid(docx);

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var table = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single();

        // NotEmpty first: Assert.All passes over an empty sequence, and CLAUDE.md records that a
        // w:tbl with NO w:tr is schema-valid - so AssertValid would not catch a rowless table
        // either, and this test would go green on a Create that emitted no rows at all.
        var createdRows = table.Elements<TableRow>().ToList();
        Assert.Equal(2, createdRows.Count);          // the header, and the one short row

        Assert.All(createdRows, r => Assert.Equal(3, r.Elements<TableCell>().Count()));
    }

    /// <summary>
    /// The other half, and it is NOT symmetric with padding. Padding a short row discards nothing;
    /// dropping cells from a long row destroys caller data with no signal — no exception, no return
    /// value recording it, and a document that looks entirely correct. This used to truncate
    /// silently. It now throws, for the same reason <see cref="DocxBlock.Heading"/> refuses to clamp
    /// an out-of-range level: the loss is only ever noticed in the finished document.
    ///
    /// The throw comes from the FACTORY, not from writing, so it lands on the line that built the
    /// bad row rather than surfacing later out of a Create call assembling many blocks.
    /// </summary>
    /// <remarks>
    /// The message is asserted WHOLE, not by substring. <c>Assert.Contains("1 column", …)</c> would
    /// pass against "1 columns" too, so it could not tell a working pluraliser from a broken one —
    /// the same vacuity that made two earlier assertions in this file useless. Both arities are
    /// covered for that reason: the singular case is the only one that can catch it.
    /// </remarks>
    [Theory]
    [InlineData(1, 2, "Row 1 has 2 cells but the table has 1 column.")]
    [InlineData(2, 3, "Row 1 has 3 cells but the table has 2 columns.")]
    public void Table_RejectsARowLongerThanItsHeader(int headers, int cells, string expected)
    {
        var ex = Assert.Throws<ArgumentException>(() => DocxBlock.Table(
            Enumerable.Range(0, headers).Select(i => $"H{i}"),
            new[] { Enumerable.Range(0, cells).Select(i => (object?)$"c{i}") }));

        Assert.Equal("rows", ex.ParamName);
        Assert.StartsWith(expected, ex.Message);
    }

    /// <summary>
    /// The boundary either side of the throw: a row of exactly the header's width is accepted, and
    /// one cell more is not. The real assertion here is "does not throw" — <c>Assert.NotNull</c> is
    /// trivially true once construction succeeds. Moving the comparison to <c>&gt;=</c> fails this
    /// and six other tests.
    /// </summary>
    [Fact]
    public void Table_AcceptsARowExactlyAsWideAsItsHeader()
    {
        var block = DocxBlock.Table(new[] { "A", "B" }, new[] { new object?[] { "x", "y" } });

        Assert.NotNull(block);
    }

    /// <summary>
    /// The writer's own guard, which the factory makes unreachable through the public API but which
    /// is NOT dead code: <c>TableBlock</c> is internal, so internal code can construct an over-wide
    /// row without passing the factory, and <c>OpenXmlValidator</c> was measured blind to a
    /// row/grid count mismatch in both directions. Reached here through
    /// <c>InternalsVisibleTo</c> — the same route <c>TableRowFinderTests</c> uses to pin an
    /// invariant whose end-to-end symptom is a plausible-looking document rather than an error.
    /// Without this test it is the one new invariant in this feature with nothing behind it.
    /// </summary>
    [Fact]
    public void Create_ThrowsWhenAnInternallyBuiltRowExceedsTheGrid()
    {
        var block = new TableBlock(
            new[] { "A" },
            new IReadOnlyList<object?>[] { new object?[] { "kept", "SURPLUS" } });

        var ex = Assert.Throws<DocumentConversionException>(() => DocxEditor.Create(new DocxBlock[] { block }));

        var inner = Assert.IsType<InvalidOperationException>(ex.InnerException);
        Assert.Equal(
            "A table row has more than 1 cell. DocxBlock.Table rejects this, so a TableBlock was "
            + "built internally without going through it.",
            inner.Message);
    }

    /// <summary>
    /// Pins the rendered form of every type <c>Format</c> handles explicitly. All five were
    /// completely untested, and their output deliberately differs from what the same value renders
    /// as through <see cref="WorkbookEditor"/> — measured against the pinned ClosedXML: a DateTime
    /// reads there in a slashed, culture-dependent form (<c>08/06/2026 13:05:30</c> or
    /// <c>06/08/2026 01:05:30 PM</c>, depending on the machine) and <c>2026-08-06 13:05:30</c> here,
    /// and a one-day TimeSpan reads <c>26:03:04</c> there and <c>1.02:03:04</c> here. The DOCX side
    /// is invariant, which is why only it is asserted below. Those differences
    /// are intentional (see the note on <see cref="DocxBlock.Table"/>), but without this test they
    /// are invisible and free to drift in either direction.
    /// </summary>
    [Fact]
    public void Create_FormatsEveryExplicitlyHandledCellType()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Table(new[] { "Value" }, new[]
            {
                new object?[] { true },
                new object?[] { false },
                // Deliberately DISTINCT values per row. Using the same date for DateTime and
                // DateOnly would make Assert.Contains("2026-08-06") pass on the DateTime row's
                // "2026-08-06 13:05:30" alone, proving nothing about DateOnly — the assertion
                // would look thorough and be vacuous. Same trap for TimeOnly against the
                // DateTime row's time portion.
                new object?[] { new DateTime(2026, 8, 6, 13, 5, 30, DateTimeKind.Unspecified) },
                new object?[] { new DateOnly(2026, 1, 2) },
                new object?[] { new TimeOnly(9, 7, 5) },
                new object?[] { new TimeSpan(1, 2, 3, 4) },
            }),
        });

        AssertValid(docx);
        var text = DocxEditor.ExtractText(docx);

        Assert.Contains("TRUE", text);
        Assert.Contains("FALSE", text);
        Assert.Contains("2026-08-06 13:05:30", text);

        // A DateOnly must render as a bare date. "2026-01-02 00:00:00" is what a naive
        // DateOnly-to-DateTime conversion would produce, so its absence is the real assertion.
        Assert.Contains("2026-01-02", text);
        Assert.DoesNotContain("2026-01-02 00:00:00", text);

        Assert.Contains("09:07:05", text);
        Assert.Contains("1.02:03:04", text);
    }

    /// <summary>
    /// Asserted on the CELL, not on extracted text. The extracted-text form of this test was
    /// vacuous: <c>AssertValid</c> plus <c>Contains("x")</c> hold whatever a null cell renders as,
    /// so the behaviour in the test's own name went unverified — mutating <c>null =&gt; string.Empty</c>
    /// to a marker string passed all 404 core tests. An empty cell is invisible to
    /// <c>ExtractText</c> by definition, so proving it is empty requires reading the cell itself.
    /// </summary>
    [Fact]
    public void Create_WritesNullCellsAsEmpty()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Table(new[] { "A", "B" }, new[] { new object?[] { null, "x" } }),
        });

        AssertValid(docx);

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var cells = doc.MainDocumentPart!.Document!.Body!.Elements<Table>().Single()
            .Elements<TableRow>().Last().Elements<TableCell>().ToList();

        Assert.Equal(2, cells.Count);
        Assert.Equal(string.Empty, cells[0].InnerText);
        Assert.Equal("x", cells[1].InnerText);
    }

    /// <summary>
    /// The Stream overload must produce the same document as the byte[] one — same content, not
    /// same bytes. An OOXML package is a ZIP whose entries carry timestamps, so byte-identity would
    /// be a flaky test wearing the costume of a strict one.
    ///
    /// StreamOverloadTests covers the SHAPE of this overload (forward-only destinations, not
    /// disposing the caller's stream, sync-vs-async call counts) across every overload at once.
    /// What it does not check is that this particular one writes the right document, which is what
    /// this adds.
    /// </summary>
    [Fact]
    public async Task CreateAsync_MatchesTheByteArrayOverload()
    {
        var blocks = new[]
        {
            DocxBlock.Heading("Quarterly Report", 1),
            DocxBlock.Paragraph("Revenue was up 12%."),
            DocxBlock.Table(new[] { "Region" }, new[] { new object?[] { "North" } }),
        };

        using var destination = new MemoryStream();
        await DocxEditor.CreateAsync(blocks, destination);

        var written = destination.ToArray();
        AssertValid(written);
        Assert.Equal(DocxEditor.ExtractText(DocxEditor.Create(blocks)), DocxEditor.ExtractText(written));
    }

    /// <summary>
    /// An ImagePart must belong to the part that owns the paragraph. Get the scope wrong and the
    /// relationship id resolves against the wrong container: Word opens the file and simply shows
    /// nothing where the image should be — no error, no exception, and every text assertion still
    /// passes. So this asserts the part is there AND that the drawing's <c>r:embed</c> actually
    /// resolves back to it; counting parts alone would pass against a drawing pointing nowhere.
    /// </summary>
    /// <remarks>
    /// BOTH formats, deliberately. With PNG alone, <c>Assert.Equal("image/png", …)</c> distinguishes
    /// only a wrong constant from a right one — hardcoding <c>AddNewPart&lt;ImagePart&gt;("image/png")</c>
    /// passed all 408 tests. Two formats is what makes it test DERIVATION from the magic bytes,
    /// which is the property the doc comments claim. The sibling
    /// <c>ReplaceImage_DeclaresTheContentTypeTheBytesActuallyAre</c> already covers both this way.
    /// </remarks>
    [Theory]
    [InlineData("png", "image/png")]
    [InlineData("jpeg", "image/jpeg")]
    public void Create_EmbedsAnImageOwnedByTheMainDocumentPart(string format, string expectedContentType)
    {
        var bytes = format == "png" ? ImageFixtures.Png() : ImageFixtures.Jpeg();

        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Paragraph("Logo below."),
            DocxBlock.Image(bytes),
        });

        AssertValid(docx);

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;

        var part = Assert.Single(main.ImageParts);
        Assert.Equal(expectedContentType, part.ContentType);
        Assert.Single(main.Document!.Body!.Descendants<Drawing>());

        var embedId = main.Document.Body.Descendants<A.Blip>().Single().Embed!.Value!;
        Assert.Same(part, main.GetPartById(embedId));

        // The part must actually HOLD the caller's bytes. Deleting the FeedData call leaves a
        // schema-valid package with an empty part — resolving r:embed, correct extents, unique ids —
        // that Word renders as a blank frame. Measured: that deletion passed all 408 tests, because
        // nothing else in this repo reads an ImagePart's stream back.
        using var stored = part.GetStream();
        using var buffer = new MemoryStream();
        stored.CopyTo(buffer);
        Assert.Equal(bytes, buffer.ToArray());
    }

    /// <summary>
    /// A duplicate <c>wp:docPr/@id</c> makes Word declare the file corrupt and offer to repair it —
    /// invisible to any test that only reads text back, and invisible to the validator too.
    /// </summary>
    [Fact]
    public void Create_GivesEveryDrawingAUniqueId()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Image(ImageFixtures.Png()),
            DocxBlock.Image(ImageFixtures.Png()),
            DocxBlock.Image(ImageFixtures.Png()),
        });

        AssertValid(docx);

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var ids = doc.MainDocumentPart!.Document!.Body!
            .Descendants<DW.DocProperties>().Select(p => p.Id!.Value).ToList();

        Assert.Equal(3, ids.Count);
        Assert.Equal(3, ids.Distinct().Count());

        // Three separate parts, not one part referenced three times: each block carried its own
        // bytes, and collapsing them would be a different behaviour that happens to share this id
        // assertion.
        Assert.Equal(3, doc.MainDocumentPart.ImageParts.Count());
    }

    /// <summary>
    /// Alt text reaches <c>wp:docPr/@descr</c>, which is what a screen reader announces.
    ///
    /// <c>@name</c> is a different thing — it labels the object in Word's selection pane — and stays
    /// the generated "Image N". Setting both from one value is what the bug was.
    /// </summary>
    [Fact]
    public void Create_PutsAltTextOnTheDrawing()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Image(ImageFixtures.Png(), altText: "Company logo, blue on white"),
        });

        AssertValid(docx);

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var props = doc.MainDocumentPart!.Document!.Body!.Descendants<DW.DocProperties>().Single();

        Assert.Equal("Company logo, blue on white", props.Description!.Value);
        Assert.Equal("Image 1", props.Name!.Value);
    }

    /// <summary>
    /// With no alt text the attribute is OMITTED, not filled with the generated name. Emitting
    /// <c>descr="Image 1"</c> is worse than emitting nothing: a screen reader reads it out as though
    /// it described the picture, so the user is told "Image 1" instead of being told there is an
    /// undescribed image. That was the behaviour before this test existed.
    /// </summary>
    [Fact]
    public void Create_OmitsAltTextWhenNoneIsGiven()
    {
        var docx = DocxEditor.Create(new[] { DocxBlock.Image(ImageFixtures.Png()) });

        AssertValid(docx);

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var props = doc.MainDocumentPart!.Document!.Body!.Descendants<DW.DocProperties>().Single();

        Assert.Null(props.Description);
        Assert.Equal("Image 1", props.Name!.Value);
    }

    /// <summary>
    /// The one place in the package where two INDEPENDENT drawing-id schemes meet: Create issues ids
    /// from a monotonic counter starting at 1, because it builds an empty document; ReplaceImage
    /// scans every part for the highest existing id, because it edits a template it did not write.
    /// Feed one to the other and a collision would produce duplicate <c>wp:docPr/@id</c> values —
    /// which makes Word declare the file corrupt and offer to repair it, while no schema check
    /// flags it and every text assertion still passes.
    ///
    /// Neither scheme is wrong; nothing tested that they compose. The whole-branch review had to run
    /// this by hand to be sure of it, which is the definition of a test that should exist.
    /// </summary>
    [Fact]
    public void Create_ThenReplaceImage_KeepsEveryDrawingIdUnique()
    {
        var created = DocxEditor.Create(new[]
        {
            DocxBlock.Image(ImageFixtures.Png()),
            DocxBlock.Paragraph("Signature: {{sig}}"),
        });

        var filled = DocxEditor.ReplaceImage(created, "{{sig}}", ImageFixtures.Jpeg());

        AssertValid(filled);

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var ids = doc.MainDocumentPart!.Document!.Body!
            .Descendants<DW.DocProperties>().Select(p => p.Id!.Value).ToList();

        Assert.Equal(2, ids.Count);
        Assert.Equal(2, ids.Distinct().Count());

        // Both images survived as separate parts, so the ids belong to two real drawings.
        Assert.Equal(2, doc.MainDocumentPart.ImageParts.Count());
    }

    /// <summary>
    /// The missing link between "Resolve returns a number" and "the document is valid". Every other
    /// size test asserts on Resolve in isolation, so the bound could be — and for one commit was —
    /// four orders of magnitude too loose while every unit test passed. This runs the largest
    /// ACCEPTED size through the real writer and the real validator.
    ///
    /// The rejected side is asserted from the public factory too, because the exception type differs
    /// by entry point unless the factory enforces it: Resolve throws inside
    /// DocxDocumentWriter.Write's try, which would wrap it as DocumentConversionException.
    /// </summary>
    [Fact]
    public void Create_AcceptsTheLargestValidImageSizeAndRejectsAnythingLarger()
    {
        // 2 x 3 px: the height derives at 1.5x, so the WIDTH that fits is bounded by the height.
        var docx = DocxEditor.Create(new[] { DocxBlock.Image(ImageFixtures.Png(), widthPoints: 112_000) });

        AssertValid(docx);

        var ex = Assert.Throws<ArgumentOutOfRangeException>(
            () => DocxBlock.Image(ImageFixtures.Png(), widthPoints: 200_000));
        Assert.Equal("widthPoints", ex.ParamName);
    }

    /// <summary>
    /// Given one dimension, the other preserves the aspect ratio. The height is asserted EXACTLY,
    /// not merely as positive: the fixture is 2x3 px, so 72pt wide must give 108pt tall. A
    /// <c>&gt; 0</c> assertion would pass against any scaling bug that still produced a number,
    /// which is the whole failure mode here — nothing downstream validates these values, so a
    /// factor-of-ten slip yields a perfectly valid document containing a wrongly sized image.
    /// </summary>
    [Fact]
    public void Create_ScalesAnImageProportionallyFromASingleDimension()
    {
        var natural = DocxEditor.Create(new[] { DocxBlock.Image(ImageFixtures.Png()) });
        var scaled = DocxEditor.Create(new[] { DocxBlock.Image(ImageFixtures.Png(), widthPoints: 72) });

        // Both sides asserted by VALUE. An inequality between them would be satisfied by any pair of
        // different numbers, so it could not tell correct scaling from a scaling bug — and it left
        // the intrinsic path unpinned end-to-end: hardcoding a width default passed all 408 tests.
        // Intrinsic is the pixel size at 96 DPI, 9525 EMU per pixel.
        Assert.Equal((2L * 9525, 3L * 9525), ExtentOf(natural));

        // 1 point = 12,700 EMU. 72pt wide, and 2:3 gives 108pt tall.
        Assert.Equal((72L * 12700, 108L * 12700), ExtentOf(scaled));

        static (long Width, long Height) ExtentOf(byte[] docx)
        {
            using var ms = new MemoryStream(docx);
            using var doc = WordprocessingDocument.Open(ms, false);
            var extent = doc.MainDocumentPart!.Document!.Body!.Descendants<DW.Extent>().First();
            return (extent.Cx!.Value, extent.Cy!.Value);
        }
    }

    /// <summary>
    /// The other half of the "only what is used" rule: a document with no headings gets no styles
    /// part at all. Asserted rather than assumed because dropping the early return in
    /// AddHeadingStyles would give every document an unused styles part, and nothing else in this
    /// file would notice — the documents would stay schema-valid and their text unchanged.
    /// </summary>
    [Fact]
    public void Create_WithNoHeadings_AddsNoStylesPart()
    {
        var docx = DocxEditor.Create(new[] { DocxBlock.Paragraph("Body only.") });

        AssertValid(docx);

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Null(doc.MainDocumentPart!.StyleDefinitionsPart);
    }
}
