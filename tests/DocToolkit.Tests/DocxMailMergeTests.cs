using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// Covers <see cref="DocxMailMerge"/> — filling a Word mail-merge template from named values.
///
/// <b>Every behaviour asserted here was measured against OfficeIMO before it was designed around</b>,
/// which is why several of these tests look like they are pinning someone else's library rather than
/// ours. They are: each one is a property this API's shape depends on, and an upstream change that
/// removed it would otherwise surface as a caller's letters quietly coming out wrong.
/// </summary>
public class DocxMailMergeTests
{
    // ---- both on-disk encodings ---------------------------------------------------------------

    [Fact]
    public void Merge_FillsASimpleField()
    {
        byte[] merged = DocxMailMerge.Merge(Simple("FirstName"), new Dictionary<string, string> { ["FirstName"] = "Khoa" });

        Assert.Equal("Khoa|", Text(merged));
    }

    [Fact]
    public void Merge_FillsAComplexField_TheFormWordItselfWrites()
    {
        // Word writes fldChar begin / instrText / separate / result / end. Most generators and
        // hand-built documents write w:fldSimple instead. An engine handling only one would leave
        // the other in the output, and a test using only one encoding would never notice.
        byte[] merged = DocxMailMerge.Merge(Complex("FirstName"), new Dictionary<string, string> { ["FirstName"] = "Khoa" });

        Assert.Equal("Dear Khoa, welcome.", Text(merged));
    }

    // ---- the unfilled field, which is why Merge is strict --------------------------------------

    [Fact]
    public void Merge_RefusesToProduceADocumentWithAnUnfilledField()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxMailMerge.Merge(Simple("FirstName", "Balance"),
                new Dictionary<string, string> { ["FirstName"] = "Khoa" }));

        // The message must name the field, or the caller has to go and find it themselves.
        Assert.Contains("Balance", ex.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("FirstName", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Merge_NamesEveryUnfilledField_NotOnlyTheFirst()
    {
        // The engine's own EnsureComplete throws on the first one it meets. Reading IsComplete
        // instead is what lets this list all of them, and this is the test that pins the difference.
        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxMailMerge.Merge(Simple("A", "B", "C"), new Dictionary<string, string>()));

        Assert.Contains("A", ex.Message, StringComparison.Ordinal);
        Assert.Contains("B", ex.Message, StringComparison.Ordinal);
        Assert.Contains("C", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MergeWithReport_ProducesTheDocumentAnyway_AndTheBytesSTILLReadAsUnfinished()
    {
        // TWO different statements, and only the second is about what a caller would suffer:
        //   1. the report says MissingValue    - the library's claim
        //   2. the shipped bytes read «Balance» - the letter that reaches a customer
        // Asserting only the first would pass against an engine that blanked the field instead,
        // which is a different and equally silent outcome.
        DocxMailMergeResult result = DocxMailMerge.MergeWithReport(
            Simple("Balance"), new Dictionary<string, string>());

        Assert.False(result.Report.IsComplete);
        Assert.Equal("Balance", Assert.Single(result.Report.MissingFieldNames));
        Assert.Equal(DocxMailMergeFieldStatus.MissingValue, Assert.Single(result.Report.Fields).Status);

        Assert.Contains("«Balance»", Text(result.Document), StringComparison.Ordinal);
    }

    [Fact]
    public void MergeWithReport_ReportsACompleteMerge()
    {
        DocxMailMergeResult result = DocxMailMerge.MergeWithReport(
            Simple("FirstName"), new Dictionary<string, string> { ["FirstName"] = "Khoa" });

        Assert.True(result.Report.IsComplete);
        Assert.Empty(result.Report.MissingFieldNames);
        Assert.Equal(1, result.Report.MergedCount);
        Assert.Equal(DocxMailMergeFieldStatus.Merged, Assert.Single(result.Report.Fields).Status);
        Assert.Equal("Khoa|", Text(result.Document));
    }

    // ---- flattening, which no text assertion can see -------------------------------------------

    [Fact]
    public void Merge_FlattensTheMergedFields_SoWordCannotReMergeTheResult()
    {
        // Measured: the TEXT is identical whether the fields are flattened or left live, so a text
        // assertion here would pass vacuously. What differs is the markup, and what it costs is
        // real - a live field re-merges or shows field shading when the result is opened in Word.
        byte[] template = Simple("FirstName");
        byte[] merged = DocxMailMerge.Merge(template, new Dictionary<string, string> { ["FirstName"] = "Khoa" });

        Assert.Equal(1, FieldCount(template));
        Assert.Equal(0, FieldCount(merged));
    }

    // ---- template inspection --------------------------------------------------------------------

    [Fact]
    public void InspectTemplate_DiscoversTheFieldNamesWithoutBeingToldThem()
    {
        // The underlying call takes the expected names as a parameter and treats null and an EMPTY
        // sequence differently: null discovers, empty asserts "this template has no fields" and so
        // reports a sound template invalid with one issue per field it does have. Array.Empty is
        // the natural defensive choice and it is the wrong one - this API passes null so that no
        // caller can express the mistake, and this test is what stops a refactor undoing it.
        DocxMailMergeTemplate template = DocxMailMerge.InspectTemplate(Simple("FirstName", "Balance"));

        Assert.True(template.IsValid);
        Assert.Empty(template.Issues);
        Assert.Equal(["Balance", "FirstName"], template.FieldNames.OrderBy(n => n, StringComparer.Ordinal));
    }

    [Fact]
    public void InspectTemplate_ReportsAMalformedField()
    {
        DocxMailMergeTemplate template = DocxMailMerge.InspectTemplate(
            Build(body => body.Append(new Paragraph(Field(" MERGEFIELD ", "«»")))));

        Assert.False(template.IsValid);
        Assert.Equal(DocxMailMergeIssueKind.MalformedField, Assert.Single(template.Issues).Kind);
        Assert.False(string.IsNullOrWhiteSpace(Assert.Single(template.Issues).Message));
    }

    [Fact]
    public void InspectTemplate_IgnoresFieldsThatAreNotMergeFields()
    {
        // A DATE field is a field, and it is none of this API's business. Reporting it would make
        // every document containing a date look like a mail-merge template.
        DocxMailMergeTemplate template = DocxMailMerge.InspectTemplate(
            Build(body => body.Append(new Paragraph(Field(" DATE \\@ \"dd/MM/yyyy\" ", "01/01/2026")))));

        Assert.True(template.IsValid);
        Assert.Empty(template.FieldNames);
    }

    [Fact]
    public void InspectTemplate_OnADocumentWithNoMergeFields_IsValidAndEmpty()
    {
        DocxMailMergeTemplate template = DocxMailMerge.InspectTemplate(
            DocxEditor.Create([DocxBlock.Paragraph("an ordinary document")]));

        Assert.True(template.IsValid);
        Assert.Empty(template.FieldNames);
    }

    [Fact]
    public void Merge_OnADocumentWithNoMergeFields_SucceedsAndChangesNothing()
    {
        // This is the case InspectTemplate exists for: passing the wrong document is not an error
        // here, so nothing about the merge itself would tell a caller they had done it.
        byte[] plain = DocxEditor.Create([DocxBlock.Paragraph("an ordinary document")]);

        DocxMailMergeResult result = DocxMailMerge.MergeWithReport(
            plain, new Dictionary<string, string> { ["Unused"] = "x" });

        Assert.True(result.Report.IsComplete);
        Assert.Empty(result.Report.Fields);
        Assert.Equal(0, result.Report.MergedCount);
        Assert.Equal("an ordinary document", Text(result.Document));
    }

    // ---- matching rules ---------------------------------------------------------------------------

    [Theory]
    [InlineData("FirstName")]
    [InlineData("firstname")]
    [InlineData("FIRSTNAME")]
    public void Merge_MatchesFieldNamesCaseInsensitively(string key)
    {
        // The engine's own matching, not the dictionary's comparer - this passes an ordinary
        // ordinal Dictionary, so if the engine ever started relying on the comparer instead, two of
        // these three would fail rather than the behaviour silently changing under a caller.
        byte[] merged = DocxMailMerge.Merge(Simple("FirstName"), new Dictionary<string, string> { [key] = "Khoa" });

        Assert.Equal("Khoa|", Text(merged));
    }

    [Fact]
    public void Merge_IgnoresValuesForFieldsTheTemplateDoesNotHave()
    {
        byte[] merged = DocxMailMerge.Merge(Simple("FirstName"),
            new Dictionary<string, string> { ["FirstName"] = "Khoa", ["Nonexistent"] = "ignored" });

        Assert.Equal("Khoa|", Text(merged));
    }

    [Fact]
    public void Merge_FillsEveryOccurrenceOfAName_AndReportsOnePerOccurrence()
    {
        DocxMailMergeResult result = DocxMailMerge.MergeWithReport(
            Simple("FirstName", "FirstName"), new Dictionary<string, string> { ["FirstName"] = "Khoa" });

        Assert.Equal("Khoa|Khoa|", Text(result.Document));
        Assert.Equal(2, result.Report.MergedCount);
        Assert.Equal(2, result.Report.Fields.Count);
    }

    // ---- the null value, which strictness cannot catch ---------------------------------------------

    [Fact]
    public void Merge_RefusesANullValue()
    {
        // Measured: the engine merges null as an empty string and reports the document COMPLETE. So
        // a database NULL produces "Your balance is " with nothing flagging it - the one silent
        // half-merge that neither the strict overload nor the report can see. Refusing it is the
        // only place this can be caught.
        var ex = Assert.Throws<ArgumentException>(
            () => DocxMailMerge.Merge(Simple("Balance"), new Dictionary<string, string> { ["Balance"] = null! }));

        Assert.Equal("values", ex.ParamName);
        Assert.Contains("Balance", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Merge_AcceptsAnEmptyString_BecauseThatIsADecision()
    {
        // The other half of the rule above: string.Empty says "leave it blank" and is honoured.
        byte[] merged = DocxMailMerge.Merge(Simple("Balance"), new Dictionary<string, string> { ["Balance"] = "" });

        Assert.Equal("|", Text(merged));
    }

    // ---- guards ---------------------------------------------------------------------------------

    [Fact]
    public void EveryByteArrayOverload_RefusesNullAndEmptyByTheParameterItDeclares()
    {
        var values = new Dictionary<string, string>();

        Assert.Equal("docx", Assert.Throws<ArgumentNullException>(() => DocxMailMerge.InspectTemplate(null!)).ParamName);
        Assert.Equal("docx", Assert.Throws<ArgumentException>(() => DocxMailMerge.InspectTemplate([])).ParamName);
        Assert.Equal("docx", Assert.Throws<ArgumentNullException>(() => DocxMailMerge.Merge(null!, values)).ParamName);
        Assert.Equal("docx", Assert.Throws<ArgumentException>(() => DocxMailMerge.Merge([], values)).ParamName);
        Assert.Equal("docx", Assert.Throws<ArgumentNullException>(() => DocxMailMerge.MergeWithReport(null!, values)).ParamName);
        Assert.Equal("docx", Assert.Throws<ArgumentException>(() => DocxMailMerge.MergeWithReport([], values)).ParamName);

        byte[] template = Simple("FirstName");
        Assert.Equal("values", Assert.Throws<ArgumentNullException>(() => DocxMailMerge.Merge(template, null!)).ParamName);
        Assert.Equal("values", Assert.Throws<ArgumentNullException>(() => DocxMailMerge.MergeWithReport(template, null!)).ParamName);
    }

    [Fact]
    public void EveryOverload_WrapsAnUnreadableDocument()
    {
        byte[] rubbish = [1, 2, 3, 4];
        var values = new Dictionary<string, string>();

        Assert.NotNull(Assert.Throws<DocumentConversionException>(() => DocxMailMerge.InspectTemplate(rubbish)).InnerException);
        Assert.NotNull(Assert.Throws<DocumentConversionException>(() => DocxMailMerge.Merge(rubbish, values)).InnerException);
        Assert.NotNull(Assert.Throws<DocumentConversionException>(() => DocxMailMerge.MergeWithReport(rubbish, values)).InnerException);
    }

    // ---- the Stream forms ---------------------------------------------------------------------------

    [Fact]
    public async Task InspectTemplateAsync_ReadsTheSameTemplateAndLeavesTheStreamOpen()
    {
        using var source = new MemoryStream(Simple("FirstName"));

        DocxMailMergeTemplate template = await DocxMailMerge.InspectTemplateAsync(source);

        Assert.Equal("FirstName", Assert.Single(template.FieldNames));
        source.Position = 0;
        Assert.True(source.ReadByte() >= 0, "the caller's stream must not be closed");
    }

    [Fact]
    public async Task MergeAsync_WritesTheMergedDocumentAndRefusesAnUnfilledField()
    {
        using var source = new MemoryStream(Simple("FirstName"));
        using var destination = new MemoryStream();

        await DocxMailMerge.MergeAsync(source, destination,
            new Dictionary<string, string> { ["FirstName"] = "Khoa" });
        Assert.Equal("Khoa|", Text(destination.ToArray()));

        using var second = new MemoryStream(Simple("FirstName"));
        using var unwritten = new MemoryStream();
        await Assert.ThrowsAsync<DocumentConversionException>(
            () => DocxMailMerge.MergeAsync(second, unwritten, new Dictionary<string, string>()));

        // The refusal has to happen BEFORE anything reaches the caller's destination, or a caller
        // who catches the exception is left holding a half-merged document they did not ask for.
        Assert.Equal(0, unwritten.Length);
    }

    [Fact]
    public async Task MergeWithReportAsync_WritesTheDocumentAndReturnsOnlyTheReport()
    {
        using var source = new MemoryStream(Simple("Balance"));
        using var destination = new MemoryStream();

        DocxMailMergeReport report = await DocxMailMerge.MergeWithReportAsync(
            source, destination, new Dictionary<string, string>());

        Assert.False(report.IsComplete);
        Assert.Contains("«Balance»", Text(destination.ToArray()), StringComparison.Ordinal);
    }

    // ---- MergeBatch: strict, one document per record --------------------------------------------

    [Fact]
    public void MergeBatch_YieldsOneDocumentPerRecord_InOrder()
    {
        var records = new IReadOnlyDictionary<string, string>[]
        {
            new Dictionary<string, string> { ["FirstName"] = "Alice" },
            new Dictionary<string, string> { ["FirstName"] = "Bob" },
            new Dictionary<string, string> { ["FirstName"] = "Carol" },
        };

        var documents = DocxMailMerge.MergeBatch(Simple("FirstName"), records).ToList();

        Assert.Equal(3, documents.Count);
        Assert.Equal("Alice|", Text(documents[0]));
        Assert.Equal("Bob|", Text(documents[1]));
        Assert.Equal("Carol|", Text(documents[2]));
    }

    [Fact]
    public void MergeBatch_RefusesOnTheBadRecord_MidEnumeration_NamingItsIndex()
    {
        var records = new IReadOnlyDictionary<string, string>[]
        {
            new Dictionary<string, string> { ["FirstName"] = "Alice", ["Balance"] = "100" },
            new Dictionary<string, string> { ["FirstName"] = "Bob" }, // Balance missing on purpose
            new Dictionary<string, string> { ["FirstName"] = "Carol", ["Balance"] = "300" },
        };
        // Counts how many records were actually pulled from the sequence, which is what proves
        // record 2 never ran -- Assert.Single(seen) below is consistent with that but does not, by
        // itself, rule out record 2 having been merged and then discarded.
        var pulledCount = 0;
        IEnumerable<IReadOnlyDictionary<string, string>> countingRecords =
            records.Select(r => { pulledCount++; return r; });

        var seen = new List<byte[]>();
        var ex = Assert.Throws<DocumentConversionException>(() =>
        {
            foreach (var document in DocxMailMerge.MergeBatch(Simple("FirstName", "Balance"), countingRecords))
                seen.Add(document);
        });

        // Record 0 was already produced and handed to the caller before the throw -- a strict
        // batch fails ON the bad record, not before it.
        Assert.Single(seen);
        // Record 2's merge never ran at all: only records 0 and 1 (the good one and the bad one)
        // were ever pulled from the sequence.
        Assert.Equal(2, pulledCount);
        // "1" alone would also match the unrelated "1 merge field(s)" substring MergeCore's own
        // message already contains -- "Record 1:" is what actually pins the index.
        Assert.Contains("Record 1:", ex.Message, StringComparison.Ordinal);
        Assert.Contains("Balance", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MergeBatch_MatchesFieldNamesCaseInsensitively()
    {
        var records = new IReadOnlyDictionary<string, string>[]
        {
            new Dictionary<string, string> { ["firstname"] = "Alice" },
        };

        var documents = DocxMailMerge.MergeBatch(Simple("FirstName"), records).ToList();

        Assert.Equal("Alice|", Text(Assert.Single(documents)));
    }

    [Fact]
    public void MergeBatch_IgnoresValuesForFieldsTheTemplateDoesNotHave()
    {
        var records = new IReadOnlyDictionary<string, string>[]
        {
            new Dictionary<string, string> { ["FirstName"] = "Alice", ["Nonexistent"] = "ignored" },
        };

        var documents = DocxMailMerge.MergeBatch(Simple("FirstName"), records).ToList();

        Assert.Equal("Alice|", Text(Assert.Single(documents)));
    }

    [Fact]
    public void MergeBatch_OnAnEmptyRecordSequence_ProducesNoDocuments()
    {
        var documents = DocxMailMerge.MergeBatch(
            Simple("FirstName"), Array.Empty<IReadOnlyDictionary<string, string>>());

        Assert.Empty(documents);
    }

    [Fact]
    public void MergeBatch_ValidatesArgumentsImmediately_NotOnlyWhenEnumerated()
    {
        // MergeBatch must NOT be an iterator method itself (no `yield return` in its own body) --
        // if it were, C# would defer the whole body, including this check, until the caller starts
        // enumerating. Calling it and never touching the result must still throw.
        Assert.Throws<ArgumentNullException>(() => DocxMailMerge.MergeBatch(null!,
            Array.Empty<IReadOnlyDictionary<string, string>>()));
        Assert.Throws<ArgumentNullException>(() => DocxMailMerge.MergeBatch(Simple("FirstName"), null!));
    }

    // ---- MergeBatch: a bad record must name "records", the parameter the caller actually
    // declared, not "values" -- the single-document methods' own parameter name, which
    // RequireValues used to bake in unconditionally ------------------------------------------------

    [Fact]
    public void MergeBatch_ANullRecordNamesTheRecordsParameter_NotValues()
    {
        var records = new IReadOnlyDictionary<string, string>?[]
        {
            new Dictionary<string, string> { ["FirstName"] = "Alice" },
            null,
            new Dictionary<string, string> { ["FirstName"] = "Carol" },
        };

        var ex = Assert.Throws<ArgumentNullException>(
            () => DocxMailMerge.MergeBatch(Simple("FirstName"), records!).ToList());

        Assert.Equal("records", ex.ParamName);
    }

    [Fact]
    public void MergeBatch_ARecordWithANullValue_NamesTheRecordsParameter_NotValues()
    {
        var records = new IReadOnlyDictionary<string, string>[]
        {
            new Dictionary<string, string> { ["FirstName"] = null! },
        };

        var ex = Assert.Throws<ArgumentException>(
            () => DocxMailMerge.MergeBatch(Simple("FirstName"), records).ToList());

        Assert.Equal("records", ex.ParamName);
        Assert.Contains("FirstName", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task MergeBatchAsync_ProducesTheSameDocuments_AsMergeBatch()
    {
        var records = new IReadOnlyDictionary<string, string>[]
        {
            new Dictionary<string, string> { ["FirstName"] = "Alice" },
            new Dictionary<string, string> { ["FirstName"] = "Bob" },
        };
        var template = Simple("FirstName");

        var sync = DocxMailMerge.MergeBatch(template, records).ToList();
        var asyncDocs = new List<byte[]>();
        await foreach (var document in DocxMailMerge.MergeBatchAsync(template, records))
            asyncDocs.Add(document);

        Assert.Equal(sync.Count, asyncDocs.Count);
        // The literal, so this test does not lean on MergeBatch_YieldsOneDocumentPerRecord_InOrder
        // to prove the sync side itself is right -- it is self-contained.
        Assert.Equal("Alice|", Text(asyncDocs[0]));
        // Parity on readable content, not bytes. The cause is durable, not incidental: MergeCore's
        // underlying OfficeIMO.Word.WordDocument.Save() assigns a random w:rsidR value and a random
        // relationship Id on every save, even for logically identical content -- measured directly
        // by diffing two independent MergeCore outputs for the same input, byte-for-byte, entry by
        // entry (word/document.xml's w:rsidR and word/_rels/document.xml.rels' relationship Id both
        // differ between runs, 0 of 5 back-to-back pairs identical). Content is what
        // MergeBatch/MergeBatchAsync promise to agree on, not bytes.
        for (var i = 0; i < sync.Count; i++)
            Assert.Equal(Text(sync[i]), Text(asyncDocs[i]));
    }

    [Fact]
    public async Task MergeBatchAsync_RefusesOnTheBadRecord_MidEnumeration()
    {
        var records = new IReadOnlyDictionary<string, string>[]
        {
            new Dictionary<string, string> { ["FirstName"] = "Alice", ["Balance"] = "100" },
            new Dictionary<string, string> { ["FirstName"] = "Bob" },
        };

        var seen = new List<byte[]>();
        var ex = await Assert.ThrowsAsync<DocumentConversionException>(async () =>
        {
            await foreach (var document in DocxMailMerge.MergeBatchAsync(
                Simple("FirstName", "Balance"), records))
                seen.Add(document);
        });

        Assert.Single(seen);
        Assert.Contains("Balance", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task MergeBatchAsync_HonoursCancellationBetweenRecords()
    {
        var records = new IReadOnlyDictionary<string, string>[]
        {
            new Dictionary<string, string> { ["FirstName"] = "Alice" },
            new Dictionary<string, string> { ["FirstName"] = "Bob" },
            new Dictionary<string, string> { ["FirstName"] = "Carol" },
        };
        // Counts how many records were actually pulled from the sequence. Assert.Single(seen)
        // below is consistent with cancellation being checked either BEFORE or AFTER the next
        // record's merge runs -- either way only one document reaches the consumer. Only
        // pulledCount discriminates: if the check ran after MoveNext() (as it incorrectly did
        // before this fix), record 1's merge would already have completed -- wastefully, and in
        // spite of the cancellation -- before the exception is thrown.
        var pulledCount = 0;
        IEnumerable<IReadOnlyDictionary<string, string>> countingRecords =
            records.Select(r => { pulledCount++; return r; });
        using var cts = new CancellationTokenSource();

        var seen = new List<byte[]>();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
        {
            await foreach (var document in DocxMailMerge.MergeBatchAsync(
                Simple("FirstName"), countingRecords, cts.Token))
            {
                seen.Add(document);
                if (seen.Count == 1) cts.Cancel();
            }
        });

        Assert.Single(seen);
        // Proves record 1's merge never STARTED, not merely that its output never arrived.
        Assert.Equal(1, pulledCount);
    }

    // ---- fixtures ------------------------------------------------------------------------------------

    private static string Text(byte[] docx) => DocxEditor.ExtractText(docx).Replace("\n", string.Empty);

    private static int FieldCount(byte[] docx)
    {
        using var ms = new MemoryStream(docx, writable: false);
        using var w = WordprocessingDocument.Open(ms, false);
        Body body = w.MainDocumentPart!.Document!.Body!;
        return body.Descendants<SimpleField>().Count() + body.Descendants<FieldChar>().Count();
    }

    private static SimpleField Field(string instruction, string shown)
        => new(new Run(new Text(shown))) { Instruction = instruction };

    /// <summary>A template using <c>w:fldSimple</c> — what generators and hand-built documents write.</summary>
    private static byte[] Simple(params string[] names) => Build(body =>
    {
        var p = new Paragraph();
        foreach (string name in names)
        {
            p.Append(Field($" MERGEFIELD {name} \\* MERGEFORMAT ", $"«{name}»"));
            p.Append(new Run(new Text("|")));
        }
        body.Append(p);
    });

    /// <summary>A template using the complex field form — what Word itself writes.</summary>
    private static byte[] Complex(string name) => Build(body =>
    {
        var p = new Paragraph(new Run(new Text("Dear ")));
        p.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
        p.Append(new Run(new FieldCode($" MERGEFIELD {name} \\* MERGEFORMAT ")
        { Space = SpaceProcessingModeValues.Preserve }));
        p.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }));
        p.Append(new Run(new Text($"«{name}»")));
        p.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        p.Append(new Run(new Text(", welcome.")));
        body.Append(p);
    });

    private static byte[] Build(Action<Body> fill)
    {
        using var ms = new MemoryStream();
        using (var d = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            MainDocumentPart main = d.AddMainDocumentPart();
            var body = new Body();
            fill(body);
            main.Document = new Document(body);
            main.Document.Save();
        }
        return ms.ToArray();
    }
}
