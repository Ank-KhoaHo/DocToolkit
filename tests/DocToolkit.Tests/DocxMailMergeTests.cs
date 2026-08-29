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

    [Fact]
    public async Task MergeBatchAsync_ThrowsForAnAlreadyCancelledToken_BeforePullingAnyRecord()
    {
        var records = new IReadOnlyDictionary<string, string>[]
        {
            new Dictionary<string, string> { ["FirstName"] = "Alice" },
        };
        var pulledCount = 0;
        IEnumerable<IReadOnlyDictionary<string, string>> countingRecords =
            records.Select(r => { pulledCount++; return r; });
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
        {
            await foreach (var document in DocxMailMerge.MergeBatchAsync(
                Simple("FirstName"), countingRecords, cts.Token))
            {
            }
        });

        // The check in MergeBatchAsync's loop runs before the first MoveNext() too -- an
        // already-cancelled token must refuse before record 0's merge ever starts, not merely
        // before its output is yielded.
        Assert.Equal(0, pulledCount);
    }

    // ---- MergeBatchWithReport: lenient, never throws for a bad record ---------------------------

    [Fact]
    public void MergeBatchWithReport_NeverThrowsForABadRecord_AndProcessesEveryRecordRegardless()
    {
        var records = new IReadOnlyDictionary<string, string>[]
        {
            new Dictionary<string, string> { ["FirstName"] = "Alice", ["Balance"] = "100" },
            new Dictionary<string, string> { ["FirstName"] = "Bob" }, // Balance missing on purpose
            new Dictionary<string, string> { ["FirstName"] = "Carol", ["Balance"] = "300" },
        };

        var items = DocxMailMerge.MergeBatchWithReport(Simple("FirstName", "Balance"), records).ToList();

        Assert.Equal(3, items.Count);

        Assert.Equal(0, items[0].RecordIndex);
        Assert.True(items[0].Report.IsComplete);
        Assert.Equal("Alice|100|", Text(items[0].Document));

        Assert.Equal(1, items[1].RecordIndex);
        Assert.False(items[1].Report.IsComplete);
        Assert.Equal("Balance", Assert.Single(items[1].Report.MissingFieldNames));
        Assert.Contains("«Balance»", Text(items[1].Document), StringComparison.Ordinal);

        // Record 2 still merged successfully -- the record AFTER a bad one is not skipped.
        Assert.Equal(2, items[2].RecordIndex);
        Assert.True(items[2].Report.IsComplete);
        Assert.Equal("Carol|300|", Text(items[2].Document));
    }

    [Fact]
    public void MergeBatchWithReport_OnAnEmptyRecordSequence_ProducesNoItems()
    {
        var items = DocxMailMerge.MergeBatchWithReport(
            Simple("FirstName"), Array.Empty<IReadOnlyDictionary<string, string>>());

        Assert.Empty(items);
    }

    [Fact]
    public void MergeBatchWithReport_ValidatesArgumentsImmediately_NotOnlyWhenEnumerated()
    {
        Assert.Throws<ArgumentNullException>(() => DocxMailMerge.MergeBatchWithReport(null!,
            Array.Empty<IReadOnlyDictionary<string, string>>()));
        Assert.Throws<ArgumentNullException>(() =>
            DocxMailMerge.MergeBatchWithReport(Simple("FirstName"), null!));
    }

    [Fact]
    public async Task MergeBatchWithReportAsync_NeverThrowsForABadRecord()
    {
        var records = new IReadOnlyDictionary<string, string>[]
        {
            new Dictionary<string, string> { ["FirstName"] = "Alice" },
            new Dictionary<string, string>(), // FirstName missing
        };

        var items = new List<DocxMailMergeBatchItem>();
        await foreach (var item in DocxMailMerge.MergeBatchWithReportAsync(Simple("FirstName"), records))
            items.Add(item);

        Assert.Equal(2, items.Count);
        Assert.True(items[0].Report.IsComplete);
        Assert.False(items[1].Report.IsComplete);
    }

    [Fact]
    public async Task MergeBatchWithReportAsync_OnAnEmptyRecordSequence_ProducesNoItems()
    {
        var items = new List<DocxMailMergeBatchItem>();
        await foreach (var item in DocxMailMerge.MergeBatchWithReportAsync(
            Simple("FirstName"), Array.Empty<IReadOnlyDictionary<string, string>>()))
            items.Add(item);

        Assert.Empty(items);
    }

    [Fact]
    public async Task MergeBatchWithReportAsync_ProducesTheSameItems_AsMergeBatchWithReport()
    {
        var records = new IReadOnlyDictionary<string, string>[]
        {
            new Dictionary<string, string> { ["FirstName"] = "Alice", ["Balance"] = "100" },
            new Dictionary<string, string> { ["FirstName"] = "Bob" }, // Balance missing on purpose
        };
        var template = Simple("FirstName", "Balance");

        var sync = DocxMailMerge.MergeBatchWithReport(template, records).ToList();
        var asyncItems = new List<DocxMailMergeBatchItem>();
        await foreach (var item in DocxMailMerge.MergeBatchWithReportAsync(template, records))
            asyncItems.Add(item);

        Assert.Equal(sync.Count, asyncItems.Count);
        // The literal, so this test does not lean on
        // MergeBatchWithReport_NeverThrowsForABadRecord_AndProcessesEveryRecordRegardless to prove
        // the sync side itself is right -- it is self-contained.
        Assert.Equal("Alice|100|", Text(asyncItems[0].Document));
        // Parity on readable content, not bytes -- see MergeBatchAsync_ProducesTheSameDocuments_AsMergeBatch
        // for why: MergeCore's underlying OfficeIMO.Word.WordDocument.Save() assigns a random
        // w:rsidR value and a random relationship Id on every save, even for logically identical
        // content, so two independent saves of the same input are never byte-identical. Content is
        // what MergeBatchWithReport/MergeBatchWithReportAsync promise to agree on, not bytes.
        for (var i = 0; i < sync.Count; i++)
        {
            Assert.Equal(Text(sync[i].Document), Text(asyncItems[i].Document));
            // DocxMailMergeBatchItem also carries a Report, unlike MergeBatch's plain byte[] --
            // parity has to cover that too, not only the document text.
            Assert.Equal(sync[i].Report.IsComplete, asyncItems[i].Report.IsComplete);
        }
    }

    [Fact]
    public async Task MergeBatchWithReportAsync_HonoursCancellationBetweenRecords()
    {
        var records = new IReadOnlyDictionary<string, string>[]
        {
            new Dictionary<string, string> { ["FirstName"] = "Alice" },
            new Dictionary<string, string> { ["FirstName"] = "Bob" },
            new Dictionary<string, string> { ["FirstName"] = "Carol" },
        };
        // Counts how many records were actually pulled from the sequence. Assert.Single(seen)
        // below is consistent with cancellation being checked either BEFORE or AFTER the next
        // record's merge runs -- either way only one item reaches the consumer. Only pulledCount
        // discriminates: if the check ran after MoveNext() (the plan document's original buggy
        // `foreach` shape), record 1's merge would already have completed -- wastefully, and in
        // spite of the cancellation -- before the exception is thrown.
        var pulledCount = 0;
        IEnumerable<IReadOnlyDictionary<string, string>> countingRecords =
            records.Select(r => { pulledCount++; return r; });
        using var cts = new CancellationTokenSource();

        var seen = new List<DocxMailMergeBatchItem>();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
        {
            await foreach (var item in DocxMailMerge.MergeBatchWithReportAsync(
                Simple("FirstName"), countingRecords, cts.Token))
            {
                seen.Add(item);
                if (seen.Count == 1) cts.Cancel();
            }
        });

        Assert.Single(seen);
        // Proves record 1's merge never STARTED, not merely that its output never arrived.
        Assert.Equal(1, pulledCount);
    }

    // ---- MergeBatchToFiles: strict, writes to disk, refuses on a path collision -----------------

    [Fact]
    public void MergeBatchToFiles_WritesOneFilePerRecord_AtTheFactorysPaths()
    {
        var dir = Directory.CreateTempSubdirectory("DocxMailMergeTests-");
        try
        {
            var templatePath = Path.Combine(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName"));
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice" },
                new Dictionary<string, string> { ["FirstName"] = "Bob" },
            };

            // Captures BOTH arguments outputPathFactory was called with, not just the index --
            // a factory that ignores its own second parameter would still pass a test that only
            // checks the index.
            var seenArgs = new List<(int Index, IReadOnlyDictionary<string, string> Record)>();
            var paths = DocxMailMerge.MergeBatchToFiles(templatePath, records, (i, r) =>
            {
                seenArgs.Add((i, r));
                return Path.Combine(dir.FullName, $"out-{i}.docx");
            });

            Assert.Equal(2, paths.Count);
            Assert.Equal("Alice|", Text(File.ReadAllBytes(paths[0])));
            Assert.Equal("Bob|", Text(File.ReadAllBytes(paths[1])));

            Assert.Equal(2, seenArgs.Count);
            Assert.Equal((0, records[0]), seenArgs[0]);
            Assert.Equal((1, records[1]), seenArgs[1]);
        }
        finally
        {
            dir.Delete(recursive: true);
        }
    }

    [Fact]
    public void MergeBatchToFiles_RefusesOnTheBadRecord_AndWritesNothingAfterIt()
    {
        var dir = Directory.CreateTempSubdirectory("DocxMailMergeTests-");
        try
        {
            var templatePath = Path.Combine(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName", "Balance"));
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice", ["Balance"] = "100" },
                new Dictionary<string, string> { ["FirstName"] = "Bob" }, // Balance missing
                new Dictionary<string, string> { ["FirstName"] = "Carol", ["Balance"] = "300" },
            };

            var ex = Assert.Throws<DocumentConversionException>(() =>
                DocxMailMerge.MergeBatchToFiles(templatePath, records,
                    (i, r) => Path.Combine(dir.FullName, $"out-{i}.docx")));

            Assert.Contains("Balance", ex.Message, StringComparison.Ordinal);
            Assert.True(File.Exists(Path.Combine(dir.FullName, "out-0.docx")));
            Assert.False(File.Exists(Path.Combine(dir.FullName, "out-1.docx")));
            Assert.False(File.Exists(Path.Combine(dir.FullName, "out-2.docx")));
        }
        finally
        {
            dir.Delete(recursive: true);
        }
    }

    [Fact]
    public void MergeBatchToFiles_RefusesACollidingOutputPathFactory_AndWritesNothingAtAll()
    {
        var dir = Directory.CreateTempSubdirectory("DocxMailMergeTests-");
        try
        {
            var templatePath = Path.Combine(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName"));
            var collidingPath = Path.Combine(dir.FullName, "collide.docx");
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice" },
                new Dictionary<string, string> { ["FirstName"] = "Bob" },
                new Dictionary<string, string> { ["FirstName"] = "Carol" },
            };

            // Records 0 and 2 collide; record 1 does not.
            var ex = Assert.Throws<ArgumentException>(() =>
                DocxMailMerge.MergeBatchToFiles(templatePath, records,
                    (i, r) => i == 1 ? Path.Combine(dir.FullName, "unique.docx") : collidingPath));

            Assert.Equal("outputPathFactory", ex.ParamName);
            // "Records 0 and 2", not bare "0"/"2" -- the temp directory's own generated name can
            // easily contain those digits as a substring, which would make a bare-digit assertion
            // pass whether or not the message actually names the right records.
            Assert.Contains("Records 0 and 2", ex.Message, StringComparison.Ordinal);
            // Nothing was written at all -- not even the record that never collided, and not even
            // the "winning" one an unguarded delegation to the engine would silently have produced.
            Assert.False(File.Exists(collidingPath));
            Assert.False(File.Exists(Path.Combine(dir.FullName, "unique.docx")));
        }
        finally
        {
            dir.Delete(recursive: true);
        }
    }

    [Fact]
    public void MergeBatchToFiles_RefusesANullRecord_BeforeCallingOutputPathFactory()
    {
        var dir = Directory.CreateTempSubdirectory("DocxMailMergeTests-");
        try
        {
            var templatePath = Path.Combine(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName"));
            var records = new IReadOnlyDictionary<string, string>?[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice" },
                null,
                new Dictionary<string, string> { ["FirstName"] = "Carol" },
            };

            // Reads the record unconditionally -- if this were ever invoked on the null entry it
            // would throw NullReferenceException itself, never reaching RequireValues's
            // ArgumentNullException. Record 0 is valid, so the factory DOES run once before the
            // null record is reached -- calledIndices pins that directly, proving validation
            // happens per-record rather than only before the very first call (a
            // validate-all-then-compute-all-paths implementation would also throw
            // ArgumentNullException here, but would never call the factory at all).
            var calledIndices = new List<int>();
            var ex = Assert.Throws<ArgumentNullException>(() =>
                DocxMailMerge.MergeBatchToFiles(templatePath, records!,
                    (i, r) => { calledIndices.Add(i); return Path.Combine(dir.FullName, r["FirstName"] + ".docx"); }));

            Assert.Equal("records", ex.ParamName);
            Assert.Equal([0], calledIndices);
            // Nothing was written -- not even record 0, whose path was computed successfully
            // before the null record was ever reached.
            Assert.False(File.Exists(Path.Combine(dir.FullName, "Alice.docx")));
            Assert.False(File.Exists(Path.Combine(dir.FullName, "Carol.docx")));
        }
        finally
        {
            dir.Delete(recursive: true);
        }
    }

    [Fact]
    public void MergeBatchToFiles_RefusesANullOrBlankOutputPath()
    {
        var dir = Directory.CreateTempSubdirectory("DocxMailMergeTests-");
        try
        {
            var templatePath = Path.Combine(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName"));
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice" },
                new Dictionary<string, string> { ["FirstName"] = "Bob" },
            };

            // A null path used to leak out of CheckNoPathCollisions's internal Dictionary as
            // ParamName "key"; a blank one used to leak out of File.WriteAllBytes as ParamName
            // "path". Neither names the actual culprit, outputPathFactory -- covering both here
            // means neither leak can come back unnoticed.
            var nullEx = Assert.Throws<ArgumentException>(() =>
                DocxMailMerge.MergeBatchToFiles(templatePath, records,
                    (i, r) => i == 0 ? Path.Combine(dir.FullName, "out-0.docx") : null!));
            Assert.Equal("outputPathFactory", nullEx.ParamName);
            // Proves the batch is refused before any write, not mid-batch -- record 0's path was
            // already computed (and would have been valid) by the time record 1's bad path is found.
            Assert.False(File.Exists(Path.Combine(dir.FullName, "out-0.docx")));

            var blankEx = Assert.Throws<ArgumentException>(() =>
                DocxMailMerge.MergeBatchToFiles(templatePath, records,
                    (i, r) => i == 0 ? Path.Combine(dir.FullName, "out-0.docx") : "   "));
            Assert.Equal("outputPathFactory", blankEx.ParamName);
            Assert.False(File.Exists(Path.Combine(dir.FullName, "out-0.docx")));
        }
        finally
        {
            dir.Delete(recursive: true);
        }
    }

    [Fact]
    public void MergeBatchToFiles_OnAnEmptyRecordSequence_WritesNothingAndReturnsEmpty()
    {
        var dir = Directory.CreateTempSubdirectory("DocxMailMergeTests-");
        try
        {
            var templatePath = Path.Combine(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName"));

            var paths = DocxMailMerge.MergeBatchToFiles(templatePath,
                Array.Empty<IReadOnlyDictionary<string, string>>(), (i, r) => "unused.docx");

            Assert.Empty(paths);
        }
        finally
        {
            dir.Delete(recursive: true);
        }
    }

    [Fact]
    public void MergeBatchToFiles_ThrowsWhenTheTemplatePathDoesNotExist()
    {
        Assert.Throws<FileNotFoundException>(() =>
            DocxMailMerge.MergeBatchToFiles("does-not-exist.docx",
                Array.Empty<IReadOnlyDictionary<string, string>>(), (i, r) => "unused.docx"));
    }

    [Fact]
    public void MergeBatchToFiles_RejectsNullArguments()
    {
        Assert.Throws<ArgumentNullException>(() => DocxMailMerge.MergeBatchToFiles(
            null!, Array.Empty<IReadOnlyDictionary<string, string>>(), (i, r) => "unused.docx"));
    }

    [Fact]
    public async Task MergeBatchToFilesAsync_ProducesTheSameTextContent_AsMergeBatchToFiles()
    {
        // NOT byte-for-byte: MergeCore's underlying OfficeIMO.Word.WordDocument.Load/Save cycle is
        // not byte-deterministic across two independent saves of logically identical content --
        // measured in Task 1 (a diff appears at the first ZIP entry's CRC-32), matching the same
        // non-determinism already documented elsewhere in this repo for OfficeIMO-authored output
        // (DocxEditorFillRowsTests.cs, WorkbookEditorServiceTests.cs). Text equality is the right
        // comparison here, the same fix Task 1 already applied to the byte[] forms' own pin test.
        var dir = Directory.CreateTempSubdirectory("DocxMailMergeTests-");
        try
        {
            var templatePath = Path.Combine(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName", "Balance"));
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice", ["Balance"] = "100" },
                new Dictionary<string, string> { ["FirstName"] = "Bob", ["Balance"] = "250" },
            };

            var syncPaths = DocxMailMerge.MergeBatchToFiles(templatePath, records,
                (i, r) => Path.Combine(dir.FullName, $"sync-{i}.docx"));
            var asyncPaths = await DocxMailMerge.MergeBatchToFilesAsync(templatePath, records,
                (i, r) => Path.Combine(dir.FullName, $"async-{i}.docx"));

            Assert.Equal(syncPaths.Count, asyncPaths.Count);
            for (var i = 0; i < syncPaths.Count; i++)
                Assert.Equal(Text(File.ReadAllBytes(syncPaths[i])), Text(File.ReadAllBytes(asyncPaths[i])));
        }
        finally
        {
            dir.Delete(recursive: true);
        }
    }

    [Fact]
    public async Task MergeBatchToFilesAsync_HonoursCancellationBetweenRecords()
    {
        var dir = Directory.CreateTempSubdirectory("DocxMailMergeTests-");
        try
        {
            var templatePath = Path.Combine(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName"));
            var paths = new[]
            {
                Path.Combine(dir.FullName, "out-0.docx"),
                Path.Combine(dir.FullName, "out-1.docx"),
                Path.Combine(dir.FullName, "out-2.docx"),
            };

            using var started = new ManualResetEventSlim(false);
            using var release = new ManualResetEventSlim(false);
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice" },
                new BlockingValues(new Dictionary<string, string> { ["FirstName"] = "Bob" }, started, release),
                new Dictionary<string, string> { ["FirstName"] = "Carol" },
            };
            using var cts = new CancellationTokenSource();

            // Task.Run, not a direct call -- if FilePipeline.ReadAsync's await happens to complete
            // synchronously (plausible for a small, already-cached template file), the WHOLE call
            // -- BlockingValues's block included -- would otherwise run inline on this thread,
            // deadlocking before it ever reaches started.Wait() below.
            var task = Task.Run(() => DocxMailMerge.MergeBatchToFilesAsync(templatePath, records, (i, r) => paths[i], cts.Token));

            // Deterministic, not a timing race -- see BlockingValues's doc comment. Record 1's
            // merge only reaches this signal once its OWN cancellation check has already passed,
            // which can only happen once record 0's merge AND write have both fully completed
            // (the loop is strictly sequential), so waiting on it proves record 0 is done rather
            // than guessing from a poll. Bounded rather than unconditional so a future regression
            // fails this test instead of hanging the run.
            try
            {
                Assert.True(started.Wait(TimeSpan.FromSeconds(30)), "record 1's merge never started.");
                cts.Cancel();
            }
            finally
            {
                // Unconditionally, even if the wait above timed out and the assertion already
                // threw -- otherwise a regression that hangs record 1 leaves it blocked forever on
                // an event the `using` declarations are about to dispose out from under it.
                release.Set();
            }

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => task);

            // Record 1's own check had already passed before it was ever blocked, so it always
            // finishes once released -- record 2 is the one this proves cancellation stops.
            Assert.True(File.Exists(paths[0]));
            Assert.True(File.Exists(paths[1]));
            Assert.False(File.Exists(paths[2]));
        }
        finally
        {
            dir.Delete(recursive: true);
        }
    }

    // ---- MergeBatchToFilesWithReport: lenient, never throws for a bad record --------------------

    [Fact]
    public void MergeBatchToFilesWithReport_NeverThrowsForABadRecord_AndWritesEveryFile()
    {
        var dir = Directory.CreateTempSubdirectory("DocxMailMergeTests-");
        try
        {
            var templatePath = Path.Combine(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName", "Balance"));
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice", ["Balance"] = "100" },
                new Dictionary<string, string> { ["FirstName"] = "Bob" }, // Balance missing
                new Dictionary<string, string> { ["FirstName"] = "Carol", ["Balance"] = "300" },
            };

            var items = DocxMailMerge.MergeBatchToFilesWithReport(templatePath, records,
                (i, r) => Path.Combine(dir.FullName, $"out-{i}.docx"));

            Assert.Equal(3, items.Count);
            Assert.True(items[0].Report.IsComplete);
            Assert.True(File.Exists(items[0].OutputPath));

            Assert.False(items[1].Report.IsComplete);
            Assert.True(File.Exists(items[1].OutputPath));
            Assert.Contains("«Balance»", Text(File.ReadAllBytes(items[1].OutputPath)), StringComparison.Ordinal);

            // Record 2 still merged successfully -- the record AFTER a bad one is not skipped,
            // matching MergeBatchWithReport's own established convention for the in-memory form.
            Assert.True(items[2].Report.IsComplete);
            Assert.True(File.Exists(items[2].OutputPath));
            Assert.Equal("Carol|300|", Text(File.ReadAllBytes(items[2].OutputPath)));
        }
        finally
        {
            dir.Delete(recursive: true);
        }
    }

    [Fact]
    public void MergeBatchToFilesWithReport_RefusesACollidingOutputPathFactory_TheSameAsTheStrictForm()
    {
        var dir = Directory.CreateTempSubdirectory("DocxMailMergeTests-");
        try
        {
            var templatePath = Path.Combine(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName"));
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice" },
                new Dictionary<string, string> { ["FirstName"] = "Bob" },
            };

            var ex = Assert.Throws<ArgumentException>(() =>
                DocxMailMerge.MergeBatchToFilesWithReport(templatePath, records,
                    (i, r) => Path.Combine(dir.FullName, "collide.docx")));

            Assert.Equal("outputPathFactory", ex.ParamName);
            Assert.False(File.Exists(Path.Combine(dir.FullName, "collide.docx")));
        }
        finally
        {
            dir.Delete(recursive: true);
        }
    }

    [Fact]
    public void MergeBatchToFilesWithReport_RefusesANullRecord_BeforeCallingOutputPathFactory()
    {
        var dir = Directory.CreateTempSubdirectory("DocxMailMergeTests-");
        try
        {
            var templatePath = Path.Combine(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName"));
            var records = new IReadOnlyDictionary<string, string>?[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice" },
                null,
                new Dictionary<string, string> { ["FirstName"] = "Carol" },
            };

            // Mirrors MergeBatchToFiles_RefusesANullRecord_BeforeCallingOutputPathFactory -- the
            // null-record refusal lives in the shared MergeBatchToFilesCore and runs regardless of
            // strict/lenient, so it must hold here too.
            var calledIndices = new List<int>();
            var ex = Assert.Throws<ArgumentNullException>(() =>
                DocxMailMerge.MergeBatchToFilesWithReport(templatePath, records!,
                    (i, r) => { calledIndices.Add(i); return Path.Combine(dir.FullName, r["FirstName"] + ".docx"); }));

            Assert.Equal("records", ex.ParamName);
            Assert.Equal([0], calledIndices);
            Assert.False(File.Exists(Path.Combine(dir.FullName, "Alice.docx")));
            Assert.False(File.Exists(Path.Combine(dir.FullName, "Carol.docx")));
        }
        finally
        {
            dir.Delete(recursive: true);
        }
    }

    [Fact]
    public void MergeBatchToFilesWithReport_RefusesANullOrBlankOutputPath()
    {
        var dir = Directory.CreateTempSubdirectory("DocxMailMergeTests-");
        try
        {
            var templatePath = Path.Combine(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName"));
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice" },
                new Dictionary<string, string> { ["FirstName"] = "Bob" },
            };

            // Mirrors MergeBatchToFiles_RefusesANullOrBlankOutputPath -- the null/blank-path
            // refusal is equally unconditional, run from the shared MergeBatchToFilesCore.
            var nullEx = Assert.Throws<ArgumentException>(() =>
                DocxMailMerge.MergeBatchToFilesWithReport(templatePath, records,
                    (i, r) => i == 0 ? Path.Combine(dir.FullName, "out-0.docx") : null!));
            Assert.Equal("outputPathFactory", nullEx.ParamName);
            Assert.False(File.Exists(Path.Combine(dir.FullName, "out-0.docx")));

            var blankEx = Assert.Throws<ArgumentException>(() =>
                DocxMailMerge.MergeBatchToFilesWithReport(templatePath, records,
                    (i, r) => i == 0 ? Path.Combine(dir.FullName, "out-0.docx") : "   "));
            Assert.Equal("outputPathFactory", blankEx.ParamName);
            Assert.False(File.Exists(Path.Combine(dir.FullName, "out-0.docx")));
        }
        finally
        {
            dir.Delete(recursive: true);
        }
    }

    [Fact]
    public async Task MergeBatchToFilesWithReportAsync_ProducesTheSameTextContent_AsMergeBatchToFilesWithReport()
    {
        // NOT byte-for-byte -- see MergeBatchToFilesAsync_ProducesTheSameTextContent_AsMergeBatchToFiles
        // above for why: MergeCore's OfficeIMO.Word save cycle is not byte-deterministic across two
        // independent saves of the same logical content.
        var dir = Directory.CreateTempSubdirectory("DocxMailMergeTests-");
        try
        {
            var templatePath = Path.Combine(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName", "Balance"));
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice", ["Balance"] = "100" },
                new Dictionary<string, string> { ["FirstName"] = "Bob" }, // Balance missing
            };

            var syncItems = DocxMailMerge.MergeBatchToFilesWithReport(templatePath, records,
                (i, r) => Path.Combine(dir.FullName, $"sync-{i}.docx"));
            var asyncItems = await DocxMailMerge.MergeBatchToFilesWithReportAsync(templatePath, records,
                (i, r) => Path.Combine(dir.FullName, $"async-{i}.docx"));

            Assert.Equal(syncItems.Count, asyncItems.Count);
            for (var i = 0; i < syncItems.Count; i++)
            {
                Assert.Equal(syncItems[i].Report.IsComplete, asyncItems[i].Report.IsComplete);
                Assert.Equal(
                    Text(File.ReadAllBytes(syncItems[i].OutputPath)),
                    Text(File.ReadAllBytes(asyncItems[i].OutputPath)));
            }
        }
        finally
        {
            dir.Delete(recursive: true);
        }
    }

    [Fact]
    public async Task MergeBatchToFilesWithReportAsync_NeverThrowsForABadRecord()
    {
        var dir = Directory.CreateTempSubdirectory("DocxMailMergeTests-");
        try
        {
            var templatePath = Path.Combine(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName"));
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice" },
                new Dictionary<string, string>(), // FirstName missing
                new Dictionary<string, string> { ["FirstName"] = "Carol" },
            };

            var items = await DocxMailMerge.MergeBatchToFilesWithReportAsync(templatePath, records,
                (i, r) => Path.Combine(dir.FullName, $"out-{i}.docx"));

            Assert.Equal(3, items.Count);
            Assert.True(items[0].Report.IsComplete);
            Assert.False(items[1].Report.IsComplete);
            // Record 2 still merged successfully -- the record AFTER a bad one is not skipped.
            Assert.True(items[2].Report.IsComplete);
            Assert.Equal("Carol|", Text(File.ReadAllBytes(items[2].OutputPath)));
        }
        finally
        {
            dir.Delete(recursive: true);
        }
    }

    [Fact]
    public async Task MergeBatchToFilesWithReportAsync_HonoursCancellationBetweenRecords()
    {
        // Same shape as MergeBatchToFilesAsync_HonoursCancellationBetweenRecords -- see
        // BlockingValues's doc comment for why this is a deterministic interleave rather than a
        // timing race.
        var dir = Directory.CreateTempSubdirectory("DocxMailMergeTests-");
        try
        {
            var templatePath = Path.Combine(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName"));
            var paths = new[]
            {
                Path.Combine(dir.FullName, "out-0.docx"),
                Path.Combine(dir.FullName, "out-1.docx"),
                Path.Combine(dir.FullName, "out-2.docx"),
            };

            using var started = new ManualResetEventSlim(false);
            using var release = new ManualResetEventSlim(false);
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice" },
                new BlockingValues(new Dictionary<string, string> { ["FirstName"] = "Bob" }, started, release),
                new Dictionary<string, string> { ["FirstName"] = "Carol" },
            };
            using var cts = new CancellationTokenSource();

            // Task.Run, not a direct call -- see MergeBatchToFilesAsync_HonoursCancellationBetweenRecords's
            // comment for why: a synchronously-completing first await would otherwise run the whole
            // call, BlockingValues's block included, inline on this thread and deadlock.
            var task = Task.Run(() => DocxMailMerge.MergeBatchToFilesWithReportAsync(templatePath, records, (i, r) => paths[i], cts.Token));

            try
            {
                Assert.True(started.Wait(TimeSpan.FromSeconds(30)), "record 1's merge never started.");
                cts.Cancel();
            }
            finally
            {
                // Unconditionally, even if the wait above timed out and the assertion already
                // threw -- otherwise a regression that hangs record 1 leaves it blocked forever on
                // an event the `using` declarations are about to dispose out from under it.
                release.Set();
            }

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => task);

            Assert.True(File.Exists(paths[0]));
            Assert.True(File.Exists(paths[1]));
            Assert.False(File.Exists(paths[2]));
        }
        finally
        {
            dir.Delete(recursive: true);
        }
    }

    // ---- fixtures ------------------------------------------------------------------------------------

    /// <summary>
    /// An <see cref="IReadOnlyDictionary{TKey, TValue}"/> that blocks the first time its
    /// <see cref="Count"/> is read, until released -- used by the file-path
    /// *_HonoursCancellationBetweenRecords tests to get a deterministic synchronization point
    /// inside <c>MergeBatchToFilesCore</c>'s per-record loop, rather than racing against wall-clock
    /// timing.
    /// </summary>
    /// <remarks>
    /// Neither <c>MergeBatchToFilesAsync</c> nor <c>MergeBatchToFilesWithReportAsync</c> is an
    /// async iterator -- each runs its whole write loop synchronously behind one <see
    /// cref="Task"/>, so there is no yield point between records for a consumer to interject a
    /// cancellation. The gap between one record's <c>File.WriteAllBytes</c> call RETURNING and the
    /// next record's cancellation check is a handful of CPU instructions -- measured directly
    /// while implementing this fix: an external thread polling for the file to appear on disk and
    /// racing to cancel before that check runs wins only occasionally, even with a multi-megabyte
    /// payload deliberately keeping the write in flight, because OS thread scheduling can delay a
    /// polling thread by more than that gap. That is not a fixable flakiness; it means the
    /// boundary this class exists to observe cannot be caught from outside by timing at all.
    ///
    /// <b>This class turns the one point <c>DocxMailMerge</c> genuinely reads a record's values
    /// into a synchronization point instead.</b> <c>DocxMailMerge.Copy</c> -- the only place a
    /// record's values are consumed to build the actual merge -- reads <see cref="Count"/> before
    /// enumerating. The earlier validation every record passes through first, before any record is
    /// merged (<c>DocxMailMerge.RequireValues</c>, called once per record while every output path
    /// is computed and collision-checked), only enumerates and never reads <see cref="Count"/>. So
    /// blocking specifically on <see cref="Count"/> is reached exactly once per record, and exactly
    /// when that record's real merge begins -- which can only happen once every record before it
    /// has been fully merged AND written, since the loop is strictly sequential.
    ///
    /// Wrapping record 1's values and waiting for that block to be reached therefore proves record
    /// 0 is completely done, with no polling and no race. It does not prove record 1 was stopped --
    /// its own cancellation check has, by construction, already passed by the time this block is
    /// reached, so record 1 always finishes once released. What it proves is one loop iteration
    /// later than the ideal case: cancelling at that point stops record 2, which exercises the
    /// identical check on the identical code path one record later.
    /// </remarks>
    private sealed class BlockingValues : IReadOnlyDictionary<string, string>
    {
        private readonly IReadOnlyDictionary<string, string> _inner;
        private readonly ManualResetEventSlim _started;
        private readonly ManualResetEventSlim _release;

        public BlockingValues(
            IReadOnlyDictionary<string, string> inner, ManualResetEventSlim started, ManualResetEventSlim release)
        {
            _inner = inner;
            _started = started;
            _release = release;
        }

        public int Count
        {
            get
            {
                _started.Set();
                _release.Wait();
                return _inner.Count;
            }
        }

        public string this[string key] => _inner[key];
        public IEnumerable<string> Keys => _inner.Keys;
        public IEnumerable<string> Values => _inner.Values;
        public bool ContainsKey(string key) => _inner.ContainsKey(key);
        public bool TryGetValue(string key, out string value) => _inner.TryGetValue(key, out value!);
        public IEnumerator<KeyValuePair<string, string>> GetEnumerator() => _inner.GetEnumerator();
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

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
