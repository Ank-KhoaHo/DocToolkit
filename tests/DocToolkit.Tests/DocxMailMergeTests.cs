using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
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
    /// <summary>
    /// Asserts the produced package is schema-valid, not merely readable — through the shared
    /// <see cref="DocxFixtures.Validate"/> helper, whose <c>Office2013</c> version is the one this
    /// repository settled on after <c>Office2007</c> was found blind to a real ordering violation.
    /// </summary>
    /// <remarks>
    /// <b>This matters most where a document is CLONED or CUT rather than typed into.</b> Expanding
    /// a repeating region duplicates a run of block-level elements; expanding a table row clones a
    /// <c>w:tr</c>; resolving a conditional block to false deletes a range. Every one of those can
    /// produce a package whose text reads perfectly and whose XML Word refuses to open, and no
    /// assertion on extracted text in this file can see the difference. Matches
    /// <see cref="DocxEditorFootnoteEndnoteTocTests"/>'s helper of the same name rather than
    /// hand-rolling a third <c>OpenXmlValidator</c> convention.
    /// </remarks>
    private static void AssertValid(byte[] docx)
    {
        var errors = DocxFixtures.Validate(docx);
        Assert.True(errors.Count == 0,
            "expected a schema-valid package, got:\n" +
            string.Join("\n", errors.Take(3).Select(e => "  " + e.Description)));
    }

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

    // ---- A75: ConditionalBlockNames / RepeatingBlockNames / expanded issue kinds --------------

    [Fact]
    public void InspectTemplate_DiscoversConditionalAndRepeatingBlockNames()
    {
        byte[] template = Build(body =>
        {
            body.Append(new Paragraph(new Run(new Text("{{#ShowDiscount}}"))));
            body.Append(new Paragraph(new Run(new Text("discount"))));
            body.Append(new Paragraph(new Run(new Text("{{/ShowDiscount}}"))));
            body.Append(new Paragraph(new Run(new Text("{{#each Items}}"))));
            body.Append(new Paragraph(new Run(new Text("item"))));
            body.Append(new Paragraph(new Run(new Text("{{/each Items}}"))));
        });

        DocxMailMergeTemplate inspection = DocxMailMerge.InspectTemplate(template);

        Assert.Equal(new[] { "ShowDiscount" }, inspection.ConditionalBlockNames);
        Assert.Equal(new[] { "Items" }, inspection.RepeatingBlockNames);
        Assert.True(inspection.IsValid);
    }

    [Fact]
    public void InspectTemplate_MalformedConditionalMarker_ReportsTheSpecificKind()
    {
        // Unmatched start -- no matching {{/ShowA}} anywhere in the document.
        byte[] template = Build(body =>
        {
            body.Append(new Paragraph(new Run(new Text("{{#ShowA}}"))));
            body.Append(new Paragraph(new Run(new Text("content"))));
        });

        DocxMailMergeTemplate inspection = DocxMailMerge.InspectTemplate(template);

        Assert.False(inspection.IsValid);
        DocxMailMergeIssue issue = Assert.Single(inspection.Issues);
        Assert.Equal(DocxMailMergeIssueKind.UnmatchedConditionalStart, issue.Kind);
        Assert.Equal("ShowA", issue.Name);
    }

    [Fact]
    public void InspectTemplate_MismatchedRepeatingEnd_ReportsTheSpecificKind()
    {
        byte[] template = Build(body =>
        {
            body.Append(new Paragraph(new Run(new Text("{{#each Items}}"))));
            body.Append(new Paragraph(new Run(new Text("content"))));
            body.Append(new Paragraph(new Run(new Text("{{/each Other}}"))));
        });

        DocxMailMergeTemplate inspection = DocxMailMerge.InspectTemplate(template);

        Assert.False(inspection.IsValid);
        DocxMailMergeIssue issue = Assert.Single(inspection.Issues);
        Assert.Equal(DocxMailMergeIssueKind.MismatchedRepeatingBlockEnd, issue.Kind);
        Assert.Equal("Other", issue.Name);
    }

    [Fact]
    public void InspectTemplate_WithNoConditionalOrRepeatingMarkers_ReportsEmptyLists()
    {
        DocxMailMergeTemplate inspection = DocxMailMerge.InspectTemplate(Simple("FirstName"));

        Assert.Empty(inspection.ConditionalBlockNames);
        Assert.Empty(inspection.RepeatingBlockNames);
    }

    [Fact]
    public void DocxMailMergeIssueKind_HasTwelveValues()
    {
        // Pins the total so a future OfficeIMO release adding a 13th kind is visible here rather
        // than silently collapsing into Other with nobody noticing.
        Assert.Equal(12, Enum.GetValues<DocxMailMergeIssueKind>().Length);
    }

    [Fact]
    public void DocxMailMergeBlockData_ExposesValuesAndRegions()
    {
        var nested = new DocxMailMergeBlockData(new Dictionary<string, string> { ["Sku"] = "A1" });
        var data = new DocxMailMergeBlockData(
            new Dictionary<string, string> { ["OrderId"] = "1001" },
            new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
            {
                ["Lines"] = new[] { nested }
            });

        Assert.Equal("1001", data.Values["OrderId"]);
        Assert.Equal("A1", data.Regions!["Lines"].Single().Values["Sku"]);
    }

    [Fact]
    public void DocxMailMergeTableRowGroup_ExposesValuesAndRows()
    {
        var group = new DocxMailMergeTableRowGroup(
            new Dictionary<string, string> { ["GroupName"] = "Fruits" },
            new[] { new Dictionary<string, string> { ["Item"] = "Apple" } });

        Assert.Equal("Fruits", group.Values["GroupName"]);
        Assert.Equal("Apple", group.Rows.Single()["Item"]);
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

    // ---- A75: MergeConditional -----------------------------------------------------------------

    [Fact]
    public void MergeConditional_IncludesTheBlockWhenTrue_AndRemovesMarkers()
    {
        byte[] merged = DocxMailMerge.MergeConditional(
            ConditionalTemplate("ShowDiscount", "Discount applies"),
            new Dictionary<string, bool> { ["ShowDiscount"] = true });

        Assert.Equal("BeforeDiscount appliesAfter", Text(merged));
        AssertValid(merged);
    }

    [Fact]
    public void MergeConditional_RemovesTheWholeBlockWhenFalse()
    {
        byte[] merged = DocxMailMerge.MergeConditional(
            ConditionalTemplate("ShowDiscount", "Discount applies"),
            new Dictionary<string, bool> { ["ShowDiscount"] = false });

        Assert.Equal("BeforeAfter", Text(merged));
        AssertValid(merged);
    }

    [Fact]
    public void MergeConditional_RefusesWhenAConditionIsMissing()
    {
        byte[] template = ConditionalTemplate("ShowDiscount", "Discount applies");

        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxMailMerge.MergeConditional(template, new Dictionary<string, bool>()));

        Assert.Contains("ShowDiscount", ex.Message);
    }

    [Fact]
    public void MergeConditional_RefusesBeforeTouchingTheDocument_OnAnUnbalancedMarker()
    {
        // Unmatched start -- InspectTemplate catches this gracefully, so the strict form should
        // refuse via DocumentConversionException, never let OfficeIMO's raw InvalidOperationException
        // escape.
        byte[] template = Build(body =>
        {
            body.Append(new Paragraph(new Run(new Text("{{#ShowA}}"))));
            body.Append(new Paragraph(new Run(new Text("content"))));
        });

        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxMailMerge.MergeConditional(template, new Dictionary<string, bool> { ["ShowA"] = true }));

        Assert.Contains("ShowA", ex.Message);
    }

    [Fact]
    public async Task MergeConditionalAsync_MatchesTheByteArrayForm()
    {
        byte[] template = ConditionalTemplate("ShowDiscount", "Discount applies");
        using var source = new MemoryStream(template);
        using var destination = new MemoryStream();

        await DocxMailMerge.MergeConditionalAsync(
            source, destination, new Dictionary<string, bool> { ["ShowDiscount"] = true });

        Assert.Equal("BeforeDiscount appliesAfter", Text(destination.ToArray()));
    }

    [Fact]
    public void MergeConditionalWithReport_PadsAMissingConditionWithFalse_AndReportsIt()
    {
        byte[] template = ConditionalTemplate("ShowDiscount", "Discount applies");

        DocxMailMergeBlockResult result = DocxMailMerge.MergeConditionalWithReport(
            template, new Dictionary<string, bool>());

        Assert.Equal("BeforeAfter", Text(result.Document));
        Assert.Equal(new[] { "ShowDiscount" }, result.Report.MissingNames);
        Assert.False(result.Report.IsComplete);
    }

    [Fact]
    public void MergeConditionalWithReport_SuppliedConditionIsNotReportedMissing()
    {
        byte[] template = ConditionalTemplate("ShowDiscount", "Discount applies");

        DocxMailMergeBlockResult result = DocxMailMerge.MergeConditionalWithReport(
            template, new Dictionary<string, bool> { ["ShowDiscount"] = true });

        Assert.Equal("BeforeDiscount appliesAfter", Text(result.Document));
        Assert.Empty(result.Report.MissingNames);
        Assert.True(result.Report.IsComplete);
    }

    [Fact]
    public void MergeConditionalWithReport_StillThrowsOnAnUnbalancedMarker()
    {
        // The one case WithReport cannot "always produce a document" for -- a genuinely unbalanced
        // marker structure makes OfficeIMO's Execute throw regardless of dictionary content.
        byte[] template = Build(body =>
        {
            body.Append(new Paragraph(new Run(new Text("{{#ShowA}}"))));
            body.Append(new Paragraph(new Run(new Text("content"))));
        });

        Assert.Throws<DocumentConversionException>(
            () => DocxMailMerge.MergeConditionalWithReport(template, new Dictionary<string, bool> { ["ShowA"] = true }));
    }

    [Theory]
    [InlineData(DocxMailMergeIssueKind.UnmatchedConditionalStart)]
    [InlineData(DocxMailMergeIssueKind.UnmatchedConditionalEnd)]
    [InlineData(DocxMailMergeIssueKind.MismatchedConditionalEnd)]
    public void MergeConditionalWithReport_StillThrows_ForEveryStructuralIssueKind(
        DocxMailMergeIssueKind expectedKind)
    {
        byte[] template = expectedKind switch
        {
            DocxMailMergeIssueKind.UnmatchedConditionalStart => Build(body =>
            {
                body.Append(new Paragraph(new Run(new Text("{{#ShowA}}"))));
                body.Append(new Paragraph(new Run(new Text("content"))));
            }),
            DocxMailMergeIssueKind.UnmatchedConditionalEnd => Build(body =>
            {
                body.Append(new Paragraph(new Run(new Text("content"))));
                body.Append(new Paragraph(new Run(new Text("{{/ShowA}}"))));
            }),
            _ => Build(body =>
            {
                body.Append(new Paragraph(new Run(new Text("{{#ShowA}}"))));
                body.Append(new Paragraph(new Run(new Text("content"))));
                body.Append(new Paragraph(new Run(new Text("{{/ShowB}}"))));
            }),
        };

        // Confirm InspectTemplate would have named this exact kind, then confirm WithReport still
        // throws for it (it cannot pad its way past a structural defect).
        DocxMailMergeTemplate inspection = DocxMailMerge.InspectTemplate(template);
        Assert.Equal(expectedKind, Assert.Single(inspection.Issues).Kind);

        Assert.Throws<DocumentConversionException>(
            () => DocxMailMerge.MergeConditionalWithReport(template, new Dictionary<string, bool> { ["ShowA"] = true, ["ShowB"] = true }));
    }

    [Fact]
    public void MergeConditional_NullConditions_Throws()
    {
        Assert.Throws<ArgumentNullException>(
            () => DocxMailMerge.MergeConditional(ConditionalTemplate("ShowA", "x"), null!));
    }

    private static byte[] ConditionalTemplate(string name, string content) => Build(body =>
    {
        body.Append(new Paragraph(new Run(new Text("Before"))));
        body.Append(new Paragraph(new Run(new Text($"{{{{#{name}}}}}"))));
        body.Append(new Paragraph(new Run(new Text(content))));
        body.Append(new Paragraph(new Run(new Text($"{{{{/{name}}}}}"))));
        body.Append(new Paragraph(new Run(new Text("After"))));
    });

    // ---- A75: MergeRepeating ---------------------------------------------------------------------

    [Fact]
    public void MergeRepeating_ExpandsOnceForEveryRecord()
    {
        byte[] merged = DocxMailMerge.MergeRepeating(
            RepeatingTemplate("Items", "Name"),
            new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>
            {
                ["Items"] = new[]
                {
                    new Dictionary<string, string> { ["Name"] = "Alice" },
                    new Dictionary<string, string> { ["Name"] = "Bob" },
                }
            });

        Assert.Equal("BeforeName: AliceName: BobAfter", Text(merged));
        AssertValid(merged);
    }

    [Fact]
    public void MergeRepeating_EmptySequenceRemovesTheWholeMarkedRegion()
    {
        byte[] merged = DocxMailMerge.MergeRepeating(
            RepeatingTemplate("Items", "Name"),
            new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>
            {
                ["Items"] = Array.Empty<IReadOnlyDictionary<string, string>>()
            });

        Assert.Equal("BeforeAfter", Text(merged));
    }

    [Fact]
    public void MergeRepeating_RefusesWhenARegionIsMissing()
    {
        byte[] template = RepeatingTemplate("Items", "Name");

        var ex = Assert.Throws<DocumentConversionException>(() => DocxMailMerge.MergeRepeating(
            template, new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>()));

        Assert.Contains("Items", ex.Message);
    }

    [Fact]
    public async Task MergeRepeatingAsync_MatchesTheByteArrayForm()
    {
        byte[] template = RepeatingTemplate("Items", "Name");
        using var source = new MemoryStream(template);
        using var destination = new MemoryStream();

        await DocxMailMerge.MergeRepeatingAsync(
            source, destination,
            new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>
            {
                ["Items"] = new[] { new Dictionary<string, string> { ["Name"] = "Alice" } }
            });

        Assert.Equal("BeforeName: AliceAfter", Text(destination.ToArray()));
    }

    [Fact]
    public void MergeRepeatingWithReport_PadsAMissingRegionWithZeroRows_AndReportsIt()
    {
        byte[] template = RepeatingTemplate("Items", "Name");

        DocxMailMergeBlockResult result = DocxMailMerge.MergeRepeatingWithReport(
            template, new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>());

        Assert.Equal("BeforeAfter", Text(result.Document));
        Assert.Equal(new[] { "Items" }, result.Report.MissingNames);
        Assert.False(result.Report.IsComplete);
    }

    [Fact]
    public void MergeRepeatingWithReport_StillThrowsOnAnUnbalancedMarker()
    {
        byte[] template = Build(body =>
        {
            body.Append(new Paragraph(new Run(new Text("{{#each Items}}"))));
            body.Append(new Paragraph(new Run(new Text("content"))));
        });

        Assert.Throws<DocumentConversionException>(() => DocxMailMerge.MergeRepeatingWithReport(
            template, new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>
            {
                ["Items"] = Array.Empty<IReadOnlyDictionary<string, string>>()
            }));
    }

    [Fact]
    public void MergeRepeating_ThenMerge_ComposesInTheRequiredOrder()
    {
        // The composition ordering the design measured: structural expansion first, then the
        // existing field-level merge -- including catching a field left unfilled inside an
        // already-expanded row, which ExecuteRepeatingBlocks itself never reports.
        byte[] template = Build(body =>
        {
            body.Append(new Paragraph(new Run(new Text("{{#each Items}}"))));
            var p = new Paragraph();
            p.Append(new Run(new Text("Name: ")));
            p.Append(Field(" MERGEFIELD Name \\* MERGEFORMAT ", "«Name»"));
            p.Append(new Run(new Text(" Qty: ")));
            p.Append(Field(" MERGEFIELD Qty \\* MERGEFORMAT ", "«Qty»"));
            body.Append(p);
            body.Append(new Paragraph(new Run(new Text("{{/each Items}}"))));
        });

        byte[] expanded = DocxMailMerge.MergeRepeating(
            template,
            new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>
            {
                ["Items"] = new IReadOnlyDictionary<string, string>[]
                {
                    new Dictionary<string, string> { ["Name"] = "Alice", ["Qty"] = "5" },
                    new Dictionary<string, string> { ["Name"] = "Bob" }, // Qty deliberately missing
                }
            });

        DocxMailMergeResult result = DocxMailMerge.MergeWithReport(expanded, new Dictionary<string, string>());

        Assert.False(result.Report.IsComplete);
        Assert.Equal(new[] { "Qty" }, result.Report.MissingFieldNames);
        Assert.Contains("Name: Alice Qty: 5", Text(result.Document));
    }

    private static byte[] RepeatingTemplate(string regionName, string fieldName) => Build(body =>
    {
        body.Append(new Paragraph(new Run(new Text("Before"))));
        body.Append(new Paragraph(new Run(new Text($"{{{{#each {regionName}}}}}"))));
        var p = new Paragraph();
        p.Append(new Run(new Text($"{fieldName}: ")));
        p.Append(Field($" MERGEFIELD {fieldName} \\* MERGEFORMAT ", $"«{fieldName}»"));
        body.Append(p);
        body.Append(new Paragraph(new Run(new Text($"{{{{/each {regionName}}}}}"))));
        body.Append(new Paragraph(new Run(new Text("After"))));
    });

    // ---- A75: MergeRepeatingRegions (nested) --------------------------------------------------

    [Fact]
    public void MergeRepeatingRegions_ExpandsNestedRegionsForEveryRecord()
    {
        byte[] template = NestedRepeatingTemplate();

        byte[] merged = DocxMailMerge.MergeRepeatingRegions(
            template,
            new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
            {
                ["Orders"] = new[]
                {
                    new DocxMailMergeBlockData(
                        new Dictionary<string, string> { ["OrderId"] = "1001" },
                        new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
                        {
                            ["Lines"] = new[]
                            {
                                new DocxMailMergeBlockData(new Dictionary<string, string> { ["Sku"] = "A1" }),
                                new DocxMailMergeBlockData(new Dictionary<string, string> { ["Sku"] = "A2" }),
                            }
                        }),
                }
            });

        Assert.Equal("Order: 1001Line: A1Line: A2", Text(merged));
        AssertValid(merged);
    }

    [Fact]
    public void MergeRepeatingRegions_CorrectlyNestedCall_DoesNotFalsePositiveAsMissing()
    {
        // The nesting trap the design measured: RepeatingBlockNames is flat, so a preflight that
        // only reads top-level dictionary keys would wrongly report "Lines" missing even when it
        // is correctly nested inside "Orders". This must NOT throw.
        byte[] template = NestedRepeatingTemplate();

        DocxMailMergeBlockResult result = DocxMailMerge.MergeRepeatingRegionsWithReport(
            template,
            new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
            {
                ["Orders"] = new[]
                {
                    new DocxMailMergeBlockData(
                        new Dictionary<string, string> { ["OrderId"] = "1001" },
                        new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
                        {
                            ["Lines"] = new[] { new DocxMailMergeBlockData(new Dictionary<string, string> { ["Sku"] = "A1" }) }
                        }),
                }
            });

        Assert.Empty(result.Report.MissingNames);
        Assert.True(result.Report.IsComplete);
    }

    [Fact]
    public void MergeRepeatingRegions_CorrectlyNestedCall_StillProducesTheExpandedDocument()
    {
        // The discriminating half of the trap test above: padding every discovered region name
        // into every nesting level (see MergeRepeatingRegionsCore) must not disturb a call that
        // supplied everything. An assertion on MissingNames alone would pass against an
        // implementation that padded a spurious region over the caller's own rows.
        byte[] template = NestedRepeatingTemplate();

        DocxMailMergeBlockResult result = DocxMailMerge.MergeRepeatingRegionsWithReport(
            template,
            new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
            {
                ["Orders"] = new[]
                {
                    new DocxMailMergeBlockData(
                        new Dictionary<string, string> { ["OrderId"] = "1001" },
                        new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
                        {
                            ["Lines"] = new[] { new DocxMailMergeBlockData(new Dictionary<string, string> { ["Sku"] = "A1" }) }
                        }),
                }
            });

        Assert.Equal("Order: 1001Line: A1", Text(result.Document));
    }

    [Fact]
    public void MergeRepeatingRegions_RefusesWhenTheTopLevelRegionIsMissing()
    {
        byte[] template = NestedRepeatingTemplate();

        var ex = Assert.Throws<DocumentConversionException>(() => DocxMailMerge.MergeRepeatingRegions(
            template, new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>()));

        Assert.Contains("Orders", ex.Message);
    }

    [Fact]
    public void MergeRepeatingRegions_RefusesWhenANestedRegionIsMissingEverywhere()
    {
        // The other half of the flat-names trap: "Lines" is nested, so a preflight that only
        // compared top-level keys could not tell this apart from the correctly-nested call above.
        byte[] template = NestedRepeatingTemplate();

        var ex = Assert.Throws<DocumentConversionException>(() => DocxMailMerge.MergeRepeatingRegions(
            template,
            new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
            {
                ["Orders"] = new[]
                {
                    new DocxMailMergeBlockData(new Dictionary<string, string> { ["OrderId"] = "1001" }),
                }
            }));

        Assert.Contains("Lines", ex.Message);
    }

    [Fact]
    public void MergeRepeatingRegionsWithReport_PadsANestedRegionMissingEverywhere_AndReportsIt()
    {
        // Measured: padding this name at the TOP level does not satisfy the engine -- it resolves a
        // nested marker from its ENCLOSING block's own regions, so the pad has to reach in there.
        byte[] template = NestedRepeatingTemplate();

        DocxMailMergeBlockResult result = DocxMailMerge.MergeRepeatingRegionsWithReport(
            template,
            new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
            {
                ["Orders"] = new[]
                {
                    new DocxMailMergeBlockData(new Dictionary<string, string> { ["OrderId"] = "1001" }),
                }
            });

        Assert.Equal("Order: 1001", Text(result.Document));
        Assert.Equal(new[] { "Lines" }, result.Report.MissingNames);
        Assert.False(result.Report.IsComplete);
    }

    [Fact]
    public void MergeRepeatingRegionsWithReport_PadsANestedRegionMissingForOneRecordOnly()
    {
        // Measured: the engine throws for this too, identically -- and it is invisible to the
        // report, because MissingNames is a NAME-level answer and "Lines" is supplied by the first
        // order. The second order's region is defaulted to zero rows and the report stays complete.
        byte[] template = NestedRepeatingTemplate();

        DocxMailMergeBlockResult result = DocxMailMerge.MergeRepeatingRegionsWithReport(
            template,
            new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
            {
                ["Orders"] = new[]
                {
                    new DocxMailMergeBlockData(
                        new Dictionary<string, string> { ["OrderId"] = "1001" },
                        new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
                        {
                            ["Lines"] = new[] { new DocxMailMergeBlockData(new Dictionary<string, string> { ["Sku"] = "A1" }) }
                        }),
                    new DocxMailMergeBlockData(new Dictionary<string, string> { ["OrderId"] = "1002" }),
                }
            });

        Assert.Equal("Order: 1001Line: A1Order: 1002", Text(result.Document));
        Assert.Empty(result.Report.MissingNames);
        Assert.True(result.Report.IsComplete);
    }

    [Fact]
    public void MergeRepeatingRegions_StrictForm_ThrowsOnAPerRowOmission_UnlikeTheLenientForm()
    {
        // Same template and data shape as MergeRepeatingRegionsWithReport_PadsANestedRegionMissingForOneRecordOnly
        // above -- only the second Orders row omits "Lines". The strict form pads nothing before
        // calling the underlying engine, so this per-row omission reaches it unpadded and it throws,
        // unlike the lenient WithReport form, which defaults the row to zero rows and reports nothing missing.
        byte[] template = NestedRepeatingTemplate();

        var ex = Assert.Throws<DocumentConversionException>(() => DocxMailMerge.MergeRepeatingRegions(
            template,
            new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
            {
                ["Orders"] = new[]
                {
                    new DocxMailMergeBlockData(
                        new Dictionary<string, string> { ["OrderId"] = "1001" },
                        new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
                        {
                            ["Lines"] = new[] { new DocxMailMergeBlockData(new Dictionary<string, string> { ["Sku"] = "A1" }) }
                        }),
                    new DocxMailMergeBlockData(new Dictionary<string, string> { ["OrderId"] = "1002" }),
                }
            }));

        // WHICH layer refused is the whole point, and the exception type alone cannot say: the
        // preflight and the engine both surface as DocumentConversionException. The preflight
        // throws with no inner exception, so an InvalidOperationException inside pins that the
        // ENGINE threw -- which is what makes this a per-row omission the name-level preflight
        // cannot see, rather than a preflight refusal wearing the same clothes.
        Assert.IsType<InvalidOperationException>(ex.InnerException);
    }

    [Fact]
    public void MergeRepeatingRegionsWithReport_StillThrowsOnAnUnbalancedMarker()
    {
        byte[] template = Build(body =>
        {
            body.Append(new Paragraph(new Run(new Text("{{#each Orders}}"))));
            body.Append(new Paragraph(new Run(new Text("content"))));
        });

        Assert.Throws<DocumentConversionException>(() => DocxMailMerge.MergeRepeatingRegionsWithReport(
            template, new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
            {
                ["Orders"] = Array.Empty<DocxMailMergeBlockData>()
            }));
    }

    [Fact]
    public async Task MergeRepeatingRegionsAsync_MatchesTheByteArrayForm()
    {
        byte[] template = NestedRepeatingTemplate();
        using var source = new MemoryStream(template);
        using var destination = new MemoryStream();

        await DocxMailMerge.MergeRepeatingRegionsAsync(
            source, destination,
            new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
            {
                ["Orders"] = new[]
                {
                    new DocxMailMergeBlockData(
                        new Dictionary<string, string> { ["OrderId"] = "1001" },
                        new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
                        {
                            ["Lines"] = new[] { new DocxMailMergeBlockData(new Dictionary<string, string> { ["Sku"] = "A1" }) }
                        }),
                }
            });

        Assert.Equal("Order: 1001Line: A1", Text(destination.ToArray()));
    }

    [Fact]
    public async Task MergeRepeatingRegionsWithReportAsync_ReportsTheMissingNestedRegion()
    {
        byte[] template = NestedRepeatingTemplate();
        using var source = new MemoryStream(template);
        using var destination = new MemoryStream();

        DocxMailMergeBlockReport report = await DocxMailMerge.MergeRepeatingRegionsWithReportAsync(
            source, destination,
            new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
            {
                ["Orders"] = new[]
                {
                    new DocxMailMergeBlockData(new Dictionary<string, string> { ["OrderId"] = "1001" }),
                }
            });

        Assert.Equal(new[] { "Lines" }, report.MissingNames);
        Assert.Equal("Order: 1001", Text(destination.ToArray()));
    }

    [Fact]
    public void MergeRepeatingRegions_DoesNotMutateTheCallersBlockData()
    {
        // The padding builds a fresh engine-side tree rather than rewriting the caller's rows.
        var lines = new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
        {
            ["Lines"] = new[] { new DocxMailMergeBlockData(new Dictionary<string, string> { ["Sku"] = "A1" }) }
        };
        var order = new DocxMailMergeBlockData(new Dictionary<string, string> { ["OrderId"] = "1001" }, lines);

        DocxMailMerge.MergeRepeatingRegionsWithReport(
            NestedRepeatingTemplate(),
            new Dictionary<string, IEnumerable<DocxMailMergeBlockData>> { ["Orders"] = new[] { order } });

        Assert.Equal(new[] { "Lines" }, order.Regions!.Keys);
        Assert.Single(lines);
    }

    private static byte[] NestedRepeatingTemplate() => Build(body =>
    {
        body.Append(new Paragraph(new Run(new Text("{{#each Orders}}"))));
        var p = new Paragraph();
        p.Append(new Run(new Text("Order: ")));
        p.Append(Field(" MERGEFIELD OrderId \\* MERGEFORMAT ", "«OrderId»"));
        body.Append(p);
        body.Append(new Paragraph(new Run(new Text("{{#each Lines}}"))));
        var p2 = new Paragraph();
        p2.Append(new Run(new Text("Line: ")));
        p2.Append(Field(" MERGEFIELD Sku \\* MERGEFORMAT ", "«Sku»"));
        body.Append(p2);
        body.Append(new Paragraph(new Run(new Text("{{/each Lines}}"))));
        body.Append(new Paragraph(new Run(new Text("{{/each Orders}}"))));
    });

    // ---- A75: MergeTableRows / MergeTableRowGroups --------------------------------------------

    [Fact]
    public void MergeTableRows_ExpandsOnceForEveryRow()
    {
        byte[] template = TableRowsTemplate();

        byte[] merged = DocxMailMerge.MergeTableRows(
            template, tableIndex: 0, templateRowIndex: 1,
            new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["Name"] = "Alice" },
                new Dictionary<string, string> { ["Name"] = "Bob" },
            });

        Assert.Equal("HeaderName: AliceName: Bob", Text(merged));
        AssertValid(merged);
    }

    [Fact]
    public void MergeTableRows_ZeroRows_RemovesTheTemplateRow()
    {
        byte[] template = TableRowsTemplate();

        byte[] merged = DocxMailMerge.MergeTableRows(
            template, tableIndex: 0, templateRowIndex: 1,
            Array.Empty<IReadOnlyDictionary<string, string>>());

        Assert.Equal("Header", Text(merged));
        AssertValid(merged);
    }

    [Fact]
    public void MergeTableRows_OutOfRangeIndex_ThrowsDocumentConversionException()
    {
        byte[] template = TableRowsTemplate();

        Assert.Throws<DocumentConversionException>(() => DocxMailMerge.MergeTableRows(
            template, tableIndex: 0, templateRowIndex: 99,
            new[] { new Dictionary<string, string>() }));
    }

    [Fact]
    public void MergeTableRows_ARecordMissingAField_LeavesItsPlaceholder_WithoutThrowing()
    {
        byte[] template = TableRowsTemplate();

        byte[] merged = DocxMailMerge.MergeTableRows(
            template, tableIndex: 0, templateRowIndex: 1,
            new IReadOnlyDictionary<string, string>[] { new Dictionary<string, string>() });

        Assert.Contains("«Name»", Text(merged));
    }

    [Fact]
    public async Task MergeTableRowsAsync_MatchesTheByteArrayForm()
    {
        byte[] template = TableRowsTemplate();
        using var source = new MemoryStream(template);
        using var destination = new MemoryStream();

        await DocxMailMerge.MergeTableRowsAsync(
            source, destination, tableIndex: 0, templateRowIndex: 1,
            new IReadOnlyDictionary<string, string>[] { new Dictionary<string, string> { ["Name"] = "Alice" } });

        Assert.Equal("HeaderName: Alice", Text(destination.ToArray()));
    }

    [Fact]
    public void MergeTableRowGroups_ExpandsGroupAndDetailRows()
    {
        byte[] template = TableRowGroupsTemplate();

        byte[] merged = DocxMailMerge.MergeTableRowGroups(
            template, tableIndex: 0, groupTemplateRowIndex: 0, detailTemplateRowIndex: 1,
            new[]
            {
                new DocxMailMergeTableRowGroup(
                    new Dictionary<string, string> { ["GroupName"] = "Fruits" },
                    new IReadOnlyDictionary<string, string>[]
                    {
                        new Dictionary<string, string> { ["Item"] = "Apple" },
                        new Dictionary<string, string> { ["Item"] = "Banana" },
                    }),
            });

        Assert.Equal("Group: FruitsDetail: AppleDetail: Banana", Text(merged));
        AssertValid(merged);
    }

    [Fact]
    public async Task MergeTableRowGroupsAsync_MatchesTheByteArrayForm()
    {
        byte[] template = TableRowGroupsTemplate();
        using var source = new MemoryStream(template);
        using var destination = new MemoryStream();

        await DocxMailMerge.MergeTableRowGroupsAsync(
            source, destination, tableIndex: 0, groupTemplateRowIndex: 0, detailTemplateRowIndex: 1,
            new[]
            {
                new DocxMailMergeTableRowGroup(
                    new Dictionary<string, string> { ["GroupName"] = "Fruits" },
                    new IReadOnlyDictionary<string, string>[] { new Dictionary<string, string> { ["Item"] = "Apple" } }),
            });

        Assert.Equal("Group: FruitsDetail: Apple", Text(destination.ToArray()));
    }

    [Fact]
    public void MergeTableRows_ThenMerge_FollowUpPassCatchesAMissingField()
    {
        byte[] template = TableRowsTemplate();

        byte[] expanded = DocxMailMerge.MergeTableRows(
            template, tableIndex: 0, templateRowIndex: 1,
            new IReadOnlyDictionary<string, string>[] { new Dictionary<string, string>() });

        DocxMailMergeResult result = DocxMailMerge.MergeWithReport(expanded, new Dictionary<string, string>());

        Assert.False(result.Report.IsComplete);
        Assert.Equal(new[] { "Name" }, result.Report.MissingFieldNames);
    }

    private static byte[] TableRowsTemplate() => Build(body =>
    {
        var tbl = new Table();
        tbl.Append(new TableProperties());
        tbl.Append(new TableGrid(new GridColumn()));
        tbl.Append(new TableRow(new TableCell(new Paragraph(new Run(new Text("Header"))))));
        var templateRow = new TableRow(new TableCell(new Paragraph(
            new Run(new Text("Name: ")),
            Field(" MERGEFIELD Name \\* MERGEFORMAT ", "«Name»"))));
        tbl.Append(templateRow);
        body.Append(tbl);
    });

    private static byte[] TableRowGroupsTemplate() => Build(body =>
    {
        var tbl = new Table();
        tbl.Append(new TableProperties());
        tbl.Append(new TableGrid(new GridColumn()));
        var groupRow = new TableRow(new TableCell(new Paragraph(
            new Run(new Text("Group: ")),
            Field(" MERGEFIELD GroupName \\* MERGEFORMAT ", "«GroupName»"))));
        tbl.Append(groupRow);
        var detailRow = new TableRow(new TableCell(new Paragraph(
            new Run(new Text("Detail: ")),
            Field(" MERGEFIELD Item \\* MERGEFORMAT ", "«Item»"))));
        tbl.Append(detailRow);
        body.Append(tbl);
    });

    // ---- A75 fix: a strict refusal is scoped to the construct that call guards -----------------

    /// <summary>
    /// The strict forms used to refuse on <c>InspectTemplate</c>'s own <c>IsValid</c>, which is
    /// false whenever the inspection found <b>anything</b> — so one malformed <c>MERGEFIELD</c>, or
    /// a Word-native <c>NEXT</c> control field, closed <see cref="DocxMailMerge.MergeConditional"/>
    /// to a document whose conditional blocks were sound and fully supplied, under a message
    /// announcing a "conditional block issue" and then quoting a merge-field problem.
    /// </summary>
    /// <remarks>
    /// <b>Both fixtures carry a positive control</b>, because the whole test is an assertion that
    /// something does NOT happen: it first confirms <c>InspectTemplate</c> really does mark this
    /// document unsound, and that every issue it names belongs to the OTHER construct. A fixture
    /// that quietly stopped carrying a problem at all would otherwise pass while proving nothing —
    /// which is exactly how the currency-switch guess in this fix's own measurement round turned
    /// out to be a non-issue (<c>\#</c>, <c>\@</c> and <c>\*</c> report nothing; <c>\b</c>,
    /// <c>\f</c> and <c>\v</c> do).
    /// </remarks>
    [Theory]
    [InlineData("malformed", DocxMailMergeIssueKind.MalformedField)]
    [InlineData("control", DocxMailMergeIssueKind.UnsupportedMailMergeControlField)]
    [InlineData("switch", DocxMailMergeIssueKind.UnsupportedFormatting)]
    public void MergeConditional_DoesNotRefuseForAMergeFieldProblemElsewhere(
        string kind, DocxMailMergeIssueKind expectedKind)
    {
        byte[] template = ConditionalTemplateWithAnUnrelatedFieldProblem(kind);

        DocxMailMergeTemplate inspection = DocxMailMerge.InspectTemplate(template);
        Assert.False(inspection.IsValid);
        Assert.NotEmpty(inspection.Issues);
        Assert.All(inspection.Issues, i => Assert.Equal(expectedKind, i.Kind));

        byte[] merged = DocxMailMerge.MergeConditional(
            template, new Dictionary<string, bool> { ["ShowDiscount"] = true });

        Assert.Contains("Discount applies", Text(merged), StringComparison.Ordinal);
        AssertValid(merged);
    }

    [Fact]
    public void MergeRepeating_DoesNotRefuseForAMalformedMergeFieldElsewhere()
    {
        byte[] template = RepeatingTemplateWithAMalformedFieldElsewhere();

        Assert.False(DocxMailMerge.InspectTemplate(template).IsValid);

        byte[] merged = DocxMailMerge.MergeRepeating(
            template,
            new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>
            {
                ["Items"] = new[] { new Dictionary<string, string> { ["Name"] = "Alice" } }
            });

        Assert.Contains("Name: Alice", Text(merged), StringComparison.Ordinal);
        AssertValid(merged);
    }

    [Fact]
    public void MergeRepeatingRegions_DoesNotRefuseForAMalformedMergeFieldElsewhere()
    {
        byte[] template = Build(body =>
        {
            body.Append(new Paragraph(new Run(new Text("{{#each Orders}}"))));
            var p = new Paragraph();
            p.Append(new Run(new Text("Order: ")));
            p.Append(Field(" MERGEFIELD OrderId \\* MERGEFORMAT ", "«OrderId»"));
            body.Append(p);
            body.Append(new Paragraph(new Run(new Text("{{/each Orders}}"))));
            body.Append(new Paragraph(Field(" MERGEFIELD ", "«»")));
        });

        Assert.False(DocxMailMerge.InspectTemplate(template).IsValid);

        byte[] merged = DocxMailMerge.MergeRepeatingRegions(
            template,
            new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
            {
                ["Orders"] = new[]
                {
                    new DocxMailMergeBlockData(new Dictionary<string, string> { ["OrderId"] = "1001" }),
                }
            });

        Assert.Contains("Order: 1001", Text(merged), StringComparison.Ordinal);
        AssertValid(merged);
    }

    [Fact]
    public void MergeConditionalWithReport_StillReportsAnUnrelatedFieldProblem()
    {
        // The other half of the narrowed refusal, and the reason Issues is documented as
        // "everything the inspection found" rather than "structural problems": the strict form
        // stopped refusing for this, so the report must not stop mentioning it.
        byte[] template = ConditionalTemplateWithAnUnrelatedFieldProblem("malformed");

        DocxMailMergeBlockResult result = DocxMailMerge.MergeConditionalWithReport(
            template, new Dictionary<string, bool> { ["ShowDiscount"] = true });

        Assert.Equal(
            DocxMailMergeIssueKind.MalformedField, Assert.Single(result.Report.Issues).Kind);
        Assert.Empty(result.Report.MissingNames);
        Assert.True(result.Report.IsComplete);
    }

    private static byte[] ConditionalTemplateWithAnUnrelatedFieldProblem(string kind) => Build(body =>
    {
        body.Append(new Paragraph(new Run(new Text("Before"))));
        body.Append(new Paragraph(new Run(new Text("{{#ShowDiscount}}"))));
        body.Append(new Paragraph(new Run(new Text("Discount applies"))));
        body.Append(new Paragraph(new Run(new Text("{{/ShowDiscount}}"))));
        body.Append(new Paragraph(UnrelatedFieldProblem(kind)));
        body.Append(new Paragraph(new Run(new Text("After"))));
    });

    /// <summary>
    /// A merge-field problem with nothing to do with conditional blocks or repeating regions —
    /// measured to set <c>InspectTemplate</c>'s <c>IsValid</c> false while reporting no issue of
    /// either construct's kind.
    /// </summary>
    private static SimpleField UnrelatedFieldProblem(string kind) => kind switch
    {
        // No field name at all.
        "malformed" => Field(" MERGEFIELD ", "«»"),
        // Word's own record-control field, which this engine reports and does not execute -- and
        // which a real Word mail-merge template carries routinely.
        "control" => Field(" NEXT ", string.Empty),
        // \b is one of the three switches (\b, \f, \v) measured outside the engine's deterministic
        // formatting profile. \#, \@ and \* are all fine and report nothing.
        _ => Field(" MERGEFIELD Balance \\b \"pre\" ", "«Balance»"),
    };

    private static byte[] RepeatingTemplateWithAMalformedFieldElsewhere() => Build(body =>
    {
        body.Append(new Paragraph(new Run(new Text("Before"))));
        body.Append(new Paragraph(new Run(new Text("{{#each Items}}"))));
        var p = new Paragraph();
        p.Append(new Run(new Text("Name: ")));
        p.Append(Field(" MERGEFIELD Name \\* MERGEFORMAT ", "«Name»"));
        body.Append(p);
        body.Append(new Paragraph(new Run(new Text("{{/each Items}}"))));
        body.Append(new Paragraph(Field(" MERGEFIELD ", "«»")));
        body.Append(new Paragraph(new Run(new Text("After"))));
    });

    // ---- A75 fix: the null-value guarantee reaches the per-record collections too --------------

    /// <summary>
    /// The class states as a guarantee that a null VALUE is refused rather than merged, because the
    /// engine writes it as an empty string and reports the document complete. Four of the five new
    /// methods never checked, so a database NULL merged into a generated row and — with
    /// <c>removeFields: true</c> — left nothing for a follow-up pass to find. Measured before the
    /// fix: <c>MergeRepeating</c> with a null <c>Name</c> produced "Before / Name: / After".
    /// </summary>
    [Fact]
    public void MergeRepeating_RefusesANullValueInARecord_NamingWhichRecord()
    {
        var ex = Assert.Throws<ArgumentException>(() => DocxMailMerge.MergeRepeating(
            RepeatingTemplate("Items", "Name"),
            new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>
            {
                ["Items"] = new[]
                {
                    new Dictionary<string, string> { ["Name"] = "Alice" },
                    new Dictionary<string, string> { ["Name"] = null! },
                }
            }));

        Assert.Equal("regions", ex.ParamName);
        Assert.Contains("Region 'Items' record 1", ex.Message, StringComparison.Ordinal);
        Assert.Contains("'Name' is null", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MergeRepeatingRegions_RefusesANullValueInANESTEDRow_NamingThePathToIt()
    {
        // Nesting is the case a top-level-only check would miss, so the null goes in the inner row
        // and the message has to name both levels to be worth anything.
        var ex = Assert.Throws<ArgumentException>(() => DocxMailMerge.MergeRepeatingRegions(
            NestedRepeatingTemplate(),
            new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
            {
                ["Orders"] = new[]
                {
                    new DocxMailMergeBlockData(
                        new Dictionary<string, string> { ["OrderId"] = "1001" },
                        new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
                        {
                            ["Lines"] = new[]
                            {
                                new DocxMailMergeBlockData(new Dictionary<string, string> { ["Sku"] = "A1" }),
                                new DocxMailMergeBlockData(new Dictionary<string, string> { ["Sku"] = null! }),
                            }
                        }),
                }
            }));

        Assert.Equal("regions", ex.ParamName);
        Assert.Contains(
            "Region 'Orders' record 0 -> region 'Lines' record 1", ex.Message, StringComparison.Ordinal);
        Assert.Contains("'Sku' is null", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MergeTableRows_RefusesANullValueInARow_NamingWhichRow()
    {
        var ex = Assert.Throws<ArgumentException>(() => DocxMailMerge.MergeTableRows(
            TableRowsTemplate(), tableIndex: 0, templateRowIndex: 1,
            new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["Name"] = "Alice" },
                new Dictionary<string, string> { ["Name"] = null! },
            }));

        Assert.Equal("rows", ex.ParamName);
        Assert.Contains("Row 1", ex.Message, StringComparison.Ordinal);
        Assert.Contains("'Name' is null", ex.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void MergeTableRowGroups_RefusesANullValue_InEitherTheGroupOrADetailRow(bool inTheGroup)
    {
        var group = inTheGroup
            ? new DocxMailMergeTableRowGroup(
                new Dictionary<string, string> { ["GroupName"] = null! },
                new IReadOnlyDictionary<string, string>[]
                {
                    new Dictionary<string, string> { ["Item"] = "Apple" },
                })
            : new DocxMailMergeTableRowGroup(
                new Dictionary<string, string> { ["GroupName"] = "Fruits" },
                new IReadOnlyDictionary<string, string>[]
                {
                    new Dictionary<string, string> { ["Item"] = "Apple" },
                    new Dictionary<string, string> { ["Item"] = null! },
                });

        var ex = Assert.Throws<ArgumentException>(() => DocxMailMerge.MergeTableRowGroups(
            TableRowGroupsTemplate(), tableIndex: 0, groupTemplateRowIndex: 0,
            detailTemplateRowIndex: 1, new[] { group }));

        Assert.Equal("groups", ex.ParamName);
        Assert.Contains(inTheGroup ? "Group 0:" : "Group 0 row 1", ex.Message, StringComparison.Ordinal);
        Assert.Contains(
            inTheGroup ? "'GroupName' is null" : "'Item' is null", ex.Message, StringComparison.Ordinal);
    }

    // ---- A75 fix: the caller's region tree is read once ----------------------------------------

    /// <summary>
    /// <c>MergeRepeatingRegionsCore</c> walked the caller's sequences twice — once to collect the
    /// supplied names, once to build what the engine takes. A genuinely single-pass source is empty
    /// on the second walk, so the document came out valid, complete-looking and missing every row,
    /// with no exception and nothing in the report.
    /// </summary>
    /// <remarks>
    /// <b>Measured both ways.</b> With the materialisation removed, this fixture produced an empty
    /// document ("" extracted text) after two walks; with it, one walk and both orders present.
    /// <see cref="SinglePass{T}"/> yields nothing rather than throwing on its later walks, on
    /// purpose — a source that threw would have made the bug loud, and the bug was silent.
    /// </remarks>
    [Fact]
    public void MergeRepeatingRegions_ReadsTheCallersSequenceOnce_NotOncePerInternalWalk()
    {
        var orders = new SinglePass<DocxMailMergeBlockData>(
        [
            new DocxMailMergeBlockData(new Dictionary<string, string> { ["OrderId"] = "1001" }),
            new DocxMailMergeBlockData(new Dictionary<string, string> { ["OrderId"] = "1002" }),
        ]);

        byte[] merged = DocxMailMerge.MergeRepeatingRegions(
            FlatOrdersTemplate(),
            new Dictionary<string, IEnumerable<DocxMailMergeBlockData>> { ["Orders"] = orders });

        // The content assertion is what the caller cares about; the walk count is what says WHY it
        // holds, so a future change that reintroduces a second walk over a materialised copy does
        // not silently re-open the door for the caller's own sequence.
        Assert.Equal("Order: 1001Order: 1002", Text(merged));
        Assert.Equal(1, orders.Walks);
        AssertValid(merged);
    }

    [Fact]
    public void MergeRepeatingRegions_ReadsANestedSequenceOnceToo()
    {
        var lines = new SinglePass<DocxMailMergeBlockData>(
        [
            new DocxMailMergeBlockData(new Dictionary<string, string> { ["Sku"] = "A1" }),
            new DocxMailMergeBlockData(new Dictionary<string, string> { ["Sku"] = "A2" }),
        ]);

        byte[] merged = DocxMailMerge.MergeRepeatingRegions(
            NestedRepeatingTemplate(),
            new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
            {
                ["Orders"] = new[]
                {
                    new DocxMailMergeBlockData(
                        new Dictionary<string, string> { ["OrderId"] = "1001" },
                        new Dictionary<string, IEnumerable<DocxMailMergeBlockData>> { ["Lines"] = lines }),
                }
            });

        Assert.Equal("Order: 1001Line: A1Line: A2", Text(merged));
        Assert.Equal(1, lines.Walks);
    }

    private static byte[] FlatOrdersTemplate() => Build(body =>
    {
        body.Append(new Paragraph(new Run(new Text("{{#each Orders}}"))));
        var p = new Paragraph();
        p.Append(new Run(new Text("Order: ")));
        p.Append(Field(" MERGEFIELD OrderId \\* MERGEFORMAT ", "«OrderId»"));
        body.Append(p);
        body.Append(new Paragraph(new Run(new Text("{{/each Orders}}"))));
    });

    // ---- A75 fix: MergeTableRows' table index is NOT DocxEditor.ReadTable's -------------------

    /// <summary>
    /// The doc comment used to say <c>tableIndex</c> selected the same table
    /// <see cref="DocxEditor.ReadTable(byte[], int)"/> would. Measured, it does not: <c>ReadTable</c>
    /// descends into a block-level content control (<c>w:sdt</c>) and the underlying mail-merge
    /// engine's own <c>Tables</c> collection does not, so a control-wrapped table is one of
    /// <c>ReadTable</c>'s and none of this method's.
    /// </summary>
    /// <remarks>
    /// The doc comment was corrected rather than the selection reimplemented — matching
    /// <c>ReadTable</c>'s content-control-aware walk would mean replacing the engine's own table
    /// lookup, which is a far larger change than the claim was worth. Both indexes are asserted
    /// here so the corrected claim is the one that is pinned.
    /// </remarks>
    [Fact]
    public void MergeTableRows_TableIndexDisagreesWithReadTable_ForAContentControlWrappedTable()
    {
        byte[] docx = WrappedAndOrdinaryTables();

        // The fixture is what it claims to be -- this repository has had a family of measurements
        // invalidated by a hand-built w:sdt that was not a content control at all.
        AssertValid(docx);

        // ReadTable sees two tables, the wrapped one first.
        Assert.Equal(2, DocxEditor.TableCount(docx));
        Assert.Equal("WRAPPED", DocxEditor.ReadTable(docx, 0)[0][0]);
        Assert.Equal("Header", DocxEditor.ReadTable(docx, 1)[0][0]);

        var rows = new IReadOnlyDictionary<string, string>[]
        {
            new Dictionary<string, string> { ["Name"] = "Alice" },
            new Dictionary<string, string> { ["Name"] = "Bob" },
        };

        // MergeTableRows' index 0 is the ORDINARY table -- ReadTable's index 1.
        byte[] merged = DocxMailMerge.MergeTableRows(docx, tableIndex: 0, templateRowIndex: 1, rows);
        Assert.Equal("WRAPPEDHeaderName: AliceName: Bobafter", Text(merged));
        AssertValid(merged);

        // ...and index 1, which ReadTable answers, is out of range here. Without this leg the
        // assertion above would also pass against an implementation that happened to agree.
        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxMailMerge.MergeTableRows(docx, tableIndex: 1, templateRowIndex: 1, rows));
        Assert.IsType<ArgumentOutOfRangeException>(ex.InnerException);
    }

    private static byte[] WrappedAndOrdinaryTables() => Build(body =>
    {
        var wrapped = new Table();
        wrapped.Append(new TableProperties());
        wrapped.Append(new TableGrid(new GridColumn()));
        wrapped.Append(new TableRow(new TableCell(new Paragraph(new Run(new Text("WRAPPED"))))));
        body.Append(new SdtBlock(
            new SdtProperties(new SdtAlias { Val = "c" }, new Tag { Val = "c" }),
            new SdtContentBlock(wrapped)));

        var ordinary = new Table();
        ordinary.Append(new TableProperties());
        ordinary.Append(new TableGrid(new GridColumn()));
        ordinary.Append(new TableRow(new TableCell(new Paragraph(new Run(new Text("Header"))))));
        ordinary.Append(new TableRow(new TableCell(new Paragraph(
            new Run(new Text("Name: ")),
            Field(" MERGEFIELD Name \\* MERGEFORMAT ", "«Name»")))));
        body.Append(ordinary);

        // A sibling paragraph, so "the table is gone" and "the read came back empty" cannot be the
        // same observation.
        body.Append(new Paragraph(new Run(new Text("after"))));
    });

    // ---- A75 fix: MergeTableRowGroups' missing test legs ---------------------------------------

    [Fact]
    public void MergeTableRowGroups_ZeroGroups_RemovesBothTemplateRows()
    {
        byte[] merged = DocxMailMerge.MergeTableRowGroups(
            TableRowGroupsTemplate(), tableIndex: 0, groupTemplateRowIndex: 0,
            detailTemplateRowIndex: 1, Array.Empty<DocxMailMergeTableRowGroup>());

        Assert.Equal(string.Empty, Text(merged));
        AssertValid(merged);
    }

    [Theory]
    [InlineData(99, 1)]
    [InlineData(0, 99)]
    public void MergeTableRowGroups_OutOfRangeRowIndex_ThrowsDocumentConversionException(
        int groupTemplateRowIndex, int detailTemplateRowIndex)
    {
        // Both row axes, because this method has two and MergeTableRows' own out-of-range test
        // cannot say anything about the second one.
        var ex = Assert.Throws<DocumentConversionException>(() => DocxMailMerge.MergeTableRowGroups(
            TableRowGroupsTemplate(), tableIndex: 0, groupTemplateRowIndex, detailTemplateRowIndex,
            new[]
            {
                new DocxMailMergeTableRowGroup(
                    new Dictionary<string, string> { ["GroupName"] = "Fruits" },
                    new IReadOnlyDictionary<string, string>[]
                    {
                        new Dictionary<string, string> { ["Item"] = "Apple" },
                    }),
            }));

        Assert.IsType<ArgumentOutOfRangeException>(ex.InnerException);
    }

    [Fact]
    public void MergeTableRowGroups_ARecordMissingAField_LeavesItsPlaceholder_WithoutThrowing()
    {
        byte[] merged = DocxMailMerge.MergeTableRowGroups(
            TableRowGroupsTemplate(), tableIndex: 0, groupTemplateRowIndex: 0,
            detailTemplateRowIndex: 1,
            new[]
            {
                new DocxMailMergeTableRowGroup(
                    new Dictionary<string, string> { ["GroupName"] = "Fruits" },
                    new IReadOnlyDictionary<string, string>[] { new Dictionary<string, string>() }),
            });

        // Same silent behaviour as MergeTableRows, and caught the same way -- by a follow-up
        // field-level pass, which is what the doc comment tells a caller to do.
        Assert.Equal("Group: FruitsDetail: «Item»", Text(merged));

        DocxMailMergeResult followUp =
            DocxMailMerge.MergeWithReport(merged, new Dictionary<string, string>());
        Assert.Equal(new[] { "Item" }, followUp.Report.MissingFieldNames);
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
            new Dictionary<string, string> { ["FirstName"] = "Alice" },
            new Dictionary<string, string> { ["FirstName"] = null! },
        };

        var ex = Assert.Throws<ArgumentException>(
            () => DocxMailMerge.MergeBatch(Simple("FirstName"), records).ToList());

        Assert.Equal("records", ex.ParamName);
        // A single-record batch can't tell $"Record {index}" apart from a hard-coded "Record 0" --
        // record 1 is what actually pins the index to the loop variable, not a constant.
        Assert.Contains("Record 1:", ex.Message, StringComparison.Ordinal);
        Assert.Contains("FirstName", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MergeBatchToFiles_ARecordWithANullValue_NamesTheRecordAndTheRecordsParameter()
    {
        var dir = Directory.CreateTempSubdirectory("DocxMailMergeTests-");
        try
        {
            var templatePath = Path.Join(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName"));
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice" },
                new Dictionary<string, string> { ["FirstName"] = null! },
            };

            var ex = Assert.Throws<ArgumentException>(() =>
                DocxMailMerge.MergeBatchToFiles(templatePath, records,
                    (i, r) => Path.Join(dir.FullName, $"out-{i}.docx")));

            Assert.Equal("records", ex.ParamName);
            // A single-record batch can't tell $"Record {index}" apart from a hard-coded "Record 0"
            // -- record 1 is what actually pins the index to the loop variable, not a constant.
            Assert.Contains("Record 1:", ex.Message, StringComparison.Ordinal);
            Assert.Contains("FirstName", ex.Message, StringComparison.Ordinal);
            // RequireValues runs in the path-computation loop, before CheckNoPathCollisions and
            // before any write -- a null value anywhere in the batch means no file is written at
            // all, unlike a strict field-miss where records before the bad one are already on disk.
            Assert.False(File.Exists(Path.Join(dir.FullName, "out-0.docx")));
        }
        finally
        {
            dir.Delete(recursive: true);
        }
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
        // "1" alone would also match the unrelated "1 merge field(s)" substring MergeCore's own
        // message already contains -- "Record 1:" is what actually pins the index.
        Assert.Contains("Record 1:", ex.Message, StringComparison.Ordinal);
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
            var templatePath = Path.Join(dir.FullName, "template.docx");
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
                return Path.Join(dir.FullName, $"out-{i}.docx");
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
            var templatePath = Path.Join(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName", "Balance"));
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice", ["Balance"] = "100" },
                new Dictionary<string, string> { ["FirstName"] = "Bob" }, // Balance missing
                new Dictionary<string, string> { ["FirstName"] = "Carol", ["Balance"] = "300" },
            };

            var ex = Assert.Throws<DocumentConversionException>(() =>
                DocxMailMerge.MergeBatchToFiles(templatePath, records,
                    (i, r) => Path.Join(dir.FullName, $"out-{i}.docx")));

            // "1" alone would also match the unrelated "1 merge field(s)" substring MergeCore's own
            // message already contains -- "Record 1:" is what actually pins the index.
            Assert.Contains("Record 1:", ex.Message, StringComparison.Ordinal);
            Assert.Contains("Balance", ex.Message, StringComparison.Ordinal);
            Assert.True(File.Exists(Path.Join(dir.FullName, "out-0.docx")));
            Assert.False(File.Exists(Path.Join(dir.FullName, "out-1.docx")));
            Assert.False(File.Exists(Path.Join(dir.FullName, "out-2.docx")));
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
            var templatePath = Path.Join(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName"));
            var collidingPath = Path.Join(dir.FullName, "collide.docx");
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice" },
                new Dictionary<string, string> { ["FirstName"] = "Bob" },
                new Dictionary<string, string> { ["FirstName"] = "Carol" },
            };

            // Records 0 and 2 collide; record 1 does not.
            var ex = Assert.Throws<ArgumentException>(() =>
                DocxMailMerge.MergeBatchToFiles(templatePath, records,
                    (i, r) => i == 1 ? Path.Join(dir.FullName, "unique.docx") : collidingPath));

            Assert.Equal("outputPathFactory", ex.ParamName);
            // "Records 0 and 2", not bare "0"/"2" -- the temp directory's own generated name can
            // easily contain those digits as a substring, which would make a bare-digit assertion
            // pass whether or not the message actually names the right records.
            Assert.Contains("Records 0 and 2", ex.Message, StringComparison.Ordinal);
            // Nothing was written at all -- not even the record that never collided, and not even
            // the "winning" one an unguarded delegation to the engine would silently have produced.
            Assert.False(File.Exists(collidingPath));
            Assert.False(File.Exists(Path.Join(dir.FullName, "unique.docx")));
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
            var templatePath = Path.Join(dir.FullName, "template.docx");
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
                    (i, r) => { calledIndices.Add(i); return Path.Join(dir.FullName, r["FirstName"] + ".docx"); }));

            Assert.Equal("records", ex.ParamName);
            Assert.Equal([0], calledIndices);
            // Nothing was written -- not even record 0, whose path was computed successfully
            // before the null record was ever reached.
            Assert.False(File.Exists(Path.Join(dir.FullName, "Alice.docx")));
            Assert.False(File.Exists(Path.Join(dir.FullName, "Carol.docx")));
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
            var templatePath = Path.Join(dir.FullName, "template.docx");
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
                    (i, r) => i == 0 ? Path.Join(dir.FullName, "out-0.docx") : null!));
            Assert.Equal("outputPathFactory", nullEx.ParamName);
            Assert.Contains("Record 1:", nullEx.Message, StringComparison.Ordinal);
            // Proves the batch is refused before any write, not mid-batch -- record 0's path was
            // already computed (and would have been valid) by the time record 1's bad path is found.
            Assert.False(File.Exists(Path.Join(dir.FullName, "out-0.docx")));

            var blankEx = Assert.Throws<ArgumentException>(() =>
                DocxMailMerge.MergeBatchToFiles(templatePath, records,
                    (i, r) => i == 0 ? Path.Join(dir.FullName, "out-0.docx") : "   "));
            Assert.Equal("outputPathFactory", blankEx.ParamName);
            Assert.Contains("Record 1:", blankEx.Message, StringComparison.Ordinal);
            Assert.False(File.Exists(Path.Join(dir.FullName, "out-0.docx")));
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
            var templatePath = Path.Join(dir.FullName, "template.docx");
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
            var templatePath = Path.Join(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName", "Balance"));
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice", ["Balance"] = "100" },
                new Dictionary<string, string> { ["FirstName"] = "Bob", ["Balance"] = "250" },
            };

            var syncPaths = DocxMailMerge.MergeBatchToFiles(templatePath, records,
                (i, r) => Path.Join(dir.FullName, $"sync-{i}.docx"));
            var asyncPaths = await DocxMailMerge.MergeBatchToFilesAsync(templatePath, records,
                (i, r) => Path.Join(dir.FullName, $"async-{i}.docx"));

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
            var templatePath = Path.Join(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName"));
            var paths = new[]
            {
                Path.Join(dir.FullName, "out-0.docx"),
                Path.Join(dir.FullName, "out-1.docx"),
                Path.Join(dir.FullName, "out-2.docx"),
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
            var templatePath = Path.Join(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName", "Balance"));
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice", ["Balance"] = "100" },
                new Dictionary<string, string> { ["FirstName"] = "Bob" }, // Balance missing
                new Dictionary<string, string> { ["FirstName"] = "Carol", ["Balance"] = "300" },
            };

            var items = DocxMailMerge.MergeBatchToFilesWithReport(templatePath, records,
                (i, r) => Path.Join(dir.FullName, $"out-{i}.docx"));

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
            var templatePath = Path.Join(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName"));
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice" },
                new Dictionary<string, string> { ["FirstName"] = "Bob" },
            };

            var ex = Assert.Throws<ArgumentException>(() =>
                DocxMailMerge.MergeBatchToFilesWithReport(templatePath, records,
                    (i, r) => Path.Join(dir.FullName, "collide.docx")));

            Assert.Equal("outputPathFactory", ex.ParamName);
            Assert.False(File.Exists(Path.Join(dir.FullName, "collide.docx")));
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
            var templatePath = Path.Join(dir.FullName, "template.docx");
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
                    (i, r) => { calledIndices.Add(i); return Path.Join(dir.FullName, r["FirstName"] + ".docx"); }));

            Assert.Equal("records", ex.ParamName);
            Assert.Equal([0], calledIndices);
            Assert.False(File.Exists(Path.Join(dir.FullName, "Alice.docx")));
            Assert.False(File.Exists(Path.Join(dir.FullName, "Carol.docx")));
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
            var templatePath = Path.Join(dir.FullName, "template.docx");
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
                    (i, r) => i == 0 ? Path.Join(dir.FullName, "out-0.docx") : null!));
            Assert.Equal("outputPathFactory", nullEx.ParamName);
            Assert.False(File.Exists(Path.Join(dir.FullName, "out-0.docx")));

            var blankEx = Assert.Throws<ArgumentException>(() =>
                DocxMailMerge.MergeBatchToFilesWithReport(templatePath, records,
                    (i, r) => i == 0 ? Path.Join(dir.FullName, "out-0.docx") : "   "));
            Assert.Equal("outputPathFactory", blankEx.ParamName);
            Assert.False(File.Exists(Path.Join(dir.FullName, "out-0.docx")));
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
            var templatePath = Path.Join(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName", "Balance"));
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice", ["Balance"] = "100" },
                new Dictionary<string, string> { ["FirstName"] = "Bob" }, // Balance missing
            };

            var syncItems = DocxMailMerge.MergeBatchToFilesWithReport(templatePath, records,
                (i, r) => Path.Join(dir.FullName, $"sync-{i}.docx"));
            var asyncItems = await DocxMailMerge.MergeBatchToFilesWithReportAsync(templatePath, records,
                (i, r) => Path.Join(dir.FullName, $"async-{i}.docx"));

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
            var templatePath = Path.Join(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName"));
            var records = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["FirstName"] = "Alice" },
                new Dictionary<string, string>(), // FirstName missing
                new Dictionary<string, string> { ["FirstName"] = "Carol" },
            };

            var items = await DocxMailMerge.MergeBatchToFilesWithReportAsync(templatePath, records,
                (i, r) => Path.Join(dir.FullName, $"out-{i}.docx"));

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
            var templatePath = Path.Join(dir.FullName, "template.docx");
            File.WriteAllBytes(templatePath, Simple("FirstName"));
            var paths = new[]
            {
                Path.Join(dir.FullName, "out-0.docx"),
                Path.Join(dir.FullName, "out-1.docx"),
                Path.Join(dir.FullName, "out-2.docx"),
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

    // ---- A75: composition across constructs, in the required order ---------------------------

    [Fact]
    public void MergeRepeating_ThenMergeConditional_ComposesInTheRequiredOrder()
    {
        // WHICH CONDITIONAL MARKERS A DOCUMENT CONTAINS IS DECIDED BY THE REPEATING PASS, which is
        // what makes the order a requirement rather than a preference. `Notices` receives no
        // records, so expanding it removes the {{#Urgent}} pair that only ever existed inside its
        // region template -- and a caller with no notices has no urgency to declare, so the
        // conditions dictionary is legitimately empty. Run the conditional pass on the UNEXPANDED
        // template and it sees a marker the caller never supplied, and refuses.
        //
        // The measurement that produced this fixture is worth stating, because the obvious fixture
        // does NOT discriminate: with the condition SUPPLIED, a conditional block nested inside a
        // region -- or one wrapping a region -- resolves to byte-identical output in either order,
        // measured both ways at true and at false. A global boolean applied uniformly commutes with
        // duplication. Only a fixture where one pass changes what the OTHER pass can see can tell
        // the two orders apart, and this is the smallest one that does.
        byte[] template = Build(body =>
        {
            body.Append(new Paragraph(new Run(new Text("Catalogue"))));
            body.Append(new Paragraph(new Run(new Text("{{#each Items}}"))));
            var p = new Paragraph();
            p.Append(Field(" MERGEFIELD Name \\* MERGEFORMAT ", "«Name»"));
            body.Append(p);
            body.Append(new Paragraph(new Run(new Text("{{/each Items}}"))));
            body.Append(new Paragraph(new Run(new Text("{{#each Notices}}"))));
            body.Append(new Paragraph(new Run(new Text("{{#Urgent}}"))));
            body.Append(new Paragraph(new Run(new Text("URGENT"))));
            body.Append(new Paragraph(new Run(new Text("{{/Urgent}}"))));
            body.Append(new Paragraph(new Run(new Text("{{/each Notices}}"))));
            body.Append(new Paragraph(new Run(new Text("End"))));
        });

        var regions = new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>
        {
            ["Items"] = new IReadOnlyDictionary<string, string>[]
            {
                new Dictionary<string, string> { ["Name"] = "Alice" },
                new Dictionary<string, string> { ["Name"] = "Bob" },
            },
            ["Notices"] = Array.Empty<IReadOnlyDictionary<string, string>>(),
        };
        var noConditions = new Dictionary<string, bool>();

        // The required order: two records genuinely multiply, and the empty region takes its
        // conditional markers with it.
        byte[] required = DocxMailMerge.MergeConditional(
            DocxMailMerge.MergeRepeating(template, regions), noConditions);

        Assert.Equal("CatalogueAliceBobEnd", Text(required));

        // The same three arguments, the other way round -- and the INNER call is the one that
        // refuses, so the repeating pass that would have come after it never runs at all.
        //
        // Nesting the two calls inside one Assert.Throws could not say that: it passes whichever
        // of them throws, so it would hold just as well against a conditional pass that succeeded
        // and a repeating pass that failed for some unrelated reason. Measured, the conditional
        // pass refuses with "Conditional block 'Urgent' was not supplied."
        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxMailMerge.MergeConditional(template, noConditions));

        Assert.Contains("Urgent", ex.Message, StringComparison.Ordinal);
        Assert.Contains("conditional block issue", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MergeTableRows_ThenMergeConditional_ThenMerge_ComposesAcrossAllThreeLayers()
    {
        // Both boundaries of the three-layer order are load-bearing here, and each fails a
        // different way when crossed:
        //
        //   rows BEFORE conditional -- MergeTableRows selects a table by INDEX, and the hidden
        //   block owns table 0. Resolving the conditional first deletes it, so the template table
        //   stops being table 1. Position is not a name; a structural pass cannot run after
        //   something that removes structure.
        //
        //   conditional BEFORE merge -- «LegacyText» lives inside the hidden block and is
        //   deliberately never supplied, because a caller hiding a section has no values for it.
        //   The field-level Merge is strict, so reaching that field before the conditional pass
        //   deletes it refuses the whole document.
        byte[] template = Build(body =>
        {
            body.Append(new Paragraph(new Run(new Text("Report"))));
            body.Append(new Paragraph(new Run(new Text("{{#ShowLegacy}}"))));

            var legacy = new Table();
            legacy.Append(new TableProperties());
            legacy.Append(new TableGrid(new GridColumn()));
            legacy.Append(new TableRow(new TableCell(new Paragraph(new Run(new Text("LEGACY"))))));
            body.Append(legacy);

            var note = new Paragraph();
            note.Append(new Run(new Text("Legacy note: ")));
            note.Append(Field(" MERGEFIELD LegacyText \\* MERGEFORMAT ", "«LegacyText»"));
            body.Append(note);
            body.Append(new Paragraph(new Run(new Text("{{/ShowLegacy}}"))));

            var current = new Table();
            current.Append(new TableProperties());
            current.Append(new TableGrid(new GridColumn()));
            current.Append(new TableRow(new TableCell(new Paragraph(
                Field(" MERGEFIELD Name \\* MERGEFORMAT ", "«Name»")))));
            body.Append(current);

            var sign = new Paragraph();
            sign.Append(new Run(new Text("Sincerely, ")));
            sign.Append(Field(" MERGEFIELD Sender \\* MERGEFORMAT ", "«Sender»"));
            body.Append(sign);
        });

        var rows = new IReadOnlyDictionary<string, string>[]
        {
            new Dictionary<string, string> { ["Name"] = "Alice" },
            new Dictionary<string, string> { ["Name"] = "Bob" },
        };
        var conditions = new Dictionary<string, bool> { ["ShowLegacy"] = false };
        var values = new Dictionary<string, string> { ["Sender"] = "Khoa" };

        // The required order. Two records, so the single template row genuinely multiplies.
        byte[] required = DocxMailMerge.Merge(
            DocxMailMerge.MergeConditional(
                DocxMailMerge.MergeTableRows(template, tableIndex: 1, templateRowIndex: 0, rows),
                conditions),
            values);

        Assert.Equal("ReportAliceBobSincerely, Khoa", Text(required));

        // Conditional first: table 1 no longer exists.
        var shifted = Assert.Throws<DocumentConversionException>(
            () => DocxMailMerge.MergeTableRows(
                DocxMailMerge.MergeConditional(template, conditions),
                tableIndex: 1, templateRowIndex: 0, rows));

        Assert.IsType<ArgumentOutOfRangeException>(shifted.InnerException);

        // ...and the template table is still there, one index lower -- so what the wrong order
        // destroyed is the meaning of the index, not the table. Without this control the
        // assertion above would pass just as happily against a conditional pass that deleted
        // everything, which would say nothing about ordering.
        Assert.Equal(
            "ReportAliceBobSincerely, Khoa",
            Text(DocxMailMerge.Merge(
                DocxMailMerge.MergeTableRows(
                    DocxMailMerge.MergeConditional(template, conditions),
                    tableIndex: 0, templateRowIndex: 0, rows),
                values)));

        // Merge before the conditional pass, with the rows already expanded so only the third
        // layer is out of place.
        var early = Assert.Throws<DocumentConversionException>(
            () => DocxMailMerge.MergeConditional(
                DocxMailMerge.Merge(
                    DocxMailMerge.MergeTableRows(template, tableIndex: 1, templateRowIndex: 0, rows),
                    values),
                conditions));

        Assert.Contains("LegacyText", early.Message, StringComparison.Ordinal);
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

    /// <summary>
    /// Yields its items on the FIRST walk and nothing on any later one, counting walks — what an
    /// iterator over a forward-only source such as a <c>DbDataReader</c> does, and deliberately
    /// <b>without throwing to say so</b>.
    /// </summary>
    /// <remarks>
    /// A double-enumeration guard that threw would be a strictly weaker fixture: the bug this
    /// stands in for produced a valid, complete-looking, empty document, and a throwing source
    /// would have turned it into an exception nobody could miss. Yielding nothing reproduces the
    /// silence, and <see cref="Walks"/> is what turns "the content is right" into "the content is
    /// right for the right reason".
    /// </remarks>
    private sealed class SinglePass<T> : IEnumerable<T>
    {
        private readonly IReadOnlyList<T> _items;
        private bool _walked;

        public SinglePass(IReadOnlyList<T> items) => _items = items;

        /// <summary>How many times something asked this sequence for an enumerator.</summary>
        public int Walks { get; private set; }

        public IEnumerator<T> GetEnumerator()
        {
            Walks++;
            if (_walked)
                return Enumerable.Empty<T>().GetEnumerator();

            _walked = true;
            return _items.GetEnumerator();
        }

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
