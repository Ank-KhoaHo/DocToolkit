using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

/// <summary>
/// None of the fixtures below carry a real Word <c>MERGEFIELD</c> - none of the twenty methods
/// covered here fill one themselves (that is <c>Merge</c>/<c>MergeWithReport</c>'s own job, per
/// this class's own doc comments: "a missing field inside one record's expansion is not caught
/// here"). What each test proves is that the SERVICE calls the SAME static method with the SAME
/// arguments in the SAME order as a direct static call would - parity plus a structural
/// discriminator (block/row/document COUNT, or which named block survived) that a swapped
/// argument or a wrong static method would visibly break. Full field-substitution semantics are
/// already exhaustively covered by the core project's own DocxMailMergeTests.cs.
/// </summary>
public class DocxMailMergeServiceTests
{
    // ---------------------------------------------------------------------------------------
    // MergeConditional family, mirrored from core 0.43.0 (A87-DI).
    // ---------------------------------------------------------------------------------------

    private static byte[] ConditionalTemplate() => DocxEditor.Create([
        DocxBlock.Paragraph("Start."),
        DocxBlock.Paragraph("{{#Bonus}}"),
        DocxBlock.Paragraph("Bonus text."),
        DocxBlock.Paragraph("{{/Bonus}}"),
        DocxBlock.Paragraph("End."),
    ]);

    [Fact]
    public void MergeConditional_MatchesTheStaticMethod()
    {
        var sut = new DocxMailMergeService();
        var docx = ConditionalTemplate();
        var conditions = new Dictionary<string, bool> { ["Bonus"] = true };

        var fromWrapper = sut.MergeConditional(docx, conditions);

        Assert.Equal(
            DocxEditor.ExtractText(DocxMailMerge.MergeConditional(docx, conditions)),
            DocxEditor.ExtractText(fromWrapper));
        Assert.Contains("Bonus text.", DocxEditor.ExtractText(fromWrapper));
    }

    [Fact]
    public async Task MergeConditionalAsync_MatchesTheStaticMethod()
    {
        var sut = new DocxMailMergeService();
        var docx = ConditionalTemplate();
        var conditions = new Dictionary<string, bool> { ["Bonus"] = false };

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        await sut.MergeConditionalAsync(source, destination, conditions);

        var text = DocxEditor.ExtractText(destination.ToArray());
        Assert.DoesNotContain("Bonus text.", text);
        Assert.Contains("Start.", text);
        Assert.Contains("End.", text);
    }

    [Fact]
    public void MergeConditionalWithReport_MatchesTheStaticMethodAndReportsWhatWasMissing()
    {
        var sut = new DocxMailMergeService();
        var docx = ConditionalTemplate();
        var empty = new Dictionary<string, bool>();

        var fromWrapper = sut.MergeConditionalWithReport(docx, empty);

        Assert.Equal(
            DocxMailMerge.MergeConditionalWithReport(docx, empty).Report.MissingNames,
            fromWrapper.Report.MissingNames);
        Assert.Equal(new[] { "Bonus" }, fromWrapper.Report.MissingNames);
        Assert.DoesNotContain("Bonus text.", DocxEditor.ExtractText(fromWrapper.Document));
    }

    [Fact]
    public async Task MergeConditionalWithReportAsync_MatchesTheStaticMethod()
    {
        var sut = new DocxMailMergeService();
        var docx = ConditionalTemplate();
        var empty = new Dictionary<string, bool>();

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        var report = await sut.MergeConditionalWithReportAsync(source, destination, empty);

        Assert.Equal(new[] { "Bonus" }, report.MissingNames);
    }

    // ---------------------------------------------------------------------------------------
    // MergeRepeating family, mirrored from core 0.43.0 (A87-DI). Discriminator is the repeated
    // paragraph's COUNT, since a swapped argument (wrong region dictionary, wrong strict flag)
    // changes how many times - or whether at all - the block expands.
    // ---------------------------------------------------------------------------------------

    private static byte[] RepeatingTemplate() => DocxEditor.Create([
        DocxBlock.Paragraph("{{#each Items}}"),
        DocxBlock.Paragraph("Row."),
        DocxBlock.Paragraph("{{/each Items}}"),
    ]);

    private static IReadOnlyDictionary<string, string>[] ThreeEmptyRecords() =>
    [
        new Dictionary<string, string>(),
        new Dictionary<string, string>(),
        new Dictionary<string, string>(),
    ];

    [Fact]
    public void MergeRepeating_MatchesTheStaticMethodAndExpandsOncePerEntry()
    {
        var sut = new DocxMailMergeService();
        var docx = RepeatingTemplate();
        var regions = new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>
        {
            ["Items"] = ThreeEmptyRecords(),
        };

        var fromWrapper = sut.MergeRepeating(docx, regions);

        Assert.Equal(
            DocxEditor.ExtractText(DocxMailMerge.MergeRepeating(docx, regions)),
            DocxEditor.ExtractText(fromWrapper));
        Assert.Equal(3, DocxEditor.ExtractText(fromWrapper).Split("Row.").Length - 1);
    }

    [Fact]
    public async Task MergeRepeatingAsync_MatchesTheStaticMethod()
    {
        var sut = new DocxMailMergeService();
        var docx = RepeatingTemplate();
        var regions = new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>
        {
            ["Items"] = ThreeEmptyRecords(),
        };

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        await sut.MergeRepeatingAsync(source, destination, regions);

        var text = DocxEditor.ExtractText(destination.ToArray());
        Assert.Equal(3, text.Split("Row.").Length - 1);
    }

    [Fact]
    public void MergeRepeatingWithReport_MatchesTheStaticMethodAndReportsWhatWasMissing()
    {
        var sut = new DocxMailMergeService();
        var docx = RepeatingTemplate();
        var empty = new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>();

        var fromWrapper = sut.MergeRepeatingWithReport(docx, empty);

        Assert.Equal(new[] { "Items" }, fromWrapper.Report.MissingNames);
        Assert.DoesNotContain("Row.", DocxEditor.ExtractText(fromWrapper.Document));
    }

    [Fact]
    public async Task MergeRepeatingWithReportAsync_MatchesTheStaticMethod()
    {
        var sut = new DocxMailMergeService();
        var docx = RepeatingTemplate();
        var empty = new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>();

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        var report = await sut.MergeRepeatingWithReportAsync(source, destination, empty);

        Assert.Equal(new[] { "Items" }, report.MissingNames);
    }

    // ---------------------------------------------------------------------------------------
    // MergeRepeatingRegions family, mirrored from core 0.43.0 (A87-DI). Two nesting levels, so
    // the discriminator distinguishes the OUTER count from the INNER count - a level-swap bug
    // (passing Orders where Lines belongs) would still "expand something" but the wrong number
    // of times at the wrong level.
    // ---------------------------------------------------------------------------------------

    private static byte[] NestedRepeatingTemplate() => DocxEditor.Create([
        DocxBlock.Paragraph("{{#each Orders}}"),
        DocxBlock.Paragraph("Order."),
        DocxBlock.Paragraph("{{#each Lines}}"),
        DocxBlock.Paragraph("Line."),
        DocxBlock.Paragraph("{{/each Lines}}"),
        DocxBlock.Paragraph("{{/each Orders}}"),
    ]);

    private static Dictionary<string, IEnumerable<DocxMailMergeBlockData>> TwoOrdersWithLines() =>
        new()
        {
            ["Orders"] =
            [
                new DocxMailMergeBlockData(
                    new Dictionary<string, string>(),
                    new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
                    {
                        ["Lines"] =
                        [
                            new DocxMailMergeBlockData(new Dictionary<string, string>()),
                            new DocxMailMergeBlockData(new Dictionary<string, string>()),
                        ],
                    }),
                new DocxMailMergeBlockData(
                    new Dictionary<string, string>(),
                    new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>
                    {
                        ["Lines"] = [new DocxMailMergeBlockData(new Dictionary<string, string>())],
                    }),
            ],
        };

    [Fact]
    public void MergeRepeatingRegions_MatchesTheStaticMethodAndExpandsBothLevelsCorrectly()
    {
        var sut = new DocxMailMergeService();
        var docx = NestedRepeatingTemplate();
        var regions = TwoOrdersWithLines();

        var fromWrapper = sut.MergeRepeatingRegions(docx, regions);

        Assert.Equal(
            DocxEditor.ExtractText(DocxMailMerge.MergeRepeatingRegions(docx, regions)),
            DocxEditor.ExtractText(fromWrapper));

        var text = DocxEditor.ExtractText(fromWrapper);
        Assert.Equal(2, text.Split("Order.").Length - 1);
        Assert.Equal(3, text.Split("Line.").Length - 1);
    }

    [Fact]
    public async Task MergeRepeatingRegionsAsync_MatchesTheStaticMethod()
    {
        var sut = new DocxMailMergeService();
        var docx = NestedRepeatingTemplate();
        var regions = TwoOrdersWithLines();

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        await sut.MergeRepeatingRegionsAsync(source, destination, regions);

        var text = DocxEditor.ExtractText(destination.ToArray());
        Assert.Equal(2, text.Split("Order.").Length - 1);
        Assert.Equal(3, text.Split("Line.").Length - 1);
    }

    [Fact]
    public void MergeRepeatingRegionsWithReport_MatchesTheStaticMethodAndReportsWhatWasMissing()
    {
        var sut = new DocxMailMergeService();
        var docx = NestedRepeatingTemplate();
        var empty = new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>();

        var fromWrapper = sut.MergeRepeatingRegionsWithReport(docx, empty);

        // Both levels are unsupplied - the whole regions dictionary is empty - so both "Orders"
        // and the nested "Lines" are named, per this method's own documented "unsupplied
        // everywhere" rule.
        Assert.Equal(new[] { "Orders", "Lines" }, fromWrapper.Report.MissingNames);
    }

    [Fact]
    public async Task MergeRepeatingRegionsWithReportAsync_MatchesTheStaticMethod()
    {
        var sut = new DocxMailMergeService();
        var docx = NestedRepeatingTemplate();
        var empty = new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>();

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        var report = await sut.MergeRepeatingRegionsWithReportAsync(source, destination, empty);

        Assert.Equal(new[] { "Orders", "Lines" }, report.MissingNames);
    }

    // ---------------------------------------------------------------------------------------
    // MergeTableRows / MergeTableRowGroups, mirrored from core 0.43.0 (A87-DI). Index-based, not
    // marker-based - the discriminator is the FULL GRID read back via DocxEditor.ReadTable,
    // compared between the wrapper's result and the static method's own, so a swapped
    // tableIndex/templateRowIndex/groupTemplateRowIndex/detailTemplateRowIndex argument shows up
    // as a grid mismatch rather than merely "some table changed".
    // ---------------------------------------------------------------------------------------

    private static byte[] TableRowsTemplate() => DocxEditor.Create([
        DocxBlock.Table(new[] { "Name" }, new[] { new object?[] { "TEMPLATE" } }),
    ]);

    [Fact]
    public void MergeTableRows_MatchesTheStaticMethodAndExpandsTheRightRow()
    {
        var sut = new DocxMailMergeService();
        var docx = TableRowsTemplate();
        var rows = new IReadOnlyDictionary<string, string>[]
        {
            new Dictionary<string, string> { ["x"] = "1" },
            new Dictionary<string, string> { ["x"] = "2" },
        };

        var fromWrapper = sut.MergeTableRows(docx, tableIndex: 0, templateRowIndex: 1, rows);

        Assert.Equal(
            DocxEditor.ReadTable(DocxMailMerge.MergeTableRows(docx, 0, 1, rows), 0),
            DocxEditor.ReadTable(fromWrapper, 0));
        // Header + 2 generated rows = 3, not the 1 the un-expanded template would leave.
        Assert.Equal(3, DocxEditor.ReadTable(fromWrapper, 0).Count);
    }

    [Fact]
    public async Task MergeTableRowsAsync_MatchesTheStaticMethod()
    {
        var sut = new DocxMailMergeService();
        var docx = TableRowsTemplate();
        var rows = new IReadOnlyDictionary<string, string>[]
        {
            new Dictionary<string, string> { ["x"] = "1" },
        };

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        await sut.MergeTableRowsAsync(source, destination, 0, 1, rows);

        Assert.Equal(2, DocxEditor.ReadTable(destination.ToArray(), 0).Count);
    }

    private static byte[] TableRowGroupsTemplate() => DocxEditor.Create([
        DocxBlock.Table(
            new[] { "Group" },
            new[]
            {
                new object?[] { "GROUP_TEMPLATE" },
                new object?[] { "DETAIL_TEMPLATE" },
            }),
    ]);

    private static DocxMailMergeTableRowGroup[] TwoGroups() =>
    [
        new(new Dictionary<string, string>(), new[]
        {
            new Dictionary<string, string>(), new Dictionary<string, string>(),
        }),
        new(new Dictionary<string, string>(), new[] { new Dictionary<string, string>() }),
    ];

    [Fact]
    public void MergeTableRowGroups_MatchesTheStaticMethodAndExpandsTheRightRows()
    {
        var sut = new DocxMailMergeService();
        var docx = TableRowGroupsTemplate();
        var groups = TwoGroups();

        var fromWrapper = sut.MergeTableRowGroups(
            docx, tableIndex: 0, groupTemplateRowIndex: 1, detailTemplateRowIndex: 2, groups);

        Assert.Equal(
            DocxEditor.ReadTable(
                DocxMailMerge.MergeTableRowGroups(docx, 0, 1, 2, groups), 0),
            DocxEditor.ReadTable(fromWrapper, 0));

        // Header + (2 groups, 2+1 detail rows) = 1 + 2 + 3 = 6 rows.
        Assert.Equal(6, DocxEditor.ReadTable(fromWrapper, 0).Count);
    }

    [Fact]
    public async Task MergeTableRowGroupsAsync_MatchesTheStaticMethod()
    {
        var sut = new DocxMailMergeService();
        var docx = TableRowGroupsTemplate();
        var groups = TwoGroups();

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        await sut.MergeTableRowGroupsAsync(source, destination, 0, 1, 2, groups);

        Assert.Equal(6, DocxEditor.ReadTable(destination.ToArray(), 0).Count);
    }

    // ---------------------------------------------------------------------------------------
    // MergeBatch family, mirrored from core 0.43.0 (A85-DI). Discriminator is the SEQUENCE
    // LENGTH and per-item ORDER, since a swapped argument or the wrong strict flag changes how
    // many documents come out, not what any single one contains (this fixture has no merge
    // fields, so every record's own document is byte-identical - what varies is the sequence).
    // ---------------------------------------------------------------------------------------

    private static byte[] BatchTemplate() => DocxEditor.Create([DocxBlock.Paragraph("Body.")]);

    private static IReadOnlyDictionary<string, string>[] TwoEmptyRecords() =>
    [
        new Dictionary<string, string>(),
        new Dictionary<string, string>(),
    ];

    [Fact]
    public void MergeBatch_MatchesTheStaticMethod()
    {
        var sut = new DocxMailMergeService();
        var docx = BatchTemplate();
        var records = TwoEmptyRecords();

        var fromWrapper = sut.MergeBatch(docx, records).ToList();
        var fromStatic = DocxMailMerge.MergeBatch(docx, records).ToList();

        Assert.Equal(fromStatic.Count, fromWrapper.Count);
        Assert.Equal(2, fromWrapper.Count);
        for (var i = 0; i < fromStatic.Count; i++)
            Assert.Equal(DocxEditor.ExtractText(fromStatic[i]), DocxEditor.ExtractText(fromWrapper[i]));
    }

    [Fact]
    public async Task MergeBatchAsync_MatchesTheStaticMethod()
    {
        var sut = new DocxMailMergeService();
        var docx = BatchTemplate();
        var records = TwoEmptyRecords();

        var fromWrapper = new List<byte[]>();
        await foreach (var doc in sut.MergeBatchAsync(docx, records)) fromWrapper.Add(doc);

        Assert.Equal(2, fromWrapper.Count);
        Assert.All(fromWrapper, d => Assert.Contains("Body.", DocxEditor.ExtractText(d)));
    }

    [Fact]
    public void MergeBatchWithReport_MatchesTheStaticMethod()
    {
        var sut = new DocxMailMergeService();
        var docx = BatchTemplate();
        var records = TwoEmptyRecords();

        var fromWrapper = sut.MergeBatchWithReport(docx, records).ToList();

        Assert.Equal(2, fromWrapper.Count);
        Assert.All(fromWrapper, item => Assert.True(item.Report.IsComplete));
    }

    [Fact]
    public async Task MergeBatchWithReportAsync_MatchesTheStaticMethod()
    {
        var sut = new DocxMailMergeService();
        var docx = BatchTemplate();
        var records = TwoEmptyRecords();

        var fromWrapper = new List<DocxMailMergeBatchItem>();
        await foreach (var item in sut.MergeBatchWithReportAsync(docx, records)) fromWrapper.Add(item);

        Assert.Equal(2, fromWrapper.Count);
        Assert.All(fromWrapper, item => Assert.True(item.Report.IsComplete));
    }
}
