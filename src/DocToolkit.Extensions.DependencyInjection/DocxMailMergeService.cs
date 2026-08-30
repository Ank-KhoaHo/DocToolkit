namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Default <see cref="IDocxMailMerge"/>, delegating to <see cref="DocToolkit.DocxMailMerge"/>.
/// </summary>
internal sealed class DocxMailMergeService : IDocxMailMerge
{
    public DocToolkit.DocxMailMergeTemplate InspectTemplate(byte[] docx)
        => DocToolkit.DocxMailMerge.InspectTemplate(docx);

    public Task<DocToolkit.DocxMailMergeTemplate> InspectTemplateAsync(
        Stream source, CancellationToken ct = default)
        => DocToolkit.DocxMailMerge.InspectTemplateAsync(source, ct);

    public byte[] Merge(byte[] docx, IReadOnlyDictionary<string, string> values)
        => DocToolkit.DocxMailMerge.Merge(docx, values);

    public Task MergeAsync(
        Stream source, Stream destination, IReadOnlyDictionary<string, string> values,
        CancellationToken ct = default)
        => DocToolkit.DocxMailMerge.MergeAsync(source, destination, values, ct);

    public DocToolkit.DocxMailMergeResult MergeWithReport(
        byte[] docx, IReadOnlyDictionary<string, string> values)
        => DocToolkit.DocxMailMerge.MergeWithReport(docx, values);

    public Task<DocToolkit.DocxMailMergeReport> MergeWithReportAsync(
        Stream source, Stream destination, IReadOnlyDictionary<string, string> values,
        CancellationToken ct = default)
        => DocToolkit.DocxMailMerge.MergeWithReportAsync(source, destination, values, ct);

    public byte[] MergeConditional(byte[] docx, IReadOnlyDictionary<string, bool> conditions)
        => DocToolkit.DocxMailMerge.MergeConditional(docx, conditions);

    public Task MergeConditionalAsync(
        Stream source, Stream destination, IReadOnlyDictionary<string, bool> conditions,
        CancellationToken ct = default)
        => DocToolkit.DocxMailMerge.MergeConditionalAsync(source, destination, conditions, ct);

    public DocToolkit.DocxMailMergeBlockResult MergeConditionalWithReport(
        byte[] docx, IReadOnlyDictionary<string, bool> conditions)
        => DocToolkit.DocxMailMerge.MergeConditionalWithReport(docx, conditions);

    public Task<DocToolkit.DocxMailMergeBlockReport> MergeConditionalWithReportAsync(
        Stream source, Stream destination, IReadOnlyDictionary<string, bool> conditions,
        CancellationToken ct = default)
        => DocToolkit.DocxMailMerge.MergeConditionalWithReportAsync(source, destination, conditions, ct);

    public byte[] MergeRepeating(
        byte[] docx, IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> regions)
        => DocToolkit.DocxMailMerge.MergeRepeating(docx, regions);

    public Task MergeRepeatingAsync(
        Stream source, Stream destination,
        IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> regions,
        CancellationToken ct = default)
        => DocToolkit.DocxMailMerge.MergeRepeatingAsync(source, destination, regions, ct);

    public DocToolkit.DocxMailMergeBlockResult MergeRepeatingWithReport(
        byte[] docx, IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> regions)
        => DocToolkit.DocxMailMerge.MergeRepeatingWithReport(docx, regions);

    public Task<DocToolkit.DocxMailMergeBlockReport> MergeRepeatingWithReportAsync(
        Stream source, Stream destination,
        IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> regions,
        CancellationToken ct = default)
        => DocToolkit.DocxMailMerge.MergeRepeatingWithReportAsync(source, destination, regions, ct);

    public byte[] MergeRepeatingRegions(
        byte[] docx, IReadOnlyDictionary<string, IEnumerable<DocToolkit.DocxMailMergeBlockData>> regions)
        => DocToolkit.DocxMailMerge.MergeRepeatingRegions(docx, regions);

    public Task MergeRepeatingRegionsAsync(
        Stream source, Stream destination,
        IReadOnlyDictionary<string, IEnumerable<DocToolkit.DocxMailMergeBlockData>> regions,
        CancellationToken ct = default)
        => DocToolkit.DocxMailMerge.MergeRepeatingRegionsAsync(source, destination, regions, ct);

    public DocToolkit.DocxMailMergeBlockResult MergeRepeatingRegionsWithReport(
        byte[] docx, IReadOnlyDictionary<string, IEnumerable<DocToolkit.DocxMailMergeBlockData>> regions)
        => DocToolkit.DocxMailMerge.MergeRepeatingRegionsWithReport(docx, regions);

    public Task<DocToolkit.DocxMailMergeBlockReport> MergeRepeatingRegionsWithReportAsync(
        Stream source, Stream destination,
        IReadOnlyDictionary<string, IEnumerable<DocToolkit.DocxMailMergeBlockData>> regions,
        CancellationToken ct = default)
        => DocToolkit.DocxMailMerge.MergeRepeatingRegionsWithReportAsync(source, destination, regions, ct);

    public byte[] MergeTableRows(
        byte[] docx, int tableIndex, int templateRowIndex,
        IEnumerable<IReadOnlyDictionary<string, string>> rows)
        => DocToolkit.DocxMailMerge.MergeTableRows(docx, tableIndex, templateRowIndex, rows);

    public Task MergeTableRowsAsync(
        Stream source, Stream destination, int tableIndex, int templateRowIndex,
        IEnumerable<IReadOnlyDictionary<string, string>> rows, CancellationToken ct = default)
        => DocToolkit.DocxMailMerge.MergeTableRowsAsync(
            source, destination, tableIndex, templateRowIndex, rows, ct);

    public byte[] MergeTableRowGroups(
        byte[] docx, int tableIndex, int groupTemplateRowIndex, int detailTemplateRowIndex,
        IEnumerable<DocToolkit.DocxMailMergeTableRowGroup> groups)
        => DocToolkit.DocxMailMerge.MergeTableRowGroups(
            docx, tableIndex, groupTemplateRowIndex, detailTemplateRowIndex, groups);

    public Task MergeTableRowGroupsAsync(
        Stream source, Stream destination, int tableIndex, int groupTemplateRowIndex,
        int detailTemplateRowIndex, IEnumerable<DocToolkit.DocxMailMergeTableRowGroup> groups,
        CancellationToken ct = default)
        => DocToolkit.DocxMailMerge.MergeTableRowGroupsAsync(
            source, destination, tableIndex, groupTemplateRowIndex, detailTemplateRowIndex, groups, ct);

    public IEnumerable<byte[]> MergeBatch(byte[] docx, IEnumerable<IReadOnlyDictionary<string, string>> records)
        => DocToolkit.DocxMailMerge.MergeBatch(docx, records);

    public IAsyncEnumerable<byte[]> MergeBatchAsync(
        byte[] docx, IEnumerable<IReadOnlyDictionary<string, string>> records,
        CancellationToken ct = default)
        => DocToolkit.DocxMailMerge.MergeBatchAsync(docx, records, ct);

    public IEnumerable<DocToolkit.DocxMailMergeBatchItem> MergeBatchWithReport(
        byte[] docx, IEnumerable<IReadOnlyDictionary<string, string>> records)
        => DocToolkit.DocxMailMerge.MergeBatchWithReport(docx, records);

    public IAsyncEnumerable<DocToolkit.DocxMailMergeBatchItem> MergeBatchWithReportAsync(
        byte[] docx, IEnumerable<IReadOnlyDictionary<string, string>> records,
        CancellationToken ct = default)
        => DocToolkit.DocxMailMerge.MergeBatchWithReportAsync(docx, records, ct);
}
