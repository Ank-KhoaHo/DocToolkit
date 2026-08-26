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
}
