namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Default <see cref="IDocxToPdfPreflight"/>, delegating to
/// <see cref="DocToolkit.DocxToPdfPreflight"/>.
/// </summary>
internal sealed class DocxToPdfPreflightService : IDocxToPdfPreflight
{
    public DocToolkit.DocxToPdfPreflightReport Inspect(byte[] docx)
        => DocToolkit.DocxToPdfPreflight.Inspect(docx);

    public Task<DocToolkit.DocxToPdfPreflightReport> InspectAsync(
        Stream source, CancellationToken ct = default)
        => DocToolkit.DocxToPdfPreflight.InspectAsync(source, ct);
}
