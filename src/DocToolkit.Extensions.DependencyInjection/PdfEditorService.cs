namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IPdfEditor"/>, delegating to <see cref="DocToolkit.PdfEditor"/>.</summary>
internal sealed class PdfEditorService : IPdfEditor
{
    public int PageCount(byte[] pdf) => DocToolkit.PdfEditor.PageCount(pdf);

    public Task<int> PageCountAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.PdfEditor.PageCountAsync(source, ct);

    public byte[] Merge(IEnumerable<byte[]> pdfs) => DocToolkit.PdfEditor.Merge(pdfs);

    public Task MergeAsync(IEnumerable<Stream> sources, Stream destination, CancellationToken ct = default)
        => DocToolkit.PdfEditor.MergeAsync(sources, destination, ct);

    public byte[] ExtractPages(byte[] pdf, int firstPage, int count)
        => DocToolkit.PdfEditor.ExtractPages(pdf, firstPage, count);

    public Task ExtractPagesAsync(
        Stream source, int firstPage, int count, Stream destination, CancellationToken ct = default)
        => DocToolkit.PdfEditor.ExtractPagesAsync(source, firstPage, count, destination, ct);

    public DocToolkit.PdfMetadata ReadMetadata(byte[] pdf) => DocToolkit.PdfEditor.ReadMetadata(pdf);

    public byte[] WithMetadata(byte[] pdf, DocToolkit.PdfMetadata metadata)
        => DocToolkit.PdfEditor.WithMetadata(pdf, metadata);
}
