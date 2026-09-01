namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IPdfEditor"/>, delegating to <see cref="DocToolkit.PdfEditor"/>.</summary>
internal sealed class PdfEditorService : IPdfEditor
{
    public int PageCount(byte[] pdf) => DocToolkit.PdfEditor.PageCount(pdf);

    public Task<int> PageCountAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.PdfEditor.PageCountAsync(source, ct);

    public IReadOnlyList<string> ExtractText(byte[] pdf) => DocToolkit.PdfEditor.ExtractText(pdf);

    public Task<IReadOnlyList<string>> ExtractTextAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.PdfEditor.ExtractTextAsync(source, ct);

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

    public byte[] RemovePages(byte[] pdf, int firstPage, int count)
        => DocToolkit.PdfEditor.RemovePages(pdf, firstPage, count);

    public Task RemovePagesAsync(
        Stream source, int firstPage, int count, Stream destination, CancellationToken ct = default)
        => DocToolkit.PdfEditor.RemovePagesAsync(source, firstPage, count, destination, ct);

    public byte[] RotatePages(byte[] pdf, int firstPage, int count, int degrees)
        => DocToolkit.PdfEditor.RotatePages(pdf, firstPage, count, degrees);

    public Task RotatePagesAsync(
        Stream source, int firstPage, int count, int degrees, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.PdfEditor.RotatePagesAsync(source, firstPage, count, degrees, destination, ct);

    public byte[] ReorderPages(byte[] pdf, IEnumerable<int> order)
        => DocToolkit.PdfEditor.ReorderPages(pdf, order);

    public Task ReorderPagesAsync(
        Stream source, IEnumerable<int> order, Stream destination, CancellationToken ct = default)
        => DocToolkit.PdfEditor.ReorderPagesAsync(source, order, destination, ct);

    public byte[] InsertPages(byte[] target, byte[] source, int atPage)
        => DocToolkit.PdfEditor.InsertPages(target, source, atPage);

    public Task InsertPagesAsync(
        Stream target, Stream source, int atPage, Stream destination, CancellationToken ct = default)
        => DocToolkit.PdfEditor.InsertPagesAsync(target, source, atPage, destination, ct);

    public byte[] Protect(byte[] pdf, DocToolkit.PdfProtection protection)
        => DocToolkit.PdfEditor.Protect(pdf, protection);

    public byte[] Unprotect(byte[] pdf, string password)
        => DocToolkit.PdfEditor.Unprotect(pdf, password);

    public Task ProtectAsync(
        Stream source, Stream destination, DocToolkit.PdfProtection protection,
        CancellationToken ct = default)
        => DocToolkit.PdfEditor.ProtectAsync(source, destination, protection, ct);

    public Task UnprotectAsync(
        Stream source, Stream destination, string password, CancellationToken ct = default)
        => DocToolkit.PdfEditor.UnprotectAsync(source, destination, password, ct);

    public Task<DocToolkit.PdfMetadata> ReadMetadataAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.PdfEditor.ReadMetadataAsync(source, ct);

    public Task WithMetadataAsync(Stream source, DocToolkit.PdfMetadata metadata, Stream destination, CancellationToken ct = default)
        => DocToolkit.PdfEditor.WithMetadataAsync(source, metadata, destination, ct);
}
