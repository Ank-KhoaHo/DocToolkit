namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IDocxEditor"/>, delegating to <see cref="DocToolkit.DocxEditor"/>.</summary>
internal sealed class DocxEditorService : IDocxEditor
{
    public byte[] Create(IEnumerable<DocToolkit.DocxBlock> blocks)
        => DocToolkit.DocxEditor.Create(blocks);

    public byte[] ReplaceText(byte[] docx, IReadOnlyDictionary<string, string> replacements)
        => DocToolkit.DocxEditor.ReplaceText(docx, replacements);

    public byte[] FillRows(
        byte[] docx, string collection,
        IEnumerable<IReadOnlyDictionary<string, string>> rows)
        => DocToolkit.DocxEditor.FillRows(docx, collection, rows);

    public byte[] ReplaceImage(
        byte[] docx, string placeholder, byte[] image,
        double? widthPoints = null, double? heightPoints = null)
        => DocToolkit.DocxEditor.ReplaceImage(docx, placeholder, image, widthPoints, heightPoints);

    public string ExtractText(byte[] docx) => DocToolkit.DocxEditor.ExtractText(docx);

    public string ExtractText(byte[] docx, bool includeHeadersAndFooters)
        => DocToolkit.DocxEditor.ExtractText(docx, includeHeadersAndFooters);

    public Task ReplaceTextAsync(
        Stream source, IReadOnlyDictionary<string, string> replacements, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.DocxEditor.ReplaceTextAsync(source, replacements, destination, ct);

    public Task<string> ExtractTextAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.DocxEditor.ExtractTextAsync(source, ct);

    public Task<string> ExtractTextAsync(Stream source, bool includeHeadersAndFooters, CancellationToken ct = default)
        => DocToolkit.DocxEditor.ExtractTextAsync(source, includeHeadersAndFooters, ct);

    public Task CreateAsync(
        IEnumerable<DocToolkit.DocxBlock> blocks, Stream destination, CancellationToken ct = default)
        => DocToolkit.DocxEditor.CreateAsync(blocks, destination, ct);

    public Task FillRowsAsync(
        Stream source, string collection,
        IEnumerable<IReadOnlyDictionary<string, string>> rows, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.DocxEditor.FillRowsAsync(source, collection, rows, destination, ct);

    public Task ReplaceImageAsync(
        Stream source, string placeholder, byte[] image, Stream destination,
        double? widthPoints = null, double? heightPoints = null,
        CancellationToken ct = default)
        => DocToolkit.DocxEditor.ReplaceImageAsync(
            source, placeholder, image, destination, widthPoints, heightPoints, ct);

    public byte[] Create(System.Collections.Generic.IEnumerable<DocToolkit.DocxBlock> blocks, DocToolkit.PageSetup page)
        => DocToolkit.DocxEditor.Create(blocks, page);

    public Task CreateAsync(System.Collections.Generic.IEnumerable<DocToolkit.DocxBlock> blocks, DocToolkit.PageSetup page, Stream destination, CancellationToken ct = default)
        => DocToolkit.DocxEditor.CreateAsync(blocks, page, destination, ct);
}
