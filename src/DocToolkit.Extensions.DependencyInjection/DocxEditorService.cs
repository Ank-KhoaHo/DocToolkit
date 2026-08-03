namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IDocxEditor"/>, delegating to <see cref="DocToolkit.DocxEditor"/>.</summary>
internal sealed class DocxEditorService : IDocxEditor
{
    public byte[] ReplaceText(byte[] docx, IReadOnlyDictionary<string, string> replacements)
        => DocToolkit.DocxEditor.ReplaceText(docx, replacements);

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
}
