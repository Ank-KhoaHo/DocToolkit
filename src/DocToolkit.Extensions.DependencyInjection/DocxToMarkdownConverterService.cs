namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IDocxToMarkdownConverter"/>, delegating to <see cref="DocToolkit.DocxToMarkdownConverter"/>.</summary>
internal sealed class DocxToMarkdownConverterService : IDocxToMarkdownConverter
{
    public string Convert(byte[] docx) => DocToolkit.DocxToMarkdownConverter.Convert(docx);

    public Task<string> ConvertAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.DocxToMarkdownConverter.ConvertAsync(source, ct);
}
