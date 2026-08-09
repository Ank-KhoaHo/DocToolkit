namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IDocxToHtmlConverter"/>, delegating to <see cref="DocToolkit.DocxToHtmlConverter"/>.</summary>
internal sealed class DocxToHtmlConverterService : IDocxToHtmlConverter
{
    public string Convert(byte[] docx) => DocToolkit.DocxToHtmlConverter.Convert(docx);

    public Task<string> ConvertAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.DocxToHtmlConverter.ConvertAsync(source, ct);
}
