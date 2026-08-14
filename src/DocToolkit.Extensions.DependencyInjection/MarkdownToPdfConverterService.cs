namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Default <see cref="IMarkdownToPdfConverter"/>, delegating to
/// <see cref="DocToolkit.MarkdownToPdfConverter"/>.
/// </summary>
internal sealed class MarkdownToPdfConverterService : IMarkdownToPdfConverter
{
    public byte[] Convert(string markdown) => DocToolkit.MarkdownToPdfConverter.Convert(markdown);

    public Task ConvertAsync(string markdown, Stream destination, CancellationToken ct = default)
        => DocToolkit.MarkdownToPdfConverter.ConvertAsync(markdown, destination, ct);

    public DocToolkit.ConversionResult<byte[]> ConvertWithReport(string markdown)
        => DocToolkit.MarkdownToPdfConverter.ConvertWithReport(markdown);
}
