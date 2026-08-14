namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Default <see cref="IMarkdownToDocxConverter"/>, delegating to
/// <see cref="DocToolkit.MarkdownToDocxConverter"/>.
/// </summary>
internal sealed class MarkdownToDocxConverterService : IMarkdownToDocxConverter
{
    public byte[] Convert(string markdown) => DocToolkit.MarkdownToDocxConverter.Convert(markdown);

    public Task ConvertAsync(string markdown, Stream destination, CancellationToken ct = default)
        => DocToolkit.MarkdownToDocxConverter.ConvertAsync(markdown, destination, ct);

    public DocToolkit.ConversionResult<byte[]> ConvertWithReport(string markdown)
        => DocToolkit.MarkdownToDocxConverter.ConvertWithReport(markdown);
}
