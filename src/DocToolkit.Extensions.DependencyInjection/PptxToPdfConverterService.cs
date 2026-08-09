namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IPptxToPdfConverter"/>, delegating to <see cref="DocToolkit.PptxToPdfConverter"/>.</summary>
internal sealed class PptxToPdfConverterService : IPptxToPdfConverter
{
    public byte[] Convert(byte[] pptx) => DocToolkit.PptxToPdfConverter.Convert(pptx);

    public Task ConvertAsync(Stream source, Stream destination, CancellationToken ct = default)
        => DocToolkit.PptxToPdfConverter.ConvertAsync(source, destination, ct);
}
