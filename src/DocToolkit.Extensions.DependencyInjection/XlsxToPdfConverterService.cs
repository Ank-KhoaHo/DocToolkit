namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IXlsxToPdfConverter"/>, delegating to <see cref="DocToolkit.XlsxToPdfConverter"/>.</summary>
internal sealed class XlsxToPdfConverterService : IXlsxToPdfConverter
{
    public byte[] Convert(byte[] xlsx) => DocToolkit.XlsxToPdfConverter.Convert(xlsx);

    public Task ConvertAsync(Stream source, Stream destination, CancellationToken ct = default)
        => DocToolkit.XlsxToPdfConverter.ConvertAsync(source, destination, ct);
}
