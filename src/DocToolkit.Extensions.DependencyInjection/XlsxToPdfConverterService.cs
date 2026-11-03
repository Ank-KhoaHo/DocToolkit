using Microsoft.Extensions.Options;

namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Default <see cref="IXlsxToPdfConverter"/>, delegating to <see cref="DocToolkit.XlsxToPdfConverter"/>
/// with <see cref="DocToolkitOptions.Fonts"/>.
/// </summary>
/// <remarks>
/// <b>Reads the options on every call</b>, as <see cref="DocxToPdfConverterService"/> does and for
/// the same reason. Fonts reached this service only from core 0.58.0, the first release whose
/// <see cref="DocToolkit.XlsxToPdfConverter"/> takes them.
/// </remarks>
internal sealed class XlsxToPdfConverterService(IOptionsMonitor<DocToolkitOptions> options)
    : IXlsxToPdfConverter
{
    public byte[] Convert(byte[] xlsx) => DocToolkit.XlsxToPdfConverter.Convert(xlsx, options.CurrentValue.Fonts);

    public Task ConvertAsync(Stream source, Stream destination, CancellationToken ct = default)
        => DocToolkit.XlsxToPdfConverter.ConvertAsync(source, destination, options.CurrentValue.Fonts, ct);
}
