using Microsoft.Extensions.Options;

namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Default <see cref="IDocxToPdfConverter"/>, delegating to <see cref="DocToolkit.DocxToPdfConverter"/>
/// with <see cref="DocToolkitOptions.Fonts"/>.
/// </summary>
/// <remarks>
/// <b>Reads the options on every call rather than capturing them at construction</b>, exactly as
/// <see cref="HtmlToDocxConverterService"/> does: these services are registered as singletons, so a
/// captured value would make a configuration change require a restart - and the one time somebody
/// most wants to change configuration is when something is going wrong.
/// </remarks>
internal sealed class DocxToPdfConverterService(IOptionsMonitor<DocToolkitOptions> options)
    : IDocxToPdfConverter
{
    public byte[] Convert(byte[] docx) =>
        DocToolkit.DocxToPdfConverter.Convert(docx, options.CurrentValue.Fonts);

    public Task ConvertAsync(Stream source, Stream destination, CancellationToken ct = default)
        => DocToolkit.DocxToPdfConverter.ConvertAsync(source, destination, options.CurrentValue.Fonts, ct);
}
