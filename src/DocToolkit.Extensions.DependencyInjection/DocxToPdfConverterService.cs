using System.IO;

namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IDocxToPdfConverter"/>, delegating to <see cref="DocToolkit.DocxToPdfConverter"/>.</summary>
internal sealed class DocxToPdfConverterService : IDocxToPdfConverter
{
    public byte[] Convert(byte[] docx) => DocToolkit.DocxToPdfConverter.Convert(docx);

    public Task ConvertAsync(Stream source, Stream destination, CancellationToken ct = default)
        => DocToolkit.DocxToPdfConverter.ConvertAsync(source, destination, ct);
}
