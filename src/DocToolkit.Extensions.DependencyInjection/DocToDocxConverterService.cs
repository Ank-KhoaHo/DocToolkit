namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IDocToDocxConverter"/>, delegating to <see cref="DocToolkit.DocToDocxConverter"/>.</summary>
internal sealed class DocToDocxConverterService : IDocToDocxConverter
{
    public byte[] Convert(byte[] doc) => DocToolkit.DocToDocxConverter.Convert(doc);

    public byte[] Convert(byte[] doc, DocToolkit.LegacyDocOptions? options)
        => DocToolkit.DocToDocxConverter.Convert(doc, options);

    public DocToolkit.ConversionResult<byte[]> ConvertWithReport(
        byte[] doc, DocToolkit.LegacyDocOptions? options = null)
        => DocToolkit.DocToDocxConverter.ConvertWithReport(doc, options);

    public string ExtractText(byte[] doc) => DocToolkit.DocToDocxConverter.ExtractText(doc);

    public Task ConvertAsync(Stream source, Stream destination, CancellationToken ct = default)
        => DocToolkit.DocToDocxConverter.ConvertAsync(source, destination, ct);

    public Task ConvertAsync(
        Stream source, Stream destination, DocToolkit.LegacyDocOptions? options,
        CancellationToken ct = default)
        => DocToolkit.DocToDocxConverter.ConvertAsync(source, destination, options, ct);

    public Task<string> ExtractTextAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.DocToDocxConverter.ExtractTextAsync(source, ct);

    public Task<DocToolkit.ConversionResult<byte[]>> ConvertWithReportAsync(
        Stream source, DocToolkit.LegacyDocOptions? options = null, CancellationToken ct = default)
        => DocToolkit.DocToDocxConverter.ConvertWithReportAsync(source, options, ct);
}
