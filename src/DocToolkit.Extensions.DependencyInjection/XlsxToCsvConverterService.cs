namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Default <see cref="IXlsxToCsvConverter"/>, delegating to
/// <see cref="DocToolkit.XlsxToCsvConverter"/>.
/// </summary>
internal sealed class XlsxToCsvConverterService : IXlsxToCsvConverter
{
    public string Convert(byte[] xlsx, string sheetName)
        => DocToolkit.XlsxToCsvConverter.Convert(xlsx, sheetName);

    public Task<string> ConvertAsync(Stream source, string sheetName, CancellationToken ct = default)
        => DocToolkit.XlsxToCsvConverter.ConvertAsync(source, sheetName, ct);
}
