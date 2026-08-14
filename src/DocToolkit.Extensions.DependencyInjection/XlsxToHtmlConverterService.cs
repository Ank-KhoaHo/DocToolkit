namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Default <see cref="IXlsxToHtmlConverter"/>, delegating to
/// <see cref="DocToolkit.XlsxToHtmlConverter"/>.
/// </summary>
internal sealed class XlsxToHtmlConverterService : IXlsxToHtmlConverter
{
    public string Convert(byte[] xlsx, string sheetName)
        => DocToolkit.XlsxToHtmlConverter.Convert(xlsx, sheetName);

    public Task<string> ConvertAsync(Stream source, string sheetName, CancellationToken ct = default)
        => DocToolkit.XlsxToHtmlConverter.ConvertAsync(source, sheetName, ct);
}
