namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IDocxCompare"/>, delegating to <see cref="DocToolkit.DocxCompare"/>.</summary>
internal sealed class DocxCompareService : IDocxCompare
{
    public byte[] Compare(byte[] original, byte[] revised, string author)
        => DocToolkit.DocxCompare.Compare(original, revised, author);

    public DocToolkit.ConversionResult<byte[]> CompareWithReport(byte[] original, byte[] revised, string author)
        => DocToolkit.DocxCompare.CompareWithReport(original, revised, author);
}
