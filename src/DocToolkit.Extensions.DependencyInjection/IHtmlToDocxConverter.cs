namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Converts HTML to a Word (.docx) package. Registered by
/// <c>AddDocToolkit</c> (see AddDocToolkit); remote image download is controlled
/// once, at registration, via <see cref="DocToolkitOptions.AllowRemoteImageDownload"/>.
/// </summary>
public interface IHtmlToDocxConverter
{
    /// <summary>Converts <paramref name="html"/> to the bytes of a .docx package.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The HTML could not be converted.</exception>
    Task<byte[]> ConvertAsync(string html, CancellationToken ct = default);
}
