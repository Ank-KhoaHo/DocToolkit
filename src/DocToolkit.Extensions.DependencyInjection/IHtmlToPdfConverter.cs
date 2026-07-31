namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Converts HTML straight to PDF by pivoting through DOCX. Registered by
/// <c>AddDocToolkit</c> (added in Task 7 of this plan); remote image download is controlled
/// once, at registration, via <see cref="DocToolkitOptions.AllowRemoteImageDownload"/>.
/// </summary>
public interface IHtmlToPdfConverter
{
    /// <summary>Converts <paramref name="html"/> straight to PDF bytes.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="html"/> is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The HTML could not be converted.</exception>
    Task<byte[]> ConvertAsync(string html, CancellationToken ct = default);
}
