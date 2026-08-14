namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Converts Markdown to a Word (.docx) package, completing the round trip
/// <see cref="IDocxToMarkdownConverter"/> opens. Registered by
/// <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.
/// </summary>
/// <remarks>
/// <b>Nothing here reaches the network or the disk.</b> A remote image reference becomes a
/// hyperlink rather than a fetch, and a local file reference is refused - reading a path out of
/// untrusted document content would let the document choose which file gets embedded in the
/// output. <c>data:</c> images are inlined, since they carry their own bytes.
///
/// <b>There is no <c>Stream source</c> overload</b>, here or on the static API: the input is a
/// <see cref="string"/> rather than a document, so a caller holding bytes decides their own
/// encoding.
/// </remarks>
public interface IMarkdownToDocxConverter
{
    /// <summary>Converts <paramref name="markdown"/> to a .docx package.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="markdown"/> is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be converted.</exception>
    byte[] Convert(string markdown);

    /// <summary>
    /// Converts <paramref name="markdown"/> and writes the .docx to <paramref name="destination"/>,
    /// which is written but neither disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="destination"/> is not writable.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be converted.</exception>
    Task ConvertAsync(string markdown, Stream destination, CancellationToken ct = default);

    /// <summary>
    /// Converts <paramref name="markdown"/> and reports what the conversion could not carry
    /// across. <see cref="Convert(string)"/> produces the same document.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="markdown"/> is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be converted.</exception>
    DocToolkit.ConversionResult<byte[]> ConvertWithReport(string markdown);
}
