using System.IO;

namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Renders a Word (.docx) package to PDF. Registered by
/// <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.
/// </summary>
public interface IDocxToPdfConverter
{
    /// <summary>Renders the .docx in <paramref name="docx"/> and returns PDF bytes.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The DOCX could not be rendered.</exception>
    byte[] Convert(byte[] docx);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and writes the rendered PDF to
    /// <paramref name="destination"/>. Neither stream is disposed, closed, sought or read back; the
    /// PDF is written straight through as the renderer produces it, so nothing here ever holds the
    /// whole rendered document in memory. The consequence of streaming is that a failure part-way
    /// through leaves whatever had already been produced on <paramref name="destination"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either stream is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="destination"/>
    /// is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The DOCX could not be rendered.</exception>
    Task ConvertAsync(Stream source, Stream destination, CancellationToken ct = default);
}
