namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Renders an Excel (.xlsx) workbook to PDF. Registered by
/// <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.
/// </summary>
public interface IXlsxToPdfConverter
{
    /// <summary>Renders the .xlsx in <paramref name="xlsx"/> and returns PDF bytes.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be rendered.</exception>
    byte[] Convert(byte[] xlsx);

    /// <summary>
    /// Reads a .xlsx from <paramref name="source"/> and writes the rendered PDF to
    /// <paramref name="destination"/>. Neither stream is disposed, closed, sought or read back; the
    /// PDF is written straight through as the renderer produces it, so a failure part-way leaves
    /// whatever had already been produced on <paramref name="destination"/>.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either stream is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="destination"/>
    /// is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be rendered.</exception>
    Task ConvertAsync(Stream source, Stream destination, CancellationToken ct = default);
}
