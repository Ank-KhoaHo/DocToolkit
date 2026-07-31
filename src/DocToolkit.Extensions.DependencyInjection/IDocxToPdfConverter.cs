namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Renders a Word (.docx) package to PDF. Registered by
/// <c>AddDocToolkit</c> (see AddDocToolkit).
/// </summary>
public interface IDocxToPdfConverter
{
    /// <summary>Renders the .docx in <paramref name="docx"/> and returns PDF bytes.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The DOCX could not be rendered.</exception>
    byte[] Convert(byte[] docx);
}
