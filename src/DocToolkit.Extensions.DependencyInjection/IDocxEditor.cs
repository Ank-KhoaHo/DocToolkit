namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Opens and edits an existing .docx package. Registered by <c>AddDocToolkit</c> (see AddDocToolkit).</summary>
public interface IDocxEditor
{
    /// <summary>Replaces every key with its value across the document body, headers, footers, footnotes and endnotes.</summary>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or edited.</exception>
    byte[] ReplaceText(byte[] docx, IReadOnlyDictionary<string, string> replacements);

    /// <summary>Returns the plain text of the document body. Headers, footers, footnotes and endnotes are not included.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    string ExtractText(byte[] docx);

    /// <summary>Returns the plain text of the document. When <paramref name="includeHeadersAndFooters"/> is true, headers and footers follow the body text; footnotes and endnotes are never included.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    string ExtractText(byte[] docx, bool includeHeadersAndFooters);
}
