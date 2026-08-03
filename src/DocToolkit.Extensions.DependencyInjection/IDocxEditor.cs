namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Opens and edits an existing .docx package. Registered by <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.</summary>
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

    /// <summary>
    /// Reads a .docx from <paramref name="source"/>, replaces every key with its value, and writes
    /// the result to <paramref name="destination"/>. See <see cref="ReplaceText"/> for exactly what
    /// counts as a match. <paramref name="source"/> is <b>read</b> to its end and
    /// <paramref name="destination"/> is <b>written</b>; neither is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="destination"/>
    /// is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or edited.</exception>
    Task ReplaceTextAsync(
        Stream source, IReadOnlyDictionary<string, string> replacements, Stream destination,
        CancellationToken ct = default);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and returns the plain text of its body. Headers,
    /// footers, footnotes and endnotes are not included. <paramref name="source"/> is <b>read</b> to
    /// its end and is neither disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    Task<string> ExtractTextAsync(Stream source, CancellationToken ct = default);

    /// <summary>
    /// Reads a .docx from <paramref name="source"/> and returns its plain text. See
    /// <see cref="ExtractText(byte[], bool)"/> for what <paramref name="includeHeadersAndFooters"/>
    /// controls. <paramref name="source"/> is <b>read</b> to its end and is neither disposed, closed
    /// nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    Task<string> ExtractTextAsync(Stream source, bool includeHeadersAndFooters, CancellationToken ct = default);
}
