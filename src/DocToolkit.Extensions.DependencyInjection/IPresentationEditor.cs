namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Creates, reads and edits PowerPoint (.pptx) presentations. Registered by <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.</summary>
public interface IPresentationEditor
{
    /// <summary>
    /// Builds a deck from <paramref name="slides"/>, one slide each — a title and bullet lines.
    /// Content comes from data rather than a template, so there is no source file to edit. An empty
    /// sequence is valid and produces a valid deck with no slides.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="slides"/> is null.</exception>
    /// <exception cref="ArgumentException">An element of <paramref name="slides"/> is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The deck could not be built.</exception>
    byte[] Create(IEnumerable<DocToolkit.PptxSlide> slides);

    /// <summary>Number of slides in the deck, as counted from the deck's slide list.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    int SlideCount(byte[] pptx);

    /// <summary>All text found on every slide, one entry per text-bearing body, in deck order.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    IReadOnlyList<string> ExtractText(byte[] pptx);

    /// <summary>Replaces every key with its value across all slide text, returning updated bytes.</summary>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or edited.</exception>
    byte[] ReplaceText(byte[] pptx, IReadOnlyDictionary<string, string> replacements);

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/> and returns its slide count, counted from the
    /// deck's slide list. <paramref name="source"/> is <b>read</b> to its end and is neither
    /// disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    Task<int> SlideCountAsync(Stream source, CancellationToken ct = default);

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/> and returns all text found on every slide, in
    /// deck order. See <see cref="ExtractText(byte[])"/> for exactly what counts as a text-bearing
    /// body. <paramref name="source"/> is <b>read</b> to its end and is neither disposed, closed nor
    /// sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    Task<IReadOnlyList<string>> ExtractTextAsync(Stream source, CancellationToken ct = default);

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/>, replaces every key with its value across all
    /// slide text, and writes the result to <paramref name="destination"/>. See
    /// <see cref="ReplaceText"/> for exactly what counts as a match. <paramref name="source"/> is
    /// <b>read</b> to its end and <paramref name="destination"/> is <b>written</b>; neither is
    /// disposed, closed or sought.
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
    /// Builds a deck from <paramref name="slides"/> and writes it to <paramref name="destination"/>.
    /// See <see cref="Create"/> for the slide semantics. <paramref name="destination"/> is
    /// <b>written</b> and is neither disposed, closed nor sought, so an HTTP response body is a
    /// valid destination.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="slides"/> or <paramref name="destination"/> is null.</exception>
    /// <exception cref="ArgumentException">An element of <paramref name="slides"/> is null, or <paramref name="destination"/> is not writable.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The deck could not be built or written.</exception>
    Task CreateAsync(
        IEnumerable<DocToolkit.PptxSlide> slides, Stream destination, CancellationToken ct = default);
}
