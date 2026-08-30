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

    /// <summary>
    /// Replaces every shape whose text is exactly <paramref name="placeholder"/> with
    /// <paramref name="image"/>, scaled to fit inside that shape's box and centred there. Position
    /// and size come from the template, so there is nothing to pass — a designer draws a box in
    /// PowerPoint where the image belongs and the image lands there. The shape's text must be
    /// nothing but the placeholder; PNG and JPEG only, detected from magic bytes rather than any
    /// filename.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="pptx"/> or <paramref name="image"/> is empty, or <paramref name="placeholder"/>
    /// is blank.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The placeholder appears nowhere, appears only inside a grouped shape, a matched shape holds
    /// other text, a matched shape has no explicit position, the image is neither PNG nor JPEG, or
    /// the package could not be edited.
    /// </exception>
    byte[] ReplaceImage(byte[] pptx, string placeholder, byte[] image);

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/>, replaces every shape whose text is exactly
    /// <paramref name="placeholder"/> with <paramref name="image"/>, and writes the result to
    /// <paramref name="destination"/>. See <see cref="ReplaceImage"/> for exactly what counts as a
    /// match. <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/>
    /// is <b>written</b>; neither is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="placeholder"/> is blank, <paramref name="image"/> is empty,
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="destination"/>
    /// is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The placeholder appears nowhere, appears only inside a grouped shape, a matched shape holds
    /// other text, a matched shape has no explicit position, the image is neither PNG nor JPEG, or
    /// the package could not be edited.
    /// </exception>
    Task ReplaceImageAsync(
        Stream source, string placeholder, byte[] image, Stream destination,
        CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="pptx"/> encrypted with <paramref name="password"/>.
    /// </summary>
    /// <remarks>
    /// <b>File encryption, not the "restrict editing" flag.</b> The result is a compound file rather
    /// than a PPTX package, so every other member here refuses it - call
    /// <see cref="Unprotect(byte[], string)"/> first.
    /// </remarks>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be read or encrypted.</exception>
    byte[] Protect(byte[] pptx, string password);

    /// <summary>A copy of <paramref name="pptx"/> with its encryption removed.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The password was wrong, the presentation was not encrypted, or it could not be read.
    /// </exception>
    byte[] Unprotect(byte[] pptx, string password);

    /// <summary>
    /// Whether <paramref name="pptx"/> is encrypted - that is, whether the other members here
    /// will refuse it. Reads the file signature; needs no password.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> is null.</exception>
    bool IsProtected(byte[] pptx);

    /// <summary>
    /// Reads a presentation from <paramref name="source"/> and writes the encrypted copy to
    /// <paramref name="destination"/>. Neither stream is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either stream is null, or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be encrypted.</exception>
    Task ProtectAsync(Stream source, Stream destination, string password, CancellationToken ct = default);

    /// <summary>
    /// Reads an encrypted presentation from <paramref name="source"/> and writes the unprotected copy to
    /// <paramref name="destination"/>. Neither stream is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either stream is null, or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The password was wrong, or it could not be read.</exception>
    Task UnprotectAsync(Stream source, Stream destination, string password, CancellationToken ct = default);

    /// <summary>
    /// All text found on slide <paramref name="index"/>. See <see cref="ExtractText(byte[])"/> for
    /// exactly what counts as a text-bearing body.
    /// </summary>
    /// <param name="pptx">The presentation to read.</param>
    /// <param name="index">1-based, because that is how a reader numbers slides.</param>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="index"/> is below 1, or above the deck's slide count.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    IReadOnlyList<string> ReadSlide(byte[] pptx, int index);

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/> and returns all text found on slide
    /// <paramref name="index"/> — see <see cref="ReadSlide(byte[], int)"/>. <paramref name="source"/>
    /// is <b>read</b> to its end and is neither disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="index"/> is below 1, or above the deck's slide count.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or read.</exception>
    Task<IReadOnlyList<string>> ReadSlideAsync(Stream source, int index, CancellationToken ct = default);

    /// <summary>A copy of <paramref name="pptx"/> with the slides at <paramref name="indices"/> removed.</summary>
    /// <param name="pptx">The presentation to remove slides from. It is not modified.</param>
    /// <param name="indices">
    /// 1-based slide numbers to remove, each exactly once and in any order — not a contiguous
    /// range.
    /// </param>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> or <paramref name="indices"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="pptx"/> is empty, or <paramref name="indices"/> contains a duplicate.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// An index in <paramref name="indices"/> is outside the deck's slide range, or removing every
    /// listed index would leave a zero-slide deck.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or edited.</exception>
    byte[] RemoveSlides(byte[] pptx, IEnumerable<int> indices);

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/>, removes the slides at
    /// <paramref name="indices"/>, and writes the result to <paramref name="destination"/> — see
    /// <see cref="RemoveSlides"/>. Neither stream is disposed, closed or sought.
    /// </summary>
    /// <param name="source">The stream the .pptx package is read from.</param>
    /// <param name="indices">1-based slide numbers to remove, each exactly once, any order.</param>
    /// <param name="destination">The stream the edited .pptx package is written to.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, <paramref name="destination"/>
    /// is not writable, or <paramref name="indices"/> contains a duplicate.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// An index in <paramref name="indices"/> is outside the deck's slide range, or removing every
    /// listed index would leave a zero-slide deck.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or edited.</exception>
    Task RemoveSlidesAsync(
        Stream source, IEnumerable<int> indices, Stream destination, CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="pptx"/> with its slides in the order given by
    /// <paramref name="order"/>, which holds 1-based slide numbers.
    /// </summary>
    /// <param name="pptx">The presentation to reorder. It is not modified.</param>
    /// <param name="order">A permutation of every slide — the same slides, in a different order.</param>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> or <paramref name="order"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="pptx"/> is empty, or <paramref name="order"/> is not a permutation of
    /// 1..<c>SlideCount</c>.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or edited.</exception>
    byte[] ReorderSlides(byte[] pptx, IEnumerable<int> order);

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/>, reorders its slides per
    /// <paramref name="order"/>, and writes the result to <paramref name="destination"/> — see
    /// <see cref="ReorderSlides"/>. Neither stream is disposed, closed or sought.
    /// </summary>
    /// <param name="source">The stream the .pptx package is read from.</param>
    /// <param name="order">A permutation of every slide's 1-based number.</param>
    /// <param name="destination">The stream the edited .pptx package is written to.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, <paramref name="destination"/>
    /// is not writable, or <paramref name="order"/> is not a permutation of every slide.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The package could not be opened or edited.</exception>
    Task ReorderSlidesAsync(
        Stream source, IEnumerable<int> order, Stream destination, CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="pptx"/> with <paramref name="slides"/> inserted so that the first
    /// of them becomes slide <paramref name="atIndex"/>. See <see cref="Create"/> for what a
    /// <see cref="DocToolkit.PptxSlide"/> becomes.
    /// </summary>
    /// <param name="pptx">The presentation to insert into. It is not modified.</param>
    /// <param name="atIndex">
    /// 1-based position the first inserted slide will occupy. <c>SlideCount + 1</c> appends.
    /// </param>
    /// <param name="slides">The slides to insert, in order.</param>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> or <paramref name="slides"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="pptx"/> is empty, or an element of <paramref name="slides"/> is null.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="atIndex"/> is below 1 or more than one past the last slide.
    /// </exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The package could not be opened or edited, or the slide the new content would attach to has
    /// no layout of its own.
    /// </exception>
    byte[] InsertSlides(byte[] pptx, int atIndex, IEnumerable<DocToolkit.PptxSlide> slides);

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/>, inserts <paramref name="slides"/> at
    /// <paramref name="atIndex"/>, and writes the result to <paramref name="destination"/> — see
    /// <see cref="InsertSlides"/> for the insertion and layout rules. Neither stream is disposed,
    /// closed or sought.
    /// </summary>
    /// <param name="source">The stream the .pptx package is read from.</param>
    /// <param name="atIndex">1-based insertion position; <c>SlideCount + 1</c> appends.</param>
    /// <param name="slides">The slides to insert, in order.</param>
    /// <param name="destination">The stream the edited .pptx package is written to.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, <paramref name="destination"/>
    /// is not writable, or an element of <paramref name="slides"/> is null.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="atIndex"/> is below 1 or more than one past the last slide.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The package could not be opened or edited, or the slide the new content would attach to has
    /// no layout of its own.
    /// </exception>
    Task InsertSlidesAsync(
        Stream source, int atIndex, IEnumerable<DocToolkit.PptxSlide> slides, Stream destination,
        CancellationToken ct = default);
}
