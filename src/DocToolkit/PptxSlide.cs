namespace DocToolkit;

/// <summary>
/// One slide: a title and zero or more bullet lines. Built by <see cref="Titled"/> and passed to
/// <see cref="PresentationEditor.Create"/>.
///
/// A single sealed type rather than the closed hierarchy <see cref="DocxBlock"/> uses. That
/// hierarchy exists because a document has four kinds of block an external assembly must not be
/// able to extend; a deck has one kind of slide in this version, so the same machinery would be
/// cost without benefit. <see cref="Titled"/> leaves room for other factories later without
/// changing what already exists.
/// </summary>
public sealed class PptxSlide
{
    private PptxSlide(string title, IReadOnlyList<string> bullets)
    {
        Title = title;
        Bullets = bullets;
    }

    /// <summary>The slide's title text. Never null; may be empty.</summary>
    public string Title { get; }

    /// <summary>The bullet lines, in order. Never null; may be empty.</summary>
    public IReadOnlyList<string> Bullets { get; }

    /// <summary>
    /// A slide with <paramref name="title"/> and one bullet line per entry in
    /// <paramref name="bullets"/>. Pass no bullets for a title-only slide.
    ///
    /// Arguments are validated immediately, so a bad value throws at the line that produced it
    /// rather than later inside a <see cref="PresentationEditor.Create"/> call assembling many
    /// slides.
    /// Bullets are materialised here; mutating the caller's array afterwards does not change the
    /// slide.
    /// </summary>
    /// <param name="title">The title text.</param>
    /// <param name="bullets">The bullet lines, in order.</param>
    /// <exception cref="ArgumentNullException"><paramref name="title"/> or <paramref name="bullets"/> is null.</exception>
    /// <exception cref="ArgumentException">An element of <paramref name="bullets"/> is null.</exception>
    public static PptxSlide Titled(string title, params string[] bullets)
    {
        ArgumentNullException.ThrowIfNull(title);
        ArgumentNullException.ThrowIfNull(bullets);

        var materialised = bullets
            .Select((b, index) => b
                ?? throw new ArgumentException($"Bullet {index + 1} was null.", nameof(bullets)))
            .ToList();

        return new PptxSlide(title, materialised);
    }
}
