namespace DocToolkit;

/// <summary>
/// The document information a PDF carries about itself — what a file manager shows in its
/// properties panel, and what a search indexer reads.
/// </summary>
/// <remarks>
/// Every property is nullable, and <see langword="null"/> means <b>absent</b> rather than blank.
/// The distinction matters to anything combining metadata from more than one source: an absent
/// title should be replaced by a fallback, while a deliberately empty one should not.
///
/// <see cref="PdfEditor.WithMetadata"/> applies the same rule in the other direction — a
/// <see langword="null"/> property leaves whatever the document already had alone, so stamping a
/// title does not silently erase an author.
/// </remarks>
public sealed class PdfMetadata
{
    /// <summary>The document's title. Not the file name, and often not set at all.</summary>
    public string? Title { get; init; }

    /// <summary>Who wrote it.</summary>
    public string? Author { get; init; }

    /// <summary>What it is about — a one-line description.</summary>
    public string? Subject { get; init; }

    /// <summary>Comma-separated keywords, by convention rather than by the specification.</summary>
    public string? Keywords { get; init; }

    /// <summary>The application that created the original document.</summary>
    public string? Creator { get; init; }
}
