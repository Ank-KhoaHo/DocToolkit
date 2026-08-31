namespace DocToolkit;

/// <summary>
/// The document properties a DOCX, XLSX or PPTX file carries about itself — what a file manager
/// shows in its properties panel, and what a search indexer reads.
/// </summary>
/// <remarks>
/// Shared across all three OOXML formats because their underlying property bags are identical in
/// shape — the same convention as <see cref="DocumentSignatureInfo"/>, which is also one type
/// reused by <c>DocxEditor</c>, <c>WorkbookEditor</c> and <c>PresentationEditor</c> rather than
/// three near-duplicates. It is <b>not</b> shared with the PDF-specific <c>PdfMetadata</c>: PDF's
/// Info dictionary is a different schema, and one field name would otherwise collide in a
/// misleading way — see <see cref="Creator"/> below.
///
/// Every property is nullable, and <see langword="null"/> means <b>absent</b> rather than blank —
/// the same rule <c>PdfMetadata</c> uses, for the same reason: an absent title should be replaced
/// by a fallback, while a deliberately empty one should not. The relevant
/// <c>With*Metadata</c> method on each editor applies the rule in the other direction — a
/// <see langword="null"/> property leaves whatever the document already had alone, so stamping a
/// title does not silently erase an author. Pass an empty string to clear one.
/// </remarks>
public sealed class DocumentMetadata
{
    /// <summary>The document's title. Not the file name, and often not set at all.</summary>
    public string? Title { get; init; }

    /// <summary>
    /// Who wrote it — OOXML's own <c>dc:creator</c> core property.
    /// </summary>
    /// <remarks>
    /// <b>A different concept from PDF's own <c>Creator</c> property.</b> The two ecosystems use
    /// the same word for two different things: OOXML's <c>Creator</c> names the person who
    /// authored the document (what <c>PdfMetadata.Author</c> means for a PDF), while PDF's own
    /// <c>Creator</c> names the application that produced the file. This type keeps OOXML's own
    /// name for its own property rather than renaming it to "Author" and inventing a false
    /// symmetry with <c>PdfMetadata</c> that the two formats do not actually share.
    /// </remarks>
    public string? Creator { get; init; }

    /// <summary>What it is about — a one-line description.</summary>
    public string? Subject { get; init; }

    /// <summary>Comma-separated keywords, by convention rather than by the specification.</summary>
    public string? Keywords { get; init; }
}
