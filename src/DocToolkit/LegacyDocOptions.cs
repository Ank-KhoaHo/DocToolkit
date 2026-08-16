namespace DocToolkit;

/// <summary>
/// Controls how <see cref="DocToDocxConverter"/> treats content a Word 97-2003 binary document
/// holds but a .docx package cannot be given.
/// </summary>
/// <remarks>
/// <b>There is one setting, and it defaults to the safe answer.</b> A legacy .doc stores pictures,
/// drawings and form fields in a binary <c>Data</c> stream that the import can see but cannot carry
/// across. By default the conversion <b>refuses</b> rather than silently producing a .docx missing
/// those payloads — set <see cref="AllowContentLoss"/> to accept the loss deliberately.
///
/// <b>This is the common case, not an edge case.</b> Measured 2026-08-16 against files produced by
/// Word itself: a .doc holding a <b>table</b> carries that stream, while plain text, bold runs and
/// headings do not. Tables are ordinary, so most real documents need the opt-in.
///
/// <b>Reading is never affected.</b> <see cref="DocToDocxConverter.ExtractText(byte[])"/> takes no
/// options and never refuses: text is not what the binary stream holds, so there is nothing for a
/// policy to decide.
/// </remarks>
public sealed class LegacyDocOptions
{
    /// <summary>
    /// Convert even when the source holds content the .docx cannot carry, instead of throwing.
    /// Defaults to <see langword="false"/>.
    /// </summary>
    /// <remarks>
    /// What is lost is the unprojected binary payload — pictures, drawings and form fields.
    /// Measured 2026-08-16: text, tables (every cell), and character formatting such as bold all
    /// survive the conversion intact, so this accepts a specific, bounded loss rather than a
    /// general "best effort".
    ///
    /// Prefer <see cref="DocToDocxConverter.ConvertWithReport(byte[], LegacyDocOptions?)"/> when
    /// setting this: it returns the same bytes and tells you exactly what was dropped, so the loss
    /// is recorded rather than merely permitted.
    /// </remarks>
    public bool AllowContentLoss { get; init; }
}
