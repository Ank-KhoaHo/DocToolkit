using OfficeIMO;

namespace DocToolkit;

/// <summary>
/// The one place an upstream conversion diagnostic becomes a <see cref="ConversionWarning"/>.
///
/// Single site for the same reason <c>SectionPropertiesFactory</c> and
/// <c>WorkbookEditor.SetCellValue</c> are: the HTML and Markdown exporters both map the same
/// upstream enum, and two mapping sites is how they come to disagree about what
/// <see cref="ConversionLossKind.Omission"/> means.
/// </summary>
internal static class ConversionDiagnostics
{
    /// <summary>
    /// Maps an upstream loss kind onto this library's own.
    /// </summary>
    /// <remarks>
    /// <b>An unrecognised value maps to <see cref="ConversionLossKind.Failure"/>, the most severe
    /// one.</b> A severity added upstream that this code has never seen must not be silently
    /// downgraded to "nothing was lost" — under-reporting loss is the exact failure this whole
    /// feature exists to prevent, so the default errs loud rather than quiet.
    /// </remarks>
    public static ConversionLossKind Map(OfficeConversionLossKind kind) => kind switch
    {
        OfficeConversionLossKind.None => ConversionLossKind.None,
        OfficeConversionLossKind.Approximation => ConversionLossKind.Approximation,
        OfficeConversionLossKind.Omission => ConversionLossKind.Omission,
        OfficeConversionLossKind.Failure => ConversionLossKind.Failure,
        _ => ConversionLossKind.Failure,
    };

    /// <summary>Builds a warning from the three fields every upstream diagnostic carries.</summary>
    /// <remarks>
    /// <b>The message is passed through VERBATIM, and there is deliberately no rewriting layer
    /// here.</b> This is the obvious place to add one, so the reason not to is worth recording.
    ///
    /// One upstream message does read badly to a consumer of this package: a remote image in
    /// Markdown reports that <c>MarkdownToWordOptions.RemoteImageResolver is not configured</c> —
    /// an option nothing in this API exposes, because leaving it null <i>is</i> the offline
    /// guarantee. It reads as a suggestion to configure something that does not exist here.
    ///
    /// <b>Measured 2026-08-15 before deciding: it is the only one.</b> Across the three upstream
    /// assemblies whose diagnostics can reach a <c>ConvertWithReport</c>, 42 message literals are
    /// diagnostic-shaped and exactly <b>one</b> names an API member the caller cannot reach — that
    /// one. The DOCX → HTML exporter's 34 describe state ("skipped because only data URI images
    /// are enabled") rather than instructing anybody to go and set something.
    ///
    /// So a rewriting layer would be a general mechanism for a single case, and it would have to
    /// match on upstream message text — a hand-maintained mapping that stops matching, silently,
    /// the moment upstream rewords a sentence. This repository has deleted several lists of
    /// exactly that shape. The proportionate fix is documentation, and
    /// <c>docfx/guides/markdown.md</c> carries it.
    ///
    /// Revisit if a second message joins it. The check is a string scrape over the upstream
    /// assemblies, so treat any count from it as a lower bound — an interpolated message assembled
    /// at run time would not appear.
    /// </remarks>
    public static ConversionWarning Warning(string? code, string? message, OfficeConversionLossKind kind)
        => new(code ?? string.Empty, message ?? string.Empty, Map(kind));
}
