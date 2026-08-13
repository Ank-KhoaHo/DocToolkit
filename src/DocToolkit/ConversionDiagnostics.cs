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
    public static ConversionWarning Warning(string? code, string? message, OfficeConversionLossKind kind)
        => new(code ?? string.Empty, message ?? string.Empty, Map(kind));
}
