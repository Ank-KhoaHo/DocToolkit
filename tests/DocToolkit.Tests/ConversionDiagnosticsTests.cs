using OfficeIMO;

namespace DocToolkit.Tests;

/// <summary>
/// <see cref="ConversionDiagnostics.Warning"/> is exercised through <c>ConvertWithReport</c> in
/// <c>DocxToHtmlMarkdownTests</c> and <c>DocxToMarkdownConverterTests</c> whenever upstream supplies
/// a code and a message, but neither ever supplies a null one - a fixture cannot force that, since
/// OfficeIMO's own diagnostic type decides it. Found on 2026-08-28 while re-measuring B30: hand
/// applying the mutation Stryker reports (both <c>?? string.Empty</c> fallbacks replaced with a
/// non-empty literal) left the whole suite green. Tested directly here, which
/// <c>src/Directory.Build.props</c>' <c>InternalsVisibleTo</c> grant from DocToolkit.Primitives to
/// DocToolkit.Tests exists to allow.
/// </summary>
public class ConversionDiagnosticsTests
{
    [Fact]
    public void Warning_TreatsANullCodeAndMessageAsEmpty_RatherThanPropagatingNull()
    {
        var warning = ConversionDiagnostics.Warning(null, null, OfficeConversionLossKind.Omission);

        Assert.Equal(string.Empty, warning.Code);
        Assert.Equal(string.Empty, warning.Message);
        Assert.Equal(ConversionLossKind.Omission, warning.Kind);
    }
}
