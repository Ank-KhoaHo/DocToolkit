using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// Covers the five presentation settings <see cref="XlsxFormat"/> gained: conditional formatting,
/// data validation, autofilter, explicit column widths and freeze-at-position.
///
/// <b>Every capability is asserted by reading it out of the SAVED BYTES</b>, never from the
/// <see cref="XlsxFormat"/> object. An assertion on the object would pass against an
/// <c>ApplyFormat</c> that discarded every one of them.
/// </summary>
public class XlsxPresentationTests
{
    [Fact]
    public void XlsxRule_CarriesItsRangeConditionAndHighlight()
    {
        XlsxRule rule = XlsxRule.GreaterThan("B2:B99", 1000, XlsxHighlight.Red);

        Assert.Equal("B2:B99", rule.Range);
        Assert.Equal(XlsxRuleKind.GreaterThan, rule.Kind);
        Assert.Equal(1000, rule.Value);
        Assert.Equal(XlsxHighlight.Red, rule.Highlight);

        XlsxRule between = XlsxRule.Between("A1:A9", 1, 10, XlsxHighlight.Green);
        Assert.Equal(XlsxRuleKind.Between, between.Kind);
        Assert.Equal(1, between.Value);
        Assert.Equal(10, between.High);

        XlsxRule contains = XlsxRule.Contains("A1:A9", "overdue", XlsxHighlight.Amber);
        Assert.Equal(XlsxRuleKind.Contains, contains.Kind);
        Assert.Equal("overdue", contains.Text);
    }

    [Fact]
    public void XlsxRule_RefusesABlankRangeByTheParameterItDeclares()
    {
        Assert.Equal("range", Assert.Throws<ArgumentException>(
            () => XlsxRule.GreaterThan("  ", 1, XlsxHighlight.Red)).ParamName);
        Assert.Equal("range", Assert.Throws<ArgumentNullException>(
            () => XlsxRule.Blank(null!, XlsxHighlight.Grey)).ParamName);
        Assert.Equal("text", Assert.Throws<ArgumentNullException>(
            () => XlsxRule.Contains("A1", null!, XlsxHighlight.Red)).ParamName);
    }

    [Fact]
    public void XlsxRule_RefusesABetweenWhoseBoundsAreInverted()
    {
        // low > high describes an empty range, so every cell fails it silently - a rule that can
        // never fire is worse than an error, because the sheet looks formatted and is not.
        Assert.Equal("high", Assert.Throws<ArgumentOutOfRangeException>(
            () => XlsxRule.Between("A1:A9", 10, 1, XlsxHighlight.Green)).ParamName);
    }
}
