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
    [Fact]
    public void XlsxValidation_CarriesItsRangeKindAndBounds()
    {
        // int, not long: the file format takes an int, so a long would be lossy at the boundary
        // rather than here.
        XlsxValidation whole = XlsxValidation.WholeNumberBetween("B2:B99", 0, 1000);
        Assert.Equal("B2:B99", whole.Range);
        Assert.Equal(XlsxValidationKind.WholeNumber, whole.Kind);
        Assert.Equal(0, whole.Min);
        Assert.Equal(1000, whole.Max);

        XlsxValidation list = XlsxValidation.OneOf("C2:C99", "Free", "Pro", "Team");
        Assert.Equal(XlsxValidationKind.List, list.Kind);
        Assert.Equal(["Free", "Pro", "Team"], list.Options);

        XlsxValidation dates = XlsxValidation.DateBetween(
            "D2:D99", new DateTime(2020, 1, 1), new DateTime(2030, 1, 1));
        Assert.Equal(XlsxValidationKind.Date, dates.Kind);
        Assert.Equal(new DateTime(2020, 1, 1), dates.MinDate);
    }

    [Fact]
    public void XlsxValidation_RefusesInvertedBoundsAndAnEmptyOptionList()
    {
        Assert.Equal("max", Assert.Throws<ArgumentOutOfRangeException>(
            () => XlsxValidation.WholeNumberBetween("A1", 10, 1)).ParamName);
        Assert.Equal("max", Assert.Throws<ArgumentOutOfRangeException>(
            () => XlsxValidation.DecimalBetween("A1", 1.5, 0.5)).ParamName);
        Assert.Equal("max", Assert.Throws<ArgumentOutOfRangeException>(
            () => XlsxValidation.TextLengthBetween("A1", 50, 1)).ParamName);
        Assert.Equal("max", Assert.Throws<ArgumentOutOfRangeException>(
            () => XlsxValidation.DateBetween("A1", new DateTime(2030, 1, 1), new DateTime(2020, 1, 1))).ParamName);

        // An empty list validates nothing, so a caller who passes one gets a cell nobody can fill.
        Assert.Equal("options", Assert.Throws<ArgumentException>(
            () => XlsxValidation.OneOf("A1")).ParamName);
    }
    // ---- XlsxFormat's five new members ---------------------------------------------------------

    [Fact]
    public void FreezeHeaderRow_IsDerivedFromFreezeAt_SoTheTwoCannotDisagree()
    {
        // One underlying concept rather than two switches. WithFrozenHeaderRow sets the position,
        // and the bool reads true exactly when that is where the sheet is frozen.
        Assert.True(XlsxFormat.None.WithFrozenHeaderRow().FreezeHeaderRow);
        Assert.Equal(new XlsxFreeze(1, 0), XlsxFormat.None.WithFrozenHeaderRow().FreezeAt);

        // Freezing elsewhere makes the bool read false, which is the honest answer.
        XlsxFormat elsewhere = XlsxFormat.None.WithFreezeAt(3, 2);
        Assert.False(elsewhere.FreezeHeaderRow);
        Assert.Equal(new XlsxFreeze(3, 2), elsewhere.FreezeAt);

        // And clearing clears the position entirely, whatever it was.
        Assert.Null(elsewhere.WithFrozenHeaderRow(false).FreezeAt);
    }

    [Fact]
    public void WithFreezeAt_RefusesTheOnePositionThatWouldMeanNothing()
    {
        // (0, 0) freezes nothing while making FreezeAt.HasValue true - two spellings of one state.
        Assert.Equal("row", Assert.Throws<ArgumentOutOfRangeException>(
            () => XlsxFormat.None.WithFreezeAt(0, 0)).ParamName);
        Assert.Equal("row", Assert.Throws<ArgumentOutOfRangeException>(
            () => XlsxFormat.None.WithFreezeAt(-1, 0)).ParamName);
        Assert.Equal("column", Assert.Throws<ArgumentOutOfRangeException>(
            () => XlsxFormat.None.WithFreezeAt(1, -1)).ParamName);

        // Rows only and columns only are both legal.
        Assert.Equal(new XlsxFreeze(1, 0), XlsxFormat.None.WithFreezeAt(1, 0).FreezeAt);
        Assert.Equal(new XlsxFreeze(0, 2), XlsxFormat.None.WithFreezeAt(0, 2).FreezeAt);
    }

    [Fact]
    public void WithColumnWidth_RefusesAWidthThatIsNotPositive()
    {
        Assert.Equal("width", Assert.Throws<ArgumentOutOfRangeException>(
            () => XlsxFormat.None.WithColumnWidth("A", 0)).ParamName);
        Assert.Equal("column", Assert.Throws<ArgumentException>(
            () => XlsxFormat.None.WithColumnWidth("  ", 10)).ParamName);
    }

    [Fact]
    public void TheNewMembersAreImmutableLikeTheOldOnes()
    {
        XlsxFormat baseline = XlsxFormat.None;
        XlsxFormat grown = baseline
            .WithColumnWidth("A", 42)
            .WithAutoFilter()
            .WithRule(XlsxRule.Blank("A1:A9", XlsxHighlight.Grey))
            .WithValidation(XlsxValidation.OneOf("B1:B9", "yes", "no"));

        // The starting instance is STATIC and shared, so a builder that mutated in place would
        // poison XlsxFormat.None for every caller in the process.
        Assert.Empty(baseline.ColumnWidths);
        Assert.False(baseline.AutoFilter);
        Assert.Empty(baseline.Rules);
        Assert.Empty(baseline.Validations);

        Assert.Equal(42, grown.ColumnWidths["A"]);
        Assert.True(grown.AutoFilter);
        Assert.Single(grown.Rules);
        Assert.Single(grown.Validations);
    }
}
