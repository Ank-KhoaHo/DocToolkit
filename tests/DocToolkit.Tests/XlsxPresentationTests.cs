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
    // ---- applied, and read back out of the SAVED BYTES ------------------------------------------

    private static byte[] Sheet() => WorkbookEditor.Create("Data",
    [
        ["Description", "Amount"],
        ["A rather long description that would otherwise be truncated", 1234.5],
        ["short", 9.0],
    ]);

    /// <summary>
    /// Reads the presentation back out of the SAVED bytes. Asserting on the <see cref="XlsxFormat"/>
    /// object would pass against an <c>ApplyFormat</c> that discarded every one of these.
    /// </summary>
    private static (double Width, int Rules, int Validations, bool Filter, int FrozenRows, int FrozenColumns, string Colour)
        Read(byte[] xlsx)
    {
        using var ms = new MemoryStream(xlsx);
        using var workbook = new ClosedXML.Excel.XLWorkbook(ms);
        ClosedXML.Excel.IXLWorksheet sheet = workbook.Worksheet("Data");

        string colour = sheet.ConditionalFormats.Any()
            ? sheet.ConditionalFormats.First().Style.Fill.BackgroundColor.ToString()
            : string.Empty;

        return (Math.Round(sheet.Column(1).Width, 2),
                sheet.ConditionalFormats.Count(),
                sheet.DataValidations.Count(),
                sheet.AutoFilter is { IsEnabled: true },
                sheet.SheetView.SplitRow,
                sheet.SheetView.SplitColumn,
                colour);
    }

    [Fact]
    public void ApplyingNothing_ProducesASheetWithNoneOfIt()
    {
        // The positive control. Without it, an ApplyFormat that always added a rule would pass every
        // assertion below.
        var read = Read(WorkbookEditor.Format(Sheet(), "Data", XlsxFormat.None));

        Assert.Equal(0, read.Rules);
        Assert.Equal(0, read.Validations);
        Assert.False(read.Filter);
        Assert.Equal(0, read.FrozenRows);
    }

    [Fact]
    public void AConditionalFormatReachesTheSavedWorkbook_WithItsColour()
    {
        // The colour matters as much as the rule: one that reloads without its formatting is a rule
        // nobody can see.
        var read = Read(WorkbookEditor.Format(Sheet(), "Data",
            XlsxFormat.None.WithRule(XlsxRule.GreaterThan("B2:B3", 100, XlsxHighlight.Red))));

        Assert.Equal(1, read.Rules);
        Assert.Equal("FFFF0000", read.Colour);
    }

    [Fact]
    public void EveryHighlightProducesADistinctColour()
    {
        // Four intents must not collapse onto one colour - a mapping that returned red for
        // everything would satisfy the test above. Asserted as a SET, so any two colliding fails.
        string[] colours =
        [
            .. new[] { XlsxHighlight.Red, XlsxHighlight.Amber, XlsxHighlight.Green, XlsxHighlight.Grey }
                .Select(h => Read(WorkbookEditor.Format(Sheet(), "Data",
                    XlsxFormat.None.WithRule(XlsxRule.Blank("A2:A3", h)))).Colour)
        ];

        Assert.Equal(4, colours.Distinct(StringComparer.Ordinal).Count());
        Assert.DoesNotContain(string.Empty, colours);
    }

    [Fact]
    public void EveryRuleKindReachesTheSavedWorkbook()
    {
        XlsxFormat format = XlsxFormat.None
            .WithRule(XlsxRule.GreaterThan("B2:B3", 100, XlsxHighlight.Red))
            .WithRule(XlsxRule.LessThan("B2:B3", 5000, XlsxHighlight.Green))
            .WithRule(XlsxRule.Between("B2:B3", 1, 10000, XlsxHighlight.Amber))
            .WithRule(XlsxRule.EqualTo("A2:A3", "short", XlsxHighlight.Grey))
            .WithRule(XlsxRule.Contains("A2:A3", "long", XlsxHighlight.Amber))
            .WithRule(XlsxRule.Blank("A2:A3", XlsxHighlight.Grey));

        Assert.Equal(6, Read(WorkbookEditor.Format(Sheet(), "Data", format)).Rules);
    }

    [Fact]
    public void EveryValidationKindReachesTheSavedWorkbook()
    {
        XlsxFormat format = XlsxFormat.None
            .WithValidation(XlsxValidation.WholeNumberBetween("B2:B3", 0, 10000))
            .WithValidation(XlsxValidation.DecimalBetween("C2:C3", 0, 1))
            .WithValidation(XlsxValidation.TextLengthBetween("A2:A3", 1, 200))
            .WithValidation(XlsxValidation.DateBetween("D2:D3", new DateTime(2020, 1, 1), new DateTime(2030, 1, 1)))
            .WithValidation(XlsxValidation.OneOf("E2:E3", "Free", "Pro", "Team"));

        Assert.Equal(5, Read(WorkbookEditor.Format(Sheet(), "Data", format)).Validations);
    }

    [Fact]
    public void AnAutoFilterReachesTheSavedWorkbook()
    {
        Assert.True(Read(WorkbookEditor.Format(Sheet(), "Data", XlsxFormat.None.WithAutoFilter())).Filter);
    }

    [Fact]
    public void AnExplicitWidthBeatsAutoFitForTheColumnItNames()
    {
        // The composition rule. Auto-fit alone widens column A well past 20; the explicit width must
        // then win, so applying them in the wrong order fails here.
        double autoFitted = Read(WorkbookEditor.Format(Sheet(), "Data",
            XlsxFormat.None.WithAutoFitColumns())).Width;
        Assert.True(autoFitted > 20, $"auto-fit should widen column A, measured {autoFitted}");

        var read = Read(WorkbookEditor.Format(Sheet(), "Data",
            XlsxFormat.None.WithAutoFitColumns().WithColumnWidth("A", 12)));

        Assert.Equal(12, read.Width);
    }

    [Fact]
    public void AFreezePositionReachesTheSavedWorkbook()
    {
        var read = Read(WorkbookEditor.Format(Sheet(), "Data", XlsxFormat.None.WithFreezeAt(3, 2)));

        Assert.Equal(3, read.FrozenRows);
        Assert.Equal(2, read.FrozenColumns);
    }

    [Fact]
    public void XlsxFormatReport_StillDoesExactlyWhatItDidBefore()
    {
        // The guard against FreezeHeaderRow becoming derived being a regression. These are the same
        // numbers the design opens with, measured before any of this existed.
        var read = Read(WorkbookEditor.Format(Sheet(), "Data", XlsxFormat.Report));

        Assert.Equal(54.14, read.Width);
        Assert.Equal(1, read.FrozenRows);
        Assert.Equal(0, read.FrozenColumns);
    }
    [Fact]
    public void ARuleCarriesBothOfItsBoundsAndItsOperator_NotJustItsExistence()
    {
        // Sabotage found the gap this closes: rewriting WhenBetween(low, high) to
        // WhenBetween(low, low) survived every other test here, because they COUNT rules rather
        // than reading them. A rule with collapsed bounds is a rule that never fires - the sheet
        // looks formatted and is not, which is the failure XlsxRule.Between's own guard exists for.
        byte[] xlsx = WorkbookEditor.Format(Sheet(), "Data",
            XlsxFormat.None.WithRule(XlsxRule.Between("B2:B3", 100, 300, XlsxHighlight.Amber)));

        using var ms = new MemoryStream(xlsx);
        using var workbook = new ClosedXML.Excel.XLWorkbook(ms);
        ClosedXML.Excel.IXLConditionalFormat rule =
            workbook.Worksheet("Data").ConditionalFormats.Single();

        Assert.Equal(ClosedXML.Excel.XLCFOperator.Between, rule.Operator);
        Assert.Equal("100", rule.Values[1].Value);
        Assert.Equal("300", rule.Values[2].Value);
    }

    [Fact]
    public void EachComparisonReachesTheFileAsItsOwnOperator()
    {
        // The counting test would also pass if every kind mapped to the same comparison. This
        // pins that the six conditions stay six distinct operators in the saved file.
        (XlsxRule Rule, ClosedXML.Excel.XLCFOperator Expected)[] cases =
        [
            (XlsxRule.GreaterThan("B2:B3", 1, XlsxHighlight.Red), ClosedXML.Excel.XLCFOperator.GreaterThan),
            (XlsxRule.LessThan("B2:B3", 1, XlsxHighlight.Red), ClosedXML.Excel.XLCFOperator.LessThan),
            (XlsxRule.EqualTo("A2:A3", "short", XlsxHighlight.Red), ClosedXML.Excel.XLCFOperator.Equal),
        ];

        foreach ((XlsxRule rule, ClosedXML.Excel.XLCFOperator expected) in cases)
        {
            using var ms = new MemoryStream(
                WorkbookEditor.Format(Sheet(), "Data", XlsxFormat.None.WithRule(rule)));
            using var workbook = new ClosedXML.Excel.XLWorkbook(ms);

            Assert.Equal(expected, workbook.Worksheet("Data").ConditionalFormats.Single().Operator);
        }
    }
}
