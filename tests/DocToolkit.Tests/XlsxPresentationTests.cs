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
        // The guard against FreezeHeaderRow becoming derived being a regression.
        //
        // The width is asserted RELATIONALLY, against auto-fit measured on this same machine,
        // rather than against the 54.14 the design opens with. That number is a property of the
        // HOST, not of this code: auto-fit derives a column width from font metrics, so a machine
        // with different fonts computes a different one. Pinning it passed on Linux and Windows
        // and failed on macOS at 50 - the same mistake CLAUDE.md already records for PDF byte
        // sizes, which vary ~100x with installed fonts and must never be asserted exactly.
        //
        // What the test is actually for survives intact: Report still auto-fits, and it still
        // freezes exactly the header row. Both are claims about this code.
        double autoFitOnly = Read(WorkbookEditor.Format(Sheet(), "Data",
            XlsxFormat.None.WithAutoFitColumns())).Width;
        var read = Read(WorkbookEditor.Format(Sheet(), "Data", XlsxFormat.Report));

        Assert.Equal(autoFitOnly, read.Width);
        Assert.True(read.Width > 20, $"Report must auto-fit column A, measured {read.Width}");
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
    // ---- the validation half, which a review found was essentially unmeasured -------------------

    /// <summary>
    /// Reads one validation back out of the SAVED bytes.
    /// </summary>
    /// <remarks>
    /// <b>Counting validations is not enough, and that was the gap.</b> A review mutated three
    /// things at once — the list kind to a text length, decimal to whole number, and a date range
    /// collapsed to a point — and the whole suite stayed green, because the only test counted five
    /// entries across five ranges. Kind, operator and bounds were all invisible.
    /// </remarks>
    private static ClosedXML.Excel.IXLDataValidation ReadValidation(byte[] xlsx, string range)
    {
        using var ms = new MemoryStream(xlsx);
        using var workbook = new ClosedXML.Excel.XLWorkbook(ms);
        return workbook.Worksheet("Data").DataValidations
            .Single(v => v.Ranges.Any(r => r.RangeAddress.ToStringRelative() == range));
    }

    [Fact]
    public void AWholeNumberValidationCarriesItsKindAndBothBounds()
    {
        ClosedXML.Excel.IXLDataValidation read = ReadValidation(
            WorkbookEditor.Format(Sheet(), "Data",
                XlsxFormat.None.WithValidation(XlsxValidation.WholeNumberBetween("B2:B3", 5, 250))),
            "B2:B3");

        Assert.Equal(ClosedXML.Excel.XLAllowedValues.WholeNumber, read.AllowedValues);
        Assert.Equal(ClosedXML.Excel.XLOperator.Between, read.Operator);
        Assert.Equal("5", read.MinValue);
        Assert.Equal("250", read.MaxValue);
    }

    [Fact]
    public void EachValidationKindReachesTheFileAsItsOwnKind()
    {
        // A single mapping that answered every kind the same way would satisfy a counting test.
        byte[] xlsx = WorkbookEditor.Format(Sheet(), "Data", XlsxFormat.None
            .WithValidation(XlsxValidation.WholeNumberBetween("B2:B3", 0, 10))
            .WithValidation(XlsxValidation.DecimalBetween("C2:C3", 0, 1))
            .WithValidation(XlsxValidation.TextLengthBetween("A2:A3", 1, 200))
            .WithValidation(XlsxValidation.DateBetween("D2:D3", new DateTime(2020, 1, 1), new DateTime(2030, 1, 1)))
            .WithValidation(XlsxValidation.OneOf("E2:E3", "Free", "Pro", "Team")));

        Assert.Equal(ClosedXML.Excel.XLAllowedValues.WholeNumber, ReadValidation(xlsx, "B2:B3").AllowedValues);
        Assert.Equal(ClosedXML.Excel.XLAllowedValues.Decimal, ReadValidation(xlsx, "C2:C3").AllowedValues);
        Assert.Equal(ClosedXML.Excel.XLAllowedValues.TextLength, ReadValidation(xlsx, "A2:A3").AllowedValues);
        Assert.Equal(ClosedXML.Excel.XLAllowedValues.Date, ReadValidation(xlsx, "D2:D3").AllowedValues);
        Assert.Equal(ClosedXML.Excel.XLAllowedValues.List, ReadValidation(xlsx, "E2:E3").AllowedValues);
    }

    [Fact]
    public void ADateValidationCarriesBothOfItsDates_NotJustOne()
    {
        ClosedXML.Excel.IXLDataValidation read = ReadValidation(
            WorkbookEditor.Format(Sheet(), "Data", XlsxFormat.None.WithValidation(
                XlsxValidation.DateBetween("D2:D3", new DateTime(2020, 1, 1), new DateTime(2030, 1, 1)))),
            "D2:D3");

        // A date lands as an Excel SERIAL number, not as text - 43831 is 2020-01-01. Converting it
        // back is what makes this assertion about the dates rather than about the encoding, and it
        // is why collapsing the range to a point cannot slip through.
        Assert.NotEqual(read.MinValue, read.MaxValue);
        Assert.Equal(new DateTime(2020, 1, 1), DateTime.FromOADate(double.Parse(read.MinValue,
            System.Globalization.CultureInfo.InvariantCulture)));
        Assert.Equal(new DateTime(2030, 1, 1), DateTime.FromOADate(double.Parse(read.MaxValue,
            System.Globalization.CultureInfo.InvariantCulture)));
    }

    [Fact]
    public void AListValidationCarriesEveryOptionAndNoMore()
    {
        ClosedXML.Excel.IXLDataValidation read = ReadValidation(
            WorkbookEditor.Format(Sheet(), "Data",
                XlsxFormat.None.WithValidation(XlsxValidation.OneOf("E2:E3", "Free", "Pro", "Team"))),
            "E2:E3");

        // The hand-rolled quoting on the apply side is the most fragile line in this feature, and
        // until now nothing exercised it at all.
        Assert.Equal(ClosedXML.Excel.XLAllowedValues.List, read.AllowedValues);
        Assert.Equal("\"Free,Pro,Team\"", read.MinValue);
    }

    [Fact]
    public void OneOf_RefusesAnOptionItCannotFaithfullyEncode()
    {
        // Measured before the guard: "Free, Pro" became TWO options in the file, and an option
        // containing a quote produced a malformed formula. For a type whose premise is a vocabulary
        // that can be guaranteed, accepting input it cannot encode is the wrong failure.
        Assert.Equal("options", Assert.Throws<ArgumentException>(
            () => XlsxValidation.OneOf("A1", "Free, Pro", "Other")).ParamName);
        Assert.Equal("options", Assert.Throws<ArgumentException>(
            () => XlsxValidation.OneOf("A1", "Say \"hi\"", "Other")).ParamName);
        Assert.Equal("options", Assert.Throws<ArgumentException>(
            () => XlsxValidation.OneOf("A1", null!, "Other")).ParamName);
        Assert.Equal("options", Assert.Throws<ArgumentException>(
            () => XlsxValidation.OneOf("A1", "  ", "Other")).ParamName);
    }

    // ---- the two rule conditions the operator test cannot reach ---------------------------------

    [Fact]
    public void ContainsAndBlankReachTheFileAsTheirOwnTYPES_NotJustOperators()
    {
        // Both read back as XLCFOperator.Equal, so the operator test cannot tell them apart - and a
        // review proved it: mapping Contains to WhenIsBlank passed the entire suite. The
        // discriminator is ConditionalFormatType, which is what the design's own table records.
        static ClosedXML.Excel.XLConditionalFormatType TypeOf(XlsxRule rule)
        {
            using var ms = new MemoryStream(
                WorkbookEditor.Format(Sheet(), "Data", XlsxFormat.None.WithRule(rule)));
            using var workbook = new ClosedXML.Excel.XLWorkbook(ms);
            return workbook.Worksheet("Data").ConditionalFormats.Single().ConditionalFormatType;
        }

        Assert.Equal(ClosedXML.Excel.XLConditionalFormatType.ContainsText,
            TypeOf(XlsxRule.Contains("A2:A3", "long", XlsxHighlight.Amber)));
        Assert.Equal(ClosedXML.Excel.XLConditionalFormatType.IsBlank,
            TypeOf(XlsxRule.Blank("A2:A3", XlsxHighlight.Grey)));
        Assert.Equal(ClosedXML.Excel.XLConditionalFormatType.CellIs,
            TypeOf(XlsxRule.GreaterThan("B2:B3", 1, XlsxHighlight.Red)));
    }

    [Fact]
    public void TheAutoFilterCoversTheUsedRangeRatherThanOneCell()
    {
        // A one-cell autofilter is useless in Excel, and asserting IsEnabled alone cannot see it -
        // measured: narrowing the filter to A1:A1 passed.
        using var ms = new MemoryStream(
            WorkbookEditor.Format(Sheet(), "Data", XlsxFormat.None.WithAutoFilter()));
        using var workbook = new ClosedXML.Excel.XLWorkbook(ms);
        ClosedXML.Excel.IXLAutoFilter filter = workbook.Worksheet("Data").AutoFilter;

        Assert.True(filter.IsEnabled);
        Assert.Equal("A1:B3", filter.Range.RangeAddress.ToStringRelative());
    }

    [Fact]
    public void AnAutoFilterOnASheetWithNoDataIsSkippedRatherThanThrowing()
    {
        // The empty-sheet case lived only in an implementation comment. It is a documented promise
        // on XlsxFormat.AutoFilter now, so it needs a test.
        byte[] empty = WorkbookEditor.Create("Data", []);

        byte[] formatted = WorkbookEditor.Format(empty, "Data", XlsxFormat.None.WithAutoFilter());

        using var ms = new MemoryStream(formatted);
        using var workbook = new ClosedXML.Excel.XLWorkbook(ms);
        Assert.False(workbook.Worksheet("Data").AutoFilter.IsEnabled);
    }

    [Fact]
    public void TheNewCollectionsCannotBeMutatedThroughTheirRuntimeType()
    {
        // The other half of the immutability guard: XlsxFormat.None and .Report are STATIC, so a
        // cast-back-and-write would poison them process-wide. The existing test covers only
        // ColumnNumberFormats; these three were wrapped correctly but unpinned.
        Assert.Throws<NotSupportedException>(
            () => ((IDictionary<string, double>)XlsxFormat.None.ColumnWidths)["A"] = 99);
        Assert.Throws<NotSupportedException>(
            () => ((IList<XlsxRule>)XlsxFormat.None.Rules).Add(XlsxRule.Blank("A1", XlsxHighlight.Grey)));
        Assert.Throws<NotSupportedException>(
            () => ((IList<XlsxValidation>)XlsxFormat.None.Validations).Add(XlsxValidation.OneOf("A1", "x")));
        Assert.Throws<NotSupportedException>(
            () => ((IList<string>)XlsxValidation.OneOf("A1", "x").Options).Add("y"));
    }
    // ---- the vocabulary is CLOSED, which means it must refuse what falls outside it ------------

    // The two enums are NOT equally reachable, and it is worth being exact about which is which.
    //
    //   XlsxHighlight     is a PARAMETER of every factory, so (XlsxHighlight)99 is reachable from
    //                     a consumer and is tested below.
    //   XlsxRuleKind      is set by the factory itself and XlsxValidationKind likewise, so an
    //                     out-of-range value cannot be constructed through the public API at all.
    //
    // Their throwing arms are therefore DEFENSIVE, and Stryker will report them uncovered. That is
    // honest rather than a gap: the alternative was a `_` arm that silently answered Blank, and an
    // unreachable throw is better than a reachable wrong answer. Recorded here so a future reader
    // does not "fix" the coverage by making the kind settable.

    [Fact]
    public void AnUndefinedHighlightDoesNotSilentlyBecomeGrey()
    {
        // The specific measurement: (XlsxHighlight)99 came back FFD3D3D3 - the same colour a
        // DELIBERATE Grey produces. An out-of-range cast was indistinguishable from a real choice.
        byte[] grey = WorkbookEditor.Format(Sheet(), "Data",
            XlsxFormat.None.WithRule(XlsxRule.Blank("A2:A3", XlsxHighlight.Grey)));
        Assert.NotEmpty(grey);

        Assert.Throws<DocumentConversionException>(() => WorkbookEditor.Format(Sheet(), "Data",
            XlsxFormat.None.WithRule(XlsxRule.Blank("A2:A3", (XlsxHighlight)99))));
    }

    [Theory]
    [InlineData("Other!A2:B2")]
    [InlineData("'Other'!A2:B2")]
    [InlineData("NoSuchSheet!A2:B2")]
    public void ASheetQualifiedRangeIsRefusedRatherThanSilentlyRetargeted(string range)
    {
        // MEASURED on a two-sheet workbook before the guard: every one of these applied to the
        // sheet Format names, with no error - the qualifier was discarded, not honoured. Even
        // "NoSuchSheet!" landed on Data. A caller who writes a qualifier means it.
        Assert.Equal("range", Assert.Throws<ArgumentException>(
            () => XlsxRule.GreaterThan(range, 1, XlsxHighlight.Red)).ParamName);
        Assert.Equal("range", Assert.Throws<ArgumentException>(
            () => XlsxValidation.WholeNumberBetween(range, 1, 2)).ParamName);
    }

    [Fact]
    public void AnUnqualifiedRangeIsStillAccepted()
    {
        // The control. Without this, a guard that refused EVERY range would pass the theory above.
        Assert.Equal("A2:B2", XlsxRule.GreaterThan("A2:B2", 1, XlsxHighlight.Red).Range);
        Assert.Equal("A2:B2", XlsxValidation.WholeNumberBetween("A2:B2", 1, 2).Range);
    }
    // ---- aimed at the survivor list, not written blind ----------------------------------------
    //
    // Widening the Stryker scope to XlsxRule.cs and XlsxValidation.cs dropped the score to
    // 93.40%, under the 95 break threshold. CLAUDE.md records exactly this shape from the last
    // widening - "adding them and lowering the gate would have bought a green run and no
    // information" - so these kill the survivors instead. Each names the mutant it exists for.

    [Fact]
    public void ADegenerateRangeIsAccepted_BecauseLowEqualToHighIsAValidRule()
    {
        // Kills `high < low` -> `high <= low` in XlsxRule, and the two `max < min` siblings in
        // XlsxValidation. A single permitted value is a real thing to ask for; the guard exists
        // to catch INVERTED bounds, not equal ones, and nothing pinned that difference.
        Assert.Equal(5, XlsxRule.Between("A1:A9", 5, 5, XlsxHighlight.Red).Value);
        Assert.Equal(5, XlsxValidation.WholeNumberBetween("A1:A9", 5, 5).Min);

        var day = new DateTime(2020, 1, 1);
        Assert.Equal(day, XlsxValidation.DateBetween("A1:A9", day, day).MinDate);

        // And the control: genuinely inverted bounds must still be refused, or the assertions
        // above would pass against a guard that had been deleted outright.
        Assert.Throws<ArgumentOutOfRangeException>(() => XlsxRule.Between("A1:A9", 9, 1, XlsxHighlight.Red));
        Assert.Throws<ArgumentOutOfRangeException>(() => XlsxValidation.WholeNumberBetween("A1:A9", 9, 1));
        Assert.Throws<ArgumentOutOfRangeException>(
            () => XlsxValidation.DateBetween("A1:A9", day.AddDays(1), day));
    }

    [Fact]
    public void EveryFactoryRefusesNullWithTheRightExceptionTYPE_NotWhateverDereferencingThrows()
    {
        // Kills the ThrowIfNull statement removals in XlsxRule.EqualTo, XlsxValidation.OneOf and
        // both XlsxFormat.With* methods. The TYPE is the assertion: without the guard the call
        // still fails, but as a NullReferenceException from a dereference further in, which is
        // the difference between a contract and an accident.
        Assert.Throws<ArgumentNullException>(() => XlsxRule.EqualTo("A1", null!, XlsxHighlight.Red));
        Assert.Throws<ArgumentNullException>(() => XlsxRule.Contains("A1", null!, XlsxHighlight.Red));
        Assert.Throws<ArgumentNullException>(() => XlsxValidation.OneOf("A1", (string[])null!));
        Assert.Throws<ArgumentNullException>(() => XlsxFormat.None.WithRule(null!));
        Assert.Throws<ArgumentNullException>(() => XlsxFormat.None.WithValidation(null!));
    }

    [Fact]
    public void ABlankRangeIsRefusedByBothVocabularies()
    {
        // Kills the ThrowIfNullOrWhiteSpace removal. A blank range otherwise sails through the
        // sheet-qualifier check - it contains no '!' - and is stored, failing much later inside
        // ClosedXML where the caller's argument no longer has a name.
        foreach (string blank in new[] { "", "   ", "\t" })
        {
            Assert.Throws<ArgumentException>(() => XlsxRule.Blank(blank, XlsxHighlight.Grey));
            Assert.Throws<ArgumentException>(() => XlsxValidation.OneOf(blank, "x"));
        }

        // Null, on BOTH vocabularies, naming the caller's argument. XlsxValidation had no
        // null-range test at all - which matters because the ThrowIfNull(range) line above it
        // is EXCLUDED from mutation as equivalent, and an exclusion resting on a test nobody
        // wrote is not a measurement. With these here, that mutant surviving means the
        // behaviour really is identical rather than merely unobserved.
        Assert.Equal("range", Assert.Throws<ArgumentNullException>(
            () => XlsxValidation.OneOf(null!, "x")).ParamName);
        Assert.Equal("range", Assert.Throws<ArgumentNullException>(
            () => XlsxValidation.WholeNumberBetween(null!, 1, 2)).ParamName);
        Assert.Equal("range", Assert.Throws<ArgumentNullException>(
            () => XlsxValidation.DateBetween(null!, DateTime.Today, DateTime.Today)).ParamName);
    }

    [Fact]
    public void AFreezeSurvivesALaterUnrelatedSetting()
    {
        // Kills both null-coalescing mutants on the private With helper's freeze argument. Every
        // With* method funnels through it, so dropping the `?? FreezeAt` fallback silently clears
        // a freeze whenever any OTHER setting is added afterwards - and nothing saw it, because
        // every existing test set the freeze last.
        XlsxFormat format = XlsxFormat.None
            .WithFreezeAt(3, 2)
            .WithAutoFilter()
            .WithColumnWidth("A", 15)
            .WithRule(XlsxRule.Blank("A1", XlsxHighlight.Grey));

        Assert.Equal(new XlsxFreeze(3, 2), format.FreezeAt);
    }

    [Fact]
    public void TheTwoMessagesThatARE_theFeature_SayWhyRatherThanJustNo()
    {
        // Kills the string mutants on these two specifically, and only these two. Most exception
        // text is left unpinned on purpose - asserting every message is brittle and this
        // repository already treats message mutants as an accepted floor. These are different:
        // each REFUSES something that used to be silently accepted, so the message carrying the
        // reason is the entire value of the guard. A caller who reads "no" and not "why" will
        // reasonably conclude the library is broken.
        var qualified = Assert.Throws<ArgumentException>(
            () => XlsxRule.GreaterThan("Other!A1:B2", 1, XlsxHighlight.Red));
        Assert.Contains("silently discarded", qualified.Message, StringComparison.Ordinal);
        Assert.Contains("sheetName", qualified.Message, StringComparison.Ordinal);

        var comma = Assert.Throws<ArgumentException>(() => XlsxValidation.OneOf("A1", "Free, Pro"));
        Assert.Contains("comma", comma.Message, StringComparison.Ordinal);
        Assert.Contains("corrupt it silently", comma.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnInstructionalMessageKeepsItsInSTRUCTION()
    {
        // The other half of the same judgement. These messages do not merely report a refusal -
        // each NAMES THE THING TO DO INSTEAD, which is the part a caller acts on and the part a
        // string mutant silently deletes. Pinned by distinctive phrase rather than whole text, so
        // rewording stays cheap while gutting does not.
        //
        // Deliberately NOT pinned: "Number format was blank", "A column width must be positive"
        // and their kind. Those restate the guard and a caller learns nothing from the words that
        // the exception type and parameter name have not already told them. Their mutants stay in
        // the accepted floor CLAUDE.md already describes - asserting every message is how a suite
        // becomes brittle, and the distinction here is whether the sentence carries information.
        var nothing = Assert.Throws<ArgumentOutOfRangeException>(() => XlsxFormat.None.WithFreezeAt(0, 0));
        Assert.Contains("WithFrozenHeaderRow(false)", nothing.Message, StringComparison.Ordinal);

        var negativeRow = Assert.Throws<ArgumentOutOfRangeException>(() => XlsxFormat.None.WithFreezeAt(-1, 0));
        Assert.Contains("cannot be negative", negativeRow.Message, StringComparison.Ordinal);

        // The two must not share a message: (0,0) is "you asked for nothing, here is how to say
        // that on purpose" and -1 is "that is not a position". Collapsing them was a review
        // finding earlier in this branch, and only distinct assertions keep them apart.
        Assert.DoesNotContain("WithFrozenHeaderRow", negativeRow.Message, StringComparison.Ordinal);

        var negativeColumn = Assert.Throws<ArgumentOutOfRangeException>(() => XlsxFormat.None.WithFreezeAt(1, -1));
        Assert.Contains("cannot be negative", negativeColumn.Message, StringComparison.Ordinal);
        Assert.Equal("column", negativeColumn.ParamName);

        var empty = Assert.Throws<ArgumentException>(() => XlsxValidation.OneOf("A1"));
        Assert.Contains("a cell nobody can fill", empty.Message, StringComparison.Ordinal);

        var blankOption = Assert.Throws<ArgumentException>(() => XlsxValidation.OneOf("A1", "x", "  "));
        Assert.Contains("nobody can pick", blankOption.Message, StringComparison.Ordinal);
    }
    [Fact]
    public void TheChainTheNugetReadmeShows_ActuallyCompilesAndRuns()
    {
        // src/DocToolkit/README.md is what nuget.org renders, so its snippets are the first code
        // a consumer copies. Its "Usage" fence says outright that it is a connected walkthrough
        // rather than a script that compiles as pasted - placeholders like logoBytes stand in for
        // the reader's own data - so gen-readme-snippets.py does not manage it and nothing
        // compiled the XlsxFormat chain that was added to it.
        //
        // A narrative fence is a fair reason for the WHOLE block not to compile. It is not a
        // reason for the individual calls to be unchecked: a wrong argument order or a renamed
        // parameter would ship to nuget.org looking authoritative. This runs the exact chain.
        byte[] report = WorkbookEditor.Format(Sheet(), "Data", XlsxFormat.Report
            .WithNumberFormat("B", "#,##0.00")
            .WithColumnWidth("A", 42)
            .WithFreezeAt(row: 2, column: 1)
            .WithAutoFilter()
            .WithRule(XlsxRule.GreaterThan("B2:B999", 10000, XlsxHighlight.Red))
            .WithValidation(XlsxValidation.OneOf("C2:C999", "Free", "Pro", "Team")));

        // Read back the two the chain would most easily get wrong: the explicit width must beat
        // Report's own auto-fit, and the named freeze must beat Report's header-row freeze.
        var read = Read(report);
        Assert.Equal(42, read.Width);
        Assert.Equal(2, read.FrozenRows);
        Assert.Equal(1, read.FrozenColumns);
        Assert.True(read.Filter);
        Assert.Equal(1, read.Rules);
        Assert.Equal(1, read.Validations);
    }
    [Fact]
    public void SettingAFreezeTwiceKeepsTheLastOne_NotTheFirst()
    {
        // Kills the surviving `freezeAt ?? FreezeAt` -> `FreezeAt ?? freezeAt` mutant. The two
        // differ ONLY when both are non-null, which is exactly the overwrite case - and every
        // other freeze test sets the position once, so nothing reached it. Under the mutant a
        // second WithFreezeAt is silently ignored, which is a wrong answer rather than a refusal.
        XlsxFormat twice = XlsxFormat.None.WithFreezeAt(1, 1).WithFreezeAt(3, 2);

        Assert.Equal(new XlsxFreeze(3, 2), twice.FreezeAt);
    }

    [Fact]
    public void BothVocabulariesExplainTheSheetQualifierInFull_NotJustTheFirstHalf()
    {
        // XlsxValidation carries its own copy of the refusal message and the theory above asserts
        // only ParamName, so every string in that copy survived. The message spans three literals
        // and the LAST one names the fix, so asserting the opening phrase alone leaves the useful
        // half unpinned - which is how a message decays into "no" with the "why" quietly gone.
        foreach (string message in new[]
        {
            Assert.Throws<ArgumentException>(
                () => XlsxRule.GreaterThan("Other!A1:B2", 1, XlsxHighlight.Red)).Message,
            Assert.Throws<ArgumentException>(
                () => XlsxValidation.WholeNumberBetween("Other!A1:B2", 1, 2)).Message,
        })
        {
            Assert.Contains("silently discarded", message, StringComparison.Ordinal);
            Assert.Contains("Pass the range alone", message, StringComparison.Ordinal);
            Assert.Contains("sheetName parameter chooses", message, StringComparison.Ordinal);
            Assert.Contains("the sheet.", message, StringComparison.Ordinal);
        }
    }
}
