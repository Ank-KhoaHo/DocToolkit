using System.Reflection;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// Covers <see cref="DocxForm"/> and <see cref="DocxFormValue"/> — Word content controls as a
/// fill-in form.
///
/// <b>Several of these tests pin the behaviour of the library underneath rather than ours</b>, and
/// deliberately: the design rests on measurements taken 2026-08-26, and an upstream change that
/// moved one would otherwise surface as a caller's form quietly filling wrong.
/// </summary>
public class DocxFormTests
{
    [Fact]
    public void DocxFormValue_CarriesEachKindAndLeavesTheOthersNull()
    {
        var text = DocxFormValue.FromText("hello");
        Assert.Equal(DocxFormValueKind.Text, text.Kind);
        Assert.Equal("hello", text.Text);
        Assert.Null(text.Checked);
        Assert.Null(text.Bytes);

        var ticked = DocxFormValue.FromChecked(true);
        Assert.Equal(DocxFormValueKind.Checked, ticked.Kind);
        Assert.True(ticked.Checked);
        Assert.Null(ticked.Text);

        var when = DocxFormValue.FromDate(new DateTime(2026, 3, 9));
        Assert.Equal(DocxFormValueKind.Date, when.Kind);
        Assert.Equal(new DateTime(2026, 3, 9), when.Date);

        var choice = DocxFormValue.FromChoice("Pro");
        Assert.Equal(DocxFormValueKind.Choice, choice.Kind);
        Assert.Equal("Pro", choice.Text);

        var picture = DocxFormValue.FromPicture([1, 2, 3], "logo.png");
        Assert.Equal(DocxFormValueKind.Picture, picture.Kind);
        Assert.Equal<byte[]>([1, 2, 3], picture.Bytes!);
        Assert.Equal("logo.png", picture.FileName);
        Assert.Null(picture.Text);
    }

    /// <summary>
    /// The bytes-only guarantee, asserted STRUCTURALLY rather than by behaviour.
    ///
    /// Upstream's <c>WordContentControlPictureValue</c> has a <c>FromFile(path)</c> factory that
    /// reads the disk. This package refuses to read local files for images, so that route must not
    /// be reachable through this API at all. A behavioural test cannot prove a negative; a surface
    /// test can, and it fails the moment somebody adds a sixth factory.
    /// </summary>
    [Fact]
    public void DocxFormValue_ExposesExactlyFiveFactories_AndNoneTakesAUri()
    {
        MethodInfo[] factories = typeof(DocxFormValue)
            .GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly)
            .Where(m => !m.IsSpecialName)
            .ToArray();

        Assert.Equal(
            ["FromChecked", "FromChoice", "FromDate", "FromPicture", "FromText"],
            factories.Select(m => m.Name).OrderBy(n => n, StringComparer.Ordinal));

        foreach (MethodInfo factory in factories)
            Assert.DoesNotContain(factory.GetParameters(), p => p.ParameterType == typeof(Uri));
    }

    [Fact]
    public void DocxFormValue_RefusesNullAndEmptyByTheParameterItDeclares()
    {
        Assert.Equal("value", Assert.Throws<ArgumentNullException>(
            () => DocxFormValue.FromText(null!)).ParamName);
        Assert.Equal("value", Assert.Throws<ArgumentNullException>(
            () => DocxFormValue.FromChoice(null!)).ParamName);
        Assert.Equal("bytes", Assert.Throws<ArgumentNullException>(
            () => DocxFormValue.FromPicture(null!, "a.png")).ParamName);
        Assert.Equal("bytes", Assert.Throws<ArgumentException>(
            () => DocxFormValue.FromPicture([], "a.png")).ParamName);
        Assert.Equal("fileName", Assert.Throws<ArgumentException>(
            () => DocxFormValue.FromPicture([1], "  ")).ParamName);
    }
    [Fact]
    public void DocxFormIssueKind_NamesOnlyTheThreeKindsThatCanActuallyFire()
    {
        // Upstream reports NINE. Measured 2026-08-26, three can be provoked: a drop-down value
        // outside its list, a string where a date belongs and a bool where a date belongs all
        // reported valid=True. Naming a kind a caller cannot receive advertises a check that does
        // not run - the failure A67 was filed to avoid - so the other six map to Other.
        Assert.Equal(
            ["Other", "MissingValue", "UnusedValue", "DuplicateKey"],
            Enum.GetNames<DocxFormIssueKind>());
    }

    [Fact]
    public void DocxFormKey_DefaultsToTheModeThatFallsBack()
    {
        // The default must fall back, so a template keyed either way works without the caller
        // knowing which. Measured: Tag gives "FullName" where Alias gives "Full name".
        Assert.Equal(DocxFormKey.TagThenAlias, default(DocxFormKey));
    }
    // ---- Inspect -------------------------------------------------------------------------------

    [Fact]
    public void Inspect_ReadsEveryControlAndItsCurrentValue()
    {
        DocxFormReport report = DocxForm.Inspect(DocxFormFixtures.Form());

        Assert.Equal(4, report.Fields.Count);
        Assert.Equal("Khoa Ho", Assert.Single(report.Fields, f => f.Key == "FullName").Value.Text);
    }

    [Theory]
    [InlineData(DocxFormKey.TagThenAlias, "FullName")]
    [InlineData(DocxFormKey.Alias, "Full name")]
    public void Inspect_KeysByTheModeAsked(DocxFormKey key, string expected)
    {
        // Measured: the two modes return genuinely different keys, which is why the mode is a
        // parameter rather than defaulted away.
        DocxFormReport report = DocxForm.Inspect(DocxFormFixtures.Form(), key);

        Assert.Contains(report.Fields, f => f.Key == expected);
    }

    // ---- Validate ------------------------------------------------------------------------------

    [Fact]
    public void Validate_ReportsAMissingControlAndAnUnusedValue()
    {
        DocxFormValidation result = DocxForm.Validate(DocxFormFixtures.Form(),
            new Dictionary<string, DocxFormValue>
            {
                ["FullName"] = DocxFormValue.FromText("Someone Else"),
                ["Nonexistent"] = DocxFormValue.FromText("spare"),
            });

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i => i.Kind == DocxFormIssueKind.MissingValue && i.Key == "Plan");
        Assert.Contains(result.Issues, i => i.Kind == DocxFormIssueKind.UnusedValue && i.Key == "Nonexistent");
        Assert.Contains("FullName", result.ExpectedKeys);
        Assert.Contains("Nonexistent", result.SuppliedKeys);
    }

    [Fact]
    public void Validate_ReportsTwoControlsSharingAName()
    {
        byte[] docx = DocxFormFixtures.Build(
            DocxFormFixtures.Control("Same", "Same", "one"),
            DocxFormFixtures.Control("Same", "Same", "two"));

        DocxFormValidation result = DocxForm.Validate(docx,
            new Dictionary<string, DocxFormValue> { ["Same"] = DocxFormValue.FromText("x") });

        Assert.False(result.IsValid);
        Assert.Equal(DocxFormIssueKind.DuplicateKey, Assert.Single(result.Issues).Kind);
    }

    [Fact]
    public void Validate_OnACompleteExactMatch_IsValid()
    {
        // The control that stops every assertion above passing vacuously.
        DocxFormValidation result = DocxForm.Validate(DocxFormFixtures.Form(),
            new Dictionary<string, DocxFormValue>
            {
                ["FullName"] = DocxFormValue.FromText("a"),
                ["Plan"] = DocxFormValue.FromChoice("Team"),
                ["Start"] = DocxFormValue.FromDate(new DateTime(2027, 3, 9)),
                ["Notes"] = DocxFormValue.FromText("b"),
            });

        Assert.True(result.IsValid);
        Assert.Empty(result.Issues);
    }

    // ---- Fill ----------------------------------------------------------------------------------

    [Fact]
    public void Fill_WritesTheValuesAndInspectReadsThemBack()
    {
        // The round trip is the point: what Inspect hands back is what Fill takes.
        byte[] filled = DocxForm.Fill(DocxFormFixtures.Form(),
            new Dictionary<string, DocxFormValue>
            {
                ["FullName"] = DocxFormValue.FromText("Someone Else"),
                ["Notes"] = DocxFormValue.FromText("updated"),
            });

        DocxFormReport report = DocxForm.Inspect(filled);
        Assert.Equal("Someone Else", Assert.Single(report.Fields, f => f.Key == "FullName").Value.Text);
        Assert.Equal("updated", Assert.Single(report.Fields, f => f.Key == "Notes").Value.Text);
    }

    [Fact]
    public void Fill_IsLenient_AndLeavesUnsuppliedControlsAtTheirOwnText()
    {
        // The measured behaviour the lenient decision rests on. Unlike a mail-merge field, a content
        // control that receives no value keeps its EXISTING content - there is no injected marker -
        // so a partial fill is a legitimate workflow rather than a half-finished document.
        byte[] filled = DocxForm.Fill(DocxFormFixtures.Form(),
            new Dictionary<string, DocxFormValue> { ["FullName"] = DocxFormValue.FromText("Only This") });

        DocxFormReport report = DocxForm.Inspect(filled);
        Assert.Equal("Only This", Assert.Single(report.Fields, f => f.Key == "FullName").Value.Text);
        Assert.Equal("Pro", Assert.Single(report.Fields, f => f.Key == "Plan").Value.Text);
        Assert.Equal("none", Assert.Single(report.Fields, f => f.Key == "Notes").Value.Text);
    }

    [Fact]
    public void Fill_LeavesTheDocumentReadableByExtractText()
    {
        // Fill rewrites the same document ExtractText was just taught to read through content
        // controls. If filling flattened or relocated a control the text would still be PRESENT and
        // every other assertion here would pass - so assert the SEPARATORS, which are what
        // distinguish a preserved structure from a rearranged one.
        byte[] filled = DocxForm.Fill(DocxFormFixtures.Form(),
            new Dictionary<string, DocxFormValue> { ["FullName"] = DocxFormValue.FromText("A") });

        Assert.Equal("A\nPro\n15 January 2026\nnone", DocxEditor.ExtractText(filled));
    }

    [Fact]
    public void Fill_RefusesANullValueInTheDictionary()
    {
        var ex = Assert.Throws<ArgumentException>(() => DocxForm.Fill(DocxFormFixtures.Form(),
            new Dictionary<string, DocxFormValue> { ["FullName"] = null! }));

        Assert.Equal("values", ex.ParamName);
        Assert.Contains("FullName", ex.Message, StringComparison.Ordinal);
    }

    // ---- guards --------------------------------------------------------------------------------

    [Fact]
    public void EveryByteArrayOverload_RefusesNullAndEmptyByTheParameterItDeclares()
    {
        var values = new Dictionary<string, DocxFormValue>();

        Assert.Equal("docx", Assert.Throws<ArgumentNullException>(() => DocxForm.Inspect(null!)).ParamName);
        Assert.Equal("docx", Assert.Throws<ArgumentException>(() => DocxForm.Inspect([])).ParamName);
        Assert.Equal("docx", Assert.Throws<ArgumentNullException>(() => DocxForm.Validate(null!, values)).ParamName);
        Assert.Equal("docx", Assert.Throws<ArgumentNullException>(() => DocxForm.Fill(null!, values)).ParamName);
        Assert.Equal("values", Assert.Throws<ArgumentNullException>(
            () => DocxForm.Fill(DocxFormFixtures.Form(), null!)).ParamName);
    }

    [Fact]
    public void EveryOverload_WrapsAnUnreadableDocument()
    {
        byte[] rubbish = [1, 2, 3, 4];
        var values = new Dictionary<string, DocxFormValue>();

        Assert.NotNull(Assert.Throws<DocumentConversionException>(() => DocxForm.Inspect(rubbish)).InnerException);
        Assert.NotNull(Assert.Throws<DocumentConversionException>(() => DocxForm.Validate(rubbish, values)).InnerException);
        Assert.NotNull(Assert.Throws<DocumentConversionException>(() => DocxForm.Fill(rubbish, values)).InnerException);
    }

    [Fact]
    public async Task TheStreamOverloadsMatchTheirByteArrayTwins_AndLeaveStreamsOpen()
    {
        using var source = new MemoryStream(DocxFormFixtures.Form());
        DocxFormReport report = await DocxForm.InspectAsync(source);
        Assert.Equal(4, report.Fields.Count);
        source.Position = 0;
        Assert.True(source.ReadByte() >= 0, "the stream the caller owns must not be closed");

        using var fillSource = new MemoryStream(DocxFormFixtures.Form());
        using var destination = new MemoryStream();
        await DocxForm.FillAsync(fillSource, destination,
            new Dictionary<string, DocxFormValue> { ["Notes"] = DocxFormValue.FromText("streamed") });

        DocxFormReport after = DocxForm.Inspect(destination.ToArray());
        Assert.Equal("streamed", Assert.Single(after.Fields, f => f.Key == "Notes").Value.Text);
    }
}
