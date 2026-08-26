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
    public void DocxFormIssueKind_SurfacesEveryKindTheLibraryBeneathReports()
    {
        // This asserted a set of THREE, on a measurement that was an artefact of its own fixtures -
        // hand-built SdtBlock markup is not a typed control, so nothing could report InvalidChoice
        // because there was no drop-down. Against a real form, three of the six "unreachable" kinds
        // fire immediately. See DocxFormFixtures for the rule that earns.
        Assert.Equal(
            ["Other", "MissingValue", "UnusedValue", "DuplicateKey", "UnmappedControl",
             "InvalidBoolean", "InvalidDate", "InvalidChoice", "InvalidImage",
             "InvalidRepeatingSection"],
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
        byte[] docx = DocxFormFixtures.DuplicateTags();

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
                ["Signed"] = DocxFormValue.FromChecked(true),
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
                ["Signed"] = DocxFormValue.FromChecked(true),
            });

        DocxFormReport report = DocxForm.Inspect(filled);
        Assert.Equal("Someone Else", Assert.Single(report.Fields, f => f.Key == "FullName").Value.Text);

        // The round trip across TYPES, not just text: a check box read back as a bool.
        DocxFormValue signed = Assert.Single(report.Fields, f => f.Key == "Signed").Value;
        Assert.Equal(DocxFormValueKind.Checked, signed.Kind);
        Assert.True(signed.Checked);
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
        Assert.Equal(new DateTime(2026, 1, 15), Assert.Single(report.Fields, f => f.Key == "Start").Value.Date);
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

        // The date picker and check box render as their own glyphs rather than as text, which is
        // the document's business and not this API's - what matters is that the block boundaries
        // survive, so the paragraphs are still four separate blocks.
        Assert.Equal(4, DocxEditor.ExtractText(filled).Split('\n').Length);
        Assert.StartsWith("A\n", DocxEditor.ExtractText(filled), StringComparison.Ordinal);
    }

    [Fact]
    public void Fill_RefusesANullValueInTheDictionary()
    {
        var ex = Assert.Throws<ArgumentException>(() => DocxForm.Fill(DocxFormFixtures.Form(),
            new Dictionary<string, DocxFormValue> { ["FullName"] = null! }));

        Assert.Equal("values", ex.ParamName);
        Assert.Contains("FullName", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Fill_WritesADateAsADate_AndNotAsPreFormattedText()
    {
        // THE MOST DANGEROUS THING THIS SUITE PINS, and an earlier version of this test declared it
        // untestable. It claimed passing a DateTime and passing that DateTime's ToString() were
        // equivalent, so no test could tell them apart. Measured on a real date picker, en-CA:
        //
        //     DateTime            -> 2027-03-09
        //     DateTime.ToString() -> 2027-09-03      <- March 9 became September 3
        //
        // ToString() uses the current culture; the library beneath re-parses with InvariantCulture.
        // The result is a valid document with a silently transposed date and nothing raised. Not an
        // equivalent mutant - a data-corruption bug that a comment was telling people not to test.
        byte[] filled = DocxForm.Fill(DocxFormFixtures.Form(),
            new Dictionary<string, DocxFormValue>
            {
                ["Start"] = DocxFormValue.FromDate(new DateTime(2027, 3, 9)),
            });

        DocxFormValue start = Assert.Single(
            DocxForm.Inspect(filled).Fields, f => f.Key == "Start").Value;

        // Assert the DAY and MONTH, not merely the year: only that distinguishes 09-03 from 03-09.
        Assert.Equal(DocxFormValueKind.Date, start.Kind);
        Assert.Equal(new DateTime(2027, 3, 9), start.Date);
    }

    [Fact]
    public void Validate_ReportsAValueThatDoesNotFitATypedControl()
    {
        // The three kinds an earlier draft called unreachable. They need a form authored the way
        // Word authors one - which is the whole reason DocxFormFixtures no longer hand-builds markup.
        DocxFormValidation result = DocxForm.Validate(DocxFormFixtures.Form(),
            new Dictionary<string, DocxFormValue>
            {
                ["FullName"] = DocxFormValue.FromText("fine"),
                ["Plan"] = DocxFormValue.FromChoice("NotAnOption"),
                ["Start"] = DocxFormValue.FromText("not a date"),
                ["Signed"] = DocxFormValue.FromText("not a bool"),
            });

        Assert.False(result.IsValid);
        Assert.Contains(result.Issues, i => i.Kind == DocxFormIssueKind.InvalidChoice && i.Key == "Plan");
        Assert.Contains(result.Issues, i => i.Kind == DocxFormIssueKind.InvalidDate && i.Key == "Start");
        Assert.Contains(result.Issues, i => i.Kind == DocxFormIssueKind.InvalidBoolean && i.Key == "Signed");
    }

    [Fact]
    public void Validate_CarriesTheUnderlyingMessageRatherThanBlankingIt()
    {
        // Nothing asserted Message anywhere, so returning string.Empty for every issue passed the
        // whole suite - and Message is the only thing carrying detail for a kind this API collapses.
        DocxFormValidation result = DocxForm.Validate(DocxFormFixtures.Form(),
            new Dictionary<string, DocxFormValue> { ["FullName"] = DocxFormValue.FromText("x") });

        DocxFormIssue issue = result.Issues.First(i => i.Kind == DocxFormIssueKind.MissingValue);
        Assert.False(string.IsNullOrWhiteSpace(issue.Message));
        Assert.Contains(issue.Key, issue.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void APictureValueReachesTheLibraryAsBYTES_NeverAsAPath()
    {
        // THE BEHAVIOURAL HALF of the bytes-only guarantee, and it was missing. The surface test
        // pins that no FACTORY takes a path; nothing pinned that the conversion uses FromBytes.
        // Rewriting that one line to FromFile(FileName) passed every other test in this suite while
        // making the shipped package read a caller-supplied string off disk.
        object? upstream = DocxFormValue.FromPicture([1, 2, 3], "logo.png").ToUpstream();

        var picture = Assert.IsType<OfficeIMO.Word.WordContentControlPictureValue>(upstream);
        Assert.Null(picture.FilePath);
        Assert.Null(picture.ExternalUri);
        Assert.False(picture.IsExternal);
        Assert.Equal<byte[]>([1, 2, 3], picture.Bytes!);
    }

    [Fact]
    public void Inspect_DoesNotClaimToReturnControlsItCannotName()
    {
        // Fields is documented as what the document EXPOSES under a key mode, not as every control,
        // because upstream drops some. Both cases are pinned here so the doc comment stays true.
        Assert.Single(DocxForm.Inspect(DocxFormFixtures.DuplicateTags()).Fields);

        // The dangerous one: a tag-only template read by alias yields nothing at all, which looks
        // exactly like a document with no form in it.
        Assert.DoesNotContain(DocxForm.Inspect(DocxFormFixtures.Form(), DocxFormKey.Alias).Fields,
            f => f.Key == "FullName");
        Assert.NotEmpty(DocxForm.Inspect(DocxFormFixtures.Form(), DocxFormKey.Tag).Fields);
    }

    [Fact]
    public void AHandBuiltSdtBlockIsNotATypedControl_WhichIsWhyFixturesAreAuthored()
    {
        // Pins the mistake that invalidated this feature's first round of measurements, so it cannot
        // be made again silently. A hand-built SdtBlock IS a content control - it has a key and it
        // extracts - but it is not a TYPED one, so no value constraint applies to it.
        byte[] handBuilt = DocxFormFixtures.UntypedBlockControl("Plan", "Pro");

        Assert.Single(DocxForm.Inspect(handBuilt).Fields);
        Assert.True(DocxForm.Validate(handBuilt,
            new Dictionary<string, DocxFormValue> { ["Plan"] = DocxFormValue.FromChoice("NotAnOption") })
            .IsValid);
    }

    [Fact]
    public void EveryKeyIsCarriedThrough_EvenOneWhoseValueMapsToNull()
    {
        // An unset date picker and an unselected drop-down both read back as null, so the advertised
        // Inspect-then-Fill round trip produced values that mapped to null. Dropping those made
        // Validate report MissingValue for a key the caller HAD supplied.
        byte[] empty = DocxFormFixtures.Authored(d =>
            d.AddParagraph().AddDatePicker(null, "Start date", "Start"));

        DocxFormValue readBack = Assert.Single(DocxForm.Inspect(empty).Fields).Value;
        DocxFormValidation result = DocxForm.Validate(empty,
            new Dictionary<string, DocxFormValue> { ["Start"] = readBack });

        Assert.Contains("Start", result.SuppliedKeys);
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

        // Against its byte[] twin, on the same input - which is what the name claims.
        byte[] form = DocxFormFixtures.Form();
        var values = new Dictionary<string, DocxFormValue>
        {
            ["FullName"] = DocxFormValue.FromText("streamed"),
        };

        using var fillSource = new MemoryStream(form);
        using var destination = new MemoryStream();
        await DocxForm.FillAsync(fillSource, destination, values);

        Assert.Equal(
            DocxForm.Inspect(DocxForm.Fill(form, values)).Fields.Select(f => f.Key + "=" + f.Value.Text),
            DocxForm.Inspect(destination.ToArray()).Fields.Select(f => f.Key + "=" + f.Value.Text));

        using var validateSource = new MemoryStream(form);
        DocxFormValidation streamed = await DocxForm.ValidateAsync(validateSource, values);
        Assert.Equal(DocxForm.Validate(form, values).IsValid, streamed.IsValid);
    }
}
