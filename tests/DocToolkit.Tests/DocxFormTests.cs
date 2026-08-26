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
}
