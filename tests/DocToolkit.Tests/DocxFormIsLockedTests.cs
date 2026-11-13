using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// <see cref="DocxFormField.IsLocked"/> and <see cref="DocxFormIssueKind.LockedControl"/> (A120).
///
/// <b>Measured before it was built</b> (<c>artifacts/a120/probe-locks-results-2026-11-13.txt</c>):
/// OfficeIMO 3.5.3's <c>ValidateContentControlValues</c> reports <c>LockedControl</c> for a supplied
/// key whose control is <c>SdtContentLocked</c> or <c>ContentLocked</c>, not <c>SdtLocked</c>, and
/// resolves the key with the same <c>DocxFormKey</c> rule as <c>ExtractContentControlValues</c>. So
/// the lock is read through OfficeIMO's own key resolution — no second reader of the OOXML.
///
/// <b>The invariant these tests hold:</b> a field reports <c>IsLocked</c> exactly when
/// <see cref="DocxForm.Fill"/> would refuse to write it. A UI that greys out locked fields must never
/// grey out one the fill accepts, or leave enabled one the fill refuses.
/// </summary>
public class DocxFormIsLockedTests
{
    private static byte[] LockedForm(string how) =>
        DocxFormLockTests.Locked(DocxFormFixtures.Form(), DocxFormLockTests.ByName(how));

    [Theory]
    [InlineData("SdtContentLocked", DocxFormKey.TagThenAlias, "FullName")]
    [InlineData("SdtContentLocked", DocxFormKey.Tag, "FullName")]
    [InlineData("SdtContentLocked", DocxFormKey.Alias, "Full name")]
    [InlineData("SdtContentLocked", DocxFormKey.AliasThenTag, "Full name")]
    [InlineData("ContentLocked", DocxFormKey.TagThenAlias, "FullName")]
    [InlineData("ContentLocked", DocxFormKey.Alias, "Full name")]
    public void AContentLockedFieldIsLocked_AndNoOtherFieldIs(string how, DocxFormKey key, string lockedKey)
    {
        var fields = DocxForm.Inspect(LockedForm(how), key).Fields;

        Assert.True(Assert.Single(fields, f => f.Key == lockedKey).IsLocked);
        // The fixture has four controls; the other three carry no lock. A reader that reported every
        // field as locked once any lock was present would pass the line above and fail here.
        Assert.Equal(3, fields.Count(f => !f.IsLocked));
    }

    [Fact]
    public void AnSdtLockedFieldIsNotLocked()
    {
        var fields = DocxForm.Inspect(LockedForm("SdtLocked")).Fields;

        Assert.Equal(4, fields.Count);
        Assert.All(fields, f => Assert.False(f.IsLocked));
    }

    [Fact]
    public void NoFieldIsLockedInAFormWithNoLocks()
    {
        var fields = DocxForm.Inspect(DocxFormFixtures.Form()).Fields;

        Assert.Equal(4, fields.Count);
        Assert.All(fields, f => Assert.False(f.IsLocked));
    }

    /// <summary>The invariant, checked directly for every lock kind.</summary>
    [Theory]
    [InlineData("SdtContentLocked")]
    [InlineData("ContentLocked")]
    [InlineData("SdtLocked")]
    public void IsLockedIsTrueExactlyWhenFillRefuses(string how)
    {
        var docx = LockedForm(how);
        bool reported = Assert.Single(DocxForm.Inspect(docx).Fields, f => f.Key == "FullName").IsLocked;

        bool refused;
        try
        {
            DocxForm.Fill(docx, new Dictionary<string, DocxFormValue> { ["FullName"] = DocxFormValue.FromText("X") });
            refused = false;
        }
        catch (InvalidOperationException)
        {
            refused = true;
        }

        Assert.Equal(refused, reported);
    }

    /// <summary>
    /// Two controls share one tag and both are locked. Extraction reports the key once; that one
    /// field must still say it is locked.
    /// </summary>
    [Fact]
    public void ADuplicatedLockedKeyIsLocked()
    {
        var docx = DocxFormLockTests.Locked(DocxFormFixtures.DuplicateTags(), LockingValues.SdtContentLocked, "Same");

        Assert.True(Assert.Single(DocxForm.Inspect(docx).Fields).IsLocked);
    }

    /// <summary>
    /// <c>Inspect</c> hands every extracted value back to OfficeIMO's validation to learn the locks,
    /// so a value kind that does not survive that round trip would make <c>Inspect</c> throw. A
    /// picture is the least ordinary value a control can hold.
    /// </summary>
    [Fact]
    public void AFormWithAPictureControlStillInspects()
    {
        var field = Assert.Single(DocxForm.Inspect(DocxFormFixtures.WithPictureControl()).Fields);

        Assert.False(field.IsLocked);
    }

    [Fact]
    public void ValidateReportsALockedControlByItsOwnKind()
    {
        var validation = DocxForm.Validate(
            LockedForm("SdtContentLocked"),
            new Dictionary<string, DocxFormValue> { ["FullName"] = DocxFormValue.FromText("X") });

        var issue = Assert.Single(validation.Issues, i => i.Key == "FullName");
        Assert.Equal(DocxFormIssueKind.LockedControl, issue.Kind);
        Assert.DoesNotContain(validation.Issues, i => i.Kind == DocxFormIssueKind.Other);
    }

    /// <summary>
    /// The new kind's number is part of the public contract: a consumer that stored it, or switched
    /// on it from another assembly, depends on it staying 10.
    /// </summary>
    [Fact]
    public void LockedControlIsTen() => Assert.Equal(10, (int)DocxFormIssueKind.LockedControl);
}
