using OfficeIMO.Word;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// <see cref="DocxFormField.Alias"/> — the display name Word shows, reported separately from the
/// key used for matching (A120).
///
/// <b>The alias comes from OfficeIMO's own control collection, not a second walk of the OOXML.</b>
/// Measured 2026-09-03: <c>WordStructuredDocumentTag</c> exposes <c>Tag</c> and <c>Alias</c>
/// together — so identity stays OfficeIMO's decision and no reader keyed on the same names was
/// added. The lock is a different story and is deliberately still absent; see the row.
/// </summary>
public class DocxFormAliasTests
{
    [Fact]
    public void ReportsTheAliasSeparatelyFromTheKey()
    {
        // The fixture authors Tag "FullName" against Alias "Full name". Asserting only that the
        // alias is non-empty would pass against a reader that returned the tag for both, which is
        // the obvious wrong implementation - so this asserts the literal AND that they differ.
        var field = Assert.Single(
            DocxForm.Inspect(DocxFormFixtures.Form()).Fields,
            f => f.Key == "FullName");

        Assert.Equal("Full name", field.Alias);
        Assert.NotEqual(field.Key, field.Alias);
    }

    [Fact]
    public void EveryFieldCarriesItsOwnAlias()
    {
        // Not one alias applied to all of them, which a lookup returning the first match would do.
        var fields = DocxForm.Inspect(DocxFormFixtures.Form()).Fields;

        Assert.Equal("Full name", Assert.Single(fields, f => f.Key == "FullName").Alias);
        Assert.Equal("Plan", Assert.Single(fields, f => f.Key == "Plan").Alias);
        Assert.Equal("Start date", Assert.Single(fields, f => f.Key == "Start").Alias);
        Assert.Equal("Signed", Assert.Single(fields, f => f.Key == "Signed").Alias);
    }

    /// <summary>
    /// Word does not require a tag to be unique. Where one name reaches two controls with
    /// DIFFERENT aliases, the honest answer is no alias rather than one of the two picked
    /// arbitrarily — being silently wrong half the time is worse than saying nothing, and
    /// <c>Validate</c> already reports the ambiguity as <c>DuplicateKey</c>.
    /// </summary>
    /// <remarks>
    /// <b>This needed its own fixture.</b> <c>DocxFormFixtures.DuplicateTags()</c> gives both
    /// controls the SAME alias, so reporting it is correct there and the ambiguity path is never
    /// reached — the first version of this test asserted against that fixture and was wrong about
    /// what it exercised rather than about the code.
    /// </remarks>
    [Fact]
    public void SaysNothingWhenTheKeyReachesTwoControlsWithDifferentAliases()
    {
        // AddStructuredDocumentTag(text, alias, tag) - one tag, two different aliases.
        var docx = DocxFormFixtures.Authored(document =>
        {
            document.AddParagraph().AddStructuredDocumentTag("one", "First label", "Same");
            document.AddParagraph().AddStructuredDocumentTag("two", "Second label", "Same");
        });

        var fields = DocxForm.Inspect(docx).Fields;

        Assert.Contains(fields, f => f.Key == "Same");
        Assert.All(fields.Where(f => f.Key == "Same"), f => Assert.Equal(string.Empty, f.Alias));
    }

    /// <summary>
    /// And the other side of it: two controls sharing a name AND an alias is not ambiguous about
    /// the alias, so that one is still reported. A guard that blanked the alias on any duplicate
    /// key would fail here.
    /// </summary>
    [Fact]
    public void StillReportsTheAliasWhenTwoControlsShareBothNames()
    {
        var fields = DocxForm.Inspect(DocxFormFixtures.DuplicateTags()).Fields;

        Assert.All(fields.Where(f => f.Key == "Same"), f => Assert.Equal("Same", f.Alias));
    }

    /// <summary>
    /// The positive control for the test above. A lookup that returned empty for EVERY key would
    /// satisfy it, and nothing else in this file would notice on the duplicate fixture alone.
    /// </summary>
    [Fact]
    public void PositiveControl_AnUnambiguousKeyStillGetsItsAlias()
    {
        var field = Assert.Single(
            DocxForm.Inspect(DocxFormFixtures.Form()).Fields,
            f => f.Key == "Signed");

        Assert.Equal("Signed", field.Alias);
    }

    [Fact]
    public async Task TheAsyncOverloadReportsTheSameAliases()
    {
        using var source = new MemoryStream(DocxFormFixtures.Form(), writable: false);

        var report = await DocxForm.InspectAsync(source);

        Assert.Equal("Full name", Assert.Single(report.Fields, f => f.Key == "FullName").Alias);
    }

    [Fact]
    public void AControlWithNoAliasReportsAnEmptyString()
    {
        // Never null: a caller rendering a label should not have to null-check a name.
        var fields = DocxForm.Inspect(DocxFormFixtures.Form()).Fields;

        Assert.All(fields, f => Assert.NotNull(f.Alias));
    }
}
