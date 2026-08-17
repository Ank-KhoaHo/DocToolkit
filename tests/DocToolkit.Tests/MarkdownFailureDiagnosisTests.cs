namespace DocToolkit.Tests;

/// <summary>
/// The two things Markdown conversion rejects, and the messages that now name them.
///
/// <b>Measured against the CommonMark 0.31.2 conformance suite - 652 examples - plus four real
/// project READMEs.</b> Markdown is the best-performing capability in this package, and both of its
/// rejections are spec-valid input arriving as an unhandled exception from inside a dependency: a
/// bare <see cref="NullReferenceException"/> and an <see cref="ArgumentOutOfRangeException"/>,
/// neither naming a construct.
///
/// <b>Neither is repaired, deliberately.</b> Both repairs would change what the document says - a
/// line feed written as a character reference is a character the author wrote, and an ordered list
/// starting at zero renumbers if it is made to start at one. So the diagnosis improves and the
/// behaviour does not.
///
/// <b>One correction to the backlog row while measuring: the ordered-list rejection is PDF-only.</b>
/// <c>0. ok</c> converts to DOCX perfectly well. The row said both converters.
/// </summary>
public class MarkdownFailureDiagnosisTests
{
    // ---- the line-feed character reference ----------------------------------------------------------

    [Theory]
    [InlineData("a&#10;b")]
    [InlineData("foo&#10;&#10;bar")]
    [InlineData("a&#x0A;b")]
    [InlineData("a&#xA;b")]
    public void ALineFeedReferenceIsNamed_OnTheDocxPath(string markdown)
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => MarkdownToDocxConverter.Convert(markdown));

        Assert.Contains("&#10;", ex.Message, StringComparison.Ordinal);
        Assert.Contains("valid CommonMark", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ALineFeedReferenceIsNamed_OnThePdfPathToo()
    {
        // The PDF path pivots through DOCX, so it fails in the same place - but a caller who wrote
        // Markdown and asked for a PDF should not have to know that.
        var ex = Assert.Throws<DocumentConversionException>(
            () => MarkdownToPdfConverter.Convert("a&#10;b"));

        Assert.Contains("&#10;", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheMessageSaysARealNewlineWorks_BecauseItDoes()
    {
        // A message that named the construct and stopped would leave the reader stuck: the whole
        // remedy is one character. Measured - a real newline converts.
        var ex = Assert.Throws<DocumentConversionException>(
            () => MarkdownToDocxConverter.Convert("a&#10;b"));

        Assert.Contains("real newline converts", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.NotEmpty(MarkdownToDocxConverter.Convert("a\nb"));
    }

    [Theory]
    [InlineData("a&#9;b")]    // tab
    [InlineData("a&#13;b")]   // carriage return
    [InlineData("a&#32;b")]   // space
    [InlineData("a&#0;b")]    // NUL
    [InlineData("a&#65;b")]   // a letter
    [InlineData("a&amp;b")]   // a named entity
    public void EveryOtherCharacterReferenceStillConverts(string markdown)
    {
        // The boundary is narrow and measured. A diagnosis that fired on character references in
        // general would be describing a problem that does not exist.
        Assert.NotEmpty(MarkdownToDocxConverter.Convert(markdown));
    }

    // ---- the ordered list starting below 1 -----------------------------------------------------------

    [Theory]
    [InlineData("0. ok")]
    [InlineData("0) ok")]
    public void AListStartingBelowOneIsNamed_OnThePdfPath(string markdown)
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => MarkdownToPdfConverter.Convert(markdown));

        Assert.Contains("ordered list starting below 1", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("0. ok")]
    [InlineData("0) ok")]
    public void TheSameListConvertsToDocx(string markdown)
    {
        // The correction to the backlog row, pinned: this is a PDF-stage limit, not a rejection of
        // the document. The message says so, and would be a lie if this ever stopped being true.
        Assert.NotEmpty(MarkdownToDocxConverter.Convert(markdown));
    }

    [Fact]
    public void TheMessageSaysTheDocxPathWorks()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => MarkdownToPdfConverter.Convert("0. ok"));

        Assert.Contains("DOCX works", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("1. ok")]
    [InlineData("2. ok")]
    [InlineData("- ok")]
    public void OrdinaryListsStillRender(string markdown)
    {
        Assert.NotEmpty(MarkdownToPdfConverter.Convert(markdown));
    }

    // ---- what must NOT be claimed --------------------------------------------------------------------

    [Fact]
    public void AnUnrecognisedFailureKeepsTheGenericWrapper()
    {
        // A vertical tab reference is invalid in XML and fails in the writer - a real failure of a
        // completely different kind. A change that put a Markdown-construct message on every
        // failure would fail here.
        var ex = Assert.Throws<DocumentConversionException>(
            () => MarkdownToDocxConverter.Convert("a&#11;b"));

        Assert.Contains("See the inner exception", ex.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("&#10;", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheFrameAloneIsNotEnough_TheInputMustContainTheReference()
    {
        // AddRun is where EVERY inline run is built, so the frame identifies the site and not the
        // cause. Without the input check, any null reference raised anywhere in that method would be
        // reported as a line-feed reference the document does not contain.
        var staged = new Staged("   at OfficeIMO.Word.Markdown.MarkdownToWordConverter.AddRun(String s)");

        Assert.Null(MarkdownFailureDiagnosis.Describe(staged, "no references here"));
        Assert.NotNull(MarkdownFailureDiagnosis.Describe(staged, "has a &#10; in it"));
    }

    private sealed class Staged(string stack) : NullReferenceException
    {
        public override string StackTrace { get; } = stack;
    }

    private sealed class StagedRange(string stack) : ArgumentOutOfRangeException
    {
        public override string StackTrace { get; } = stack;
    }

    [Fact]
    public void TheInputAloneIsNotEnough_TheFrameMustMatchToo()
    {
        // The mirror of the test above, and it needed writing: nothing else here has a null
        // reference from somewhere OTHER than AddRun while the input happens to contain a
        // character reference, so dropping the frame check survived every assertion in the file.
        // Mutation testing found that.
        var elsewhere = new Staged("   at Some.Other.Component.DoesSomething(String s)");

        Assert.Null(MarkdownFailureDiagnosis.Describe(elsewhere, "this text has a &#10; in it"));
    }

    [Fact]
    public void AnOutOfRangeFromElsewhereIsNotAListStart()
    {
        // Same gap on the other rule: every ArgumentOutOfRangeException reaching these tests came
        // from the numbered-list block, so matching on the type alone was indistinguishable from
        // matching on type AND frame. A range error from anywhere else must not be reported as a
        // list that starts at zero.
        var elsewhere = new StagedRange("   at Some.Other.Component.Index(Int32 i)");

        Assert.Null(MarkdownFailureDiagnosis.Describe(elsewhere, "1. ok"));
    }

    [Fact]
    public void ADifferentExceptionTypeGetsNoDiagnosis()
    {
        Assert.Null(MarkdownFailureDiagnosis.Describe(
            new InvalidOperationException("anything"), "a&#10;b"));
    }

    [Fact]
    public void OrdinaryMarkdownStillConverts()
    {
        Assert.NotEmpty(MarkdownToDocxConverter.Convert("# Title\n\nBody with **bold**.\n\n- one\n- two"));
        Assert.NotEmpty(MarkdownToPdfConverter.Convert("# Title\n\nBody with **bold**."));
    }
}
