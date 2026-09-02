using System.Text.RegularExpressions;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// Pattern-based find-and-replace on a real .docx (A116).
///
/// <c>RunTextSplicerTests</c> already proves the splice itself — that a match spanning runs is
/// replaced exactly like one inside a single run, and that runs a match does not overlap are never
/// written to. This file proves the other half: that the pattern reaches every part of the package
/// the literal overload reaches, and that the guard on an unbounded pattern is real.
///
/// Assertions are on exact strings rather than substrings. A replacement that ran twice, or that
/// relocated text between runs, still contains every token you would search for — only the whole
/// string discriminates.
/// </summary>
public class DocxEditorRegexReplaceTests
{
    private static Regex Pattern(string text) =>
        new(text, RegexOptions.None, TimeSpan.FromSeconds(2));

    private static byte[] Doc(params string[] paragraphs) =>
        DocxEditor.Create([.. paragraphs.Select(DocxBlock.Paragraph)]);

    [Fact]
    public void ReplacesEveryMatchInTheBody()
    {
        var docx = Doc("Invoice 2026-04-17 and 2026-05-01");

        var edited = DocxEditor.ReplaceText(docx, Pattern(@"\d{4}-\d{2}-\d{2}"), "[date]");

        Assert.Equal("Invoice [date] and [date]", DocxEditor.ExtractText(edited));
    }

    [Fact]
    public void ExpandsCaptureGroupsInTheReplacement()
    {
        var docx = Doc("due 2026-04-17");

        var edited = DocxEditor.ReplaceText(docx, Pattern(@"(\d{4})-(\d{2})-(\d{2})"), "$3/$2/$1");

        Assert.Equal("due 17/04/2026", DocxEditor.ExtractText(edited));
    }

    [Fact]
    public void ADoubledDollarIsALiteralDollar()
    {
        // The replacement is a template, so this is the escape a caller needs and the one thing
        // they cannot discover from the signature.
        var docx = Doc("costs 42 units");

        var edited = DocxEditor.ReplaceText(docx, Pattern(@"\d+"), "$$100");

        Assert.Equal("costs $100 units", DocxEditor.ExtractText(edited));
    }

    [Fact]
    public void APatternThatMatchesNothingReturnsTheDocumentUnchanged()
    {
        var docx = Doc("Acme Corporation");

        var edited = DocxEditor.ReplaceText(docx, Pattern(@"\d+"), "N");

        Assert.Equal("Acme Corporation", DocxEditor.ExtractText(edited));
    }

    [Fact]
    public void TheReplacementIsNotItselfRescanned()
    {
        // A loop that re-ran the pattern over its own output would keep matching here. The literal
        // overload makes the same guarantee, and this is the pattern-shaped way to break it.
        var docx = Doc("a1");

        var edited = DocxEditor.ReplaceText(docx, Pattern(@"\d"), "2");

        Assert.Equal("a2", DocxEditor.ExtractText(edited));
    }

    [Fact]
    public void ReachesHeadersAndFootersLikeTheLiteralOverload()
    {
        var page = PageSetup.A4
            .WithHeader(DocxHeader.Text("Ref 2026-04-17"))
            .WithFooter(DocxHeader.Text("Ref 2026-05-01"));
        var docx = DocxEditor.Create([DocxBlock.Paragraph("Body 2026-06-02")], page);

        var edited = DocxEditor.ReplaceText(docx, Pattern(@"\d{4}-\d{2}-\d{2}"), "[d]");

        var all = DocxEditor.ExtractText(edited, includeHeadersAndFooters: true);
        Assert.Contains("Ref [d]", all, StringComparison.Ordinal);
        Assert.Contains("Body [d]", all, StringComparison.Ordinal);
        Assert.DoesNotContain("2026-", all, StringComparison.Ordinal);
    }

    // ---------- the unbounded-pattern guard ----------

    [Fact]
    public void RefusesAPatternWithNoMatchTimeout()
    {
        var docx = Doc("anything");

        var ex = Assert.Throws<ArgumentException>(
            () => DocxEditor.ReplaceText(docx, new Regex(@"\d+"), "x"));

        Assert.Equal("pattern", ex.ParamName);
        Assert.Contains("no match timeout", ex.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// The positive control for the test above. Without it, the refusal could be firing on every
    /// pattern — including a correctly constructed one — and both tests would still look right.
    /// </summary>
    [Fact]
    public void PositiveControl_APatternWithATimeoutIsAccepted()
    {
        var docx = Doc("value 42");

        var edited = DocxEditor.ReplaceText(docx, Pattern(@"\d+"), "N");

        Assert.Equal("value N", DocxEditor.ExtractText(edited));
    }

    [Fact]
    public async Task TheAsyncOverloadRefusesAnUnboundedPatternToo()
    {
        using var source = new MemoryStream(Doc("anything"));
        using var destination = new MemoryStream();

        var ex = await Assert.ThrowsAsync<ArgumentException>(
            () => DocxEditor.ReplaceTextAsync(source, new Regex(@"\d+"), "x", destination));

        Assert.Equal("pattern", ex.ParamName);
    }

    // ---------- argument guards and the Stream overload ----------

    [Fact]
    public void RejectsNullPatternAndNullReplacement()
    {
        var docx = Doc("x");

        Assert.Throws<ArgumentNullException>(() => DocxEditor.ReplaceText(docx, null!, "y"));
        Assert.Throws<ArgumentNullException>(() => DocxEditor.ReplaceText(docx, Pattern("x"), null!));
    }

    [Fact]
    public void RejectsEmptyContent()
    {
        var ex = Assert.Throws<ArgumentException>(
            () => DocxEditor.ReplaceText([], Pattern("x"), "y"));

        Assert.Equal("docx", ex.ParamName);
    }

    [Fact]
    public async Task TheStreamOverloadAgreesWithTheByteArrayForm_AndLeavesBothStreamsOpen()
    {
        var docx = Doc("Invoice 2026-04-17");
        using var source = new MemoryStream(docx, writable: false);
        using var destination = new MemoryStream();

        await DocxEditor.ReplaceTextAsync(source, Pattern(@"\d{4}-\d{2}-\d{2}"), "[date]", destination);

        Assert.Equal("Invoice [date]", DocxEditor.ExtractText(destination.ToArray()));
        Assert.True(source.CanRead, "the caller's source was closed");
        Assert.True(destination.CanWrite, "the caller's destination was closed");
    }

    [Fact]
    public async Task TheStreamOverloadRefusesAnAlreadyCancelledToken()
    {
        using var source = new MemoryStream(Doc("x"));
        using var destination = new MemoryStream();
        using var cts = new CancellationTokenSource();
        await cts.CancelAsync();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => DocxEditor.ReplaceTextAsync(source, Pattern("x"), "y", destination, cts.Token));
    }
}
