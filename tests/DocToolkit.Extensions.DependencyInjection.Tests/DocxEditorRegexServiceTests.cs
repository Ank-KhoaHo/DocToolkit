using System.Text.RegularExpressions;
using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

/// <summary>
/// The regex <c>ReplaceText</c> overloads on <see cref="IDocxEditor"/> (A116-DI).
///
/// <b>The interface-completeness test could not have caught these missing.</b> It compares method
/// NAMES, and <c>ReplaceText</c> was already on the interface — verified by bumping the floor to
/// 0.54.0 and watching it stay green with both overloads absent. So the mechanism that found
/// A114-DI is structurally blind here, and only the filed row catches it.
///
/// Each test asserts a literal before it asserts delegation: "the wrapper matches the thing it
/// wraps" holds identically when both do nothing.
/// </summary>
public class DocxEditorRegexServiceTests
{
    private static DocxEditorService Sut() =>
        new(new TestOptionsMonitor<DocToolkitOptions>(new DocToolkitOptions()));

    private static Regex Pattern(string text) => new(text, RegexOptions.None, TimeSpan.FromSeconds(2));

    private static byte[] Doc(string text) => DocxEditor.Create([DocxBlock.Paragraph(text)]);

    [Fact]
    public void ReplaceTextMatchesAPatternAndExpandsCaptureGroups()
    {
        var sut = Sut();

        var edited = sut.ReplaceText(Doc("due 2026-04-17"), Pattern(@"(\d{4})-(\d{2})-(\d{2})"), "$3/$2/$1");

        Assert.Equal("due 17/04/2026", DocxEditor.ExtractText(edited));
    }

    [Fact]
    public void ReplaceTextRefusesAPatternWithNoMatchTimeout()
    {
        var sut = Sut();

        var ex = Assert.Throws<ArgumentException>(
            () => sut.ReplaceText(Doc("value 42"), new Regex(@"\d+"), "N"));

        Assert.Equal("pattern", ex.ParamName);
    }

    /// <summary>
    /// The positive control for the refusal above. Without it, a wrapper that threw on every
    /// pattern would satisfy that test and nothing here would notice.
    /// </summary>
    [Fact]
    public void PositiveControl_ABoundedPatternIsAccepted()
    {
        var sut = Sut();

        var edited = sut.ReplaceText(Doc("value 42"), Pattern(@"\d+"), "N");

        Assert.Equal("value N", DocxEditor.ExtractText(edited));
    }

    [Fact]
    public async Task ReplaceTextAsyncMatchesTheByteArrayForm_AndLeavesBothStreamsOpen()
    {
        var sut = Sut();
        using var source = new MemoryStream(Doc("Invoice 2026-04-17"), writable: false);
        using var destination = new MemoryStream();

        await sut.ReplaceTextAsync(source, Pattern(@"\d{4}-\d{2}-\d{2}"), "[date]", destination);

        Assert.Equal("Invoice [date]", DocxEditor.ExtractText(destination.ToArray()));
        Assert.True(source.CanRead, "ReplaceTextAsync disposed a source stream it does not own.");
        Assert.True(destination.CanWrite, "ReplaceTextAsync closed a destination it does not own.");
    }

    [Fact]
    public async Task ReplaceTextAsyncRefusesAnUnboundedPatternToo()
    {
        var sut = Sut();
        using var source = new MemoryStream(Doc("value 42"));
        using var destination = new MemoryStream();

        var ex = await Assert.ThrowsAsync<ArgumentException>(
            () => sut.ReplaceTextAsync(source, new Regex(@"\d+"), "N", destination));

        Assert.Equal("pattern", ex.ParamName);
    }

    [Fact]
    public async Task ReplaceTextAsyncRefusesAnAlreadyCancelledToken()
    {
        var sut = Sut();
        using var cts = new CancellationTokenSource();
        await cts.CancelAsync();
        using var source = new MemoryStream(Doc("value 42"));
        using var destination = new MemoryStream();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => sut.ReplaceTextAsync(source, Pattern(@"\d+"), "N", destination, cts.Token));
    }
}
