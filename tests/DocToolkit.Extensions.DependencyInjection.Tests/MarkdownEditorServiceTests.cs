using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class MarkdownEditorServiceTests
{
    private const string Sample = """
        ---
        title: Report
        ---
        # Overview
        Intro text.

        | Region | Total |
        |---|---|
        | North | 1200 |

        ## Details
        More text.
        """;

    [Fact]
    public void ReadFrontMatter_MatchesTheStaticMethod()
    {
        var sut = new MarkdownEditorService();

        var fromWrapper = sut.ReadFrontMatter(Sample);

        Assert.Equal(MarkdownEditor.ReadFrontMatter(Sample), fromWrapper);
        Assert.Equal("Report", fromWrapper["title"]);
    }

    [Fact]
    public void FindHeading_MatchesTheStaticMethodAndReturnsTheRealHeading()
    {
        var sut = new MarkdownEditorService();

        var fromWrapper = sut.FindHeading(Sample, "Details");

        Assert.Equal(MarkdownEditor.FindHeading(Sample, "Details")?.Text, fromWrapper?.Text);
        Assert.NotNull(fromWrapper);
        Assert.Equal(2, fromWrapper!.Level);
    }

    [Fact]
    public void TableCount_MatchesTheStaticMethodAndReturnsTheRealCount()
    {
        var sut = new MarkdownEditorService();

        var count = sut.TableCount(Sample);

        Assert.Equal(MarkdownEditor.TableCount(Sample), count);
        Assert.Equal(1, count);
    }

    [Fact]
    public void ReadTable_MatchesTheStaticMethodAndReturnsTheRealGrid()
    {
        var sut = new MarkdownEditorService();

        var table = sut.ReadTable(Sample, 0);

        Assert.Equal(MarkdownEditor.ReadTable(Sample, 0), table);
        Assert.Equal(new[] { "Region", "Total" }, table[0]);
        Assert.Equal(new[] { "North", "1200" }, table[1]);
    }

    [Fact]
    public void ReplaceSection_MatchesTheStaticMethodAndReplacesTheRightSection()
    {
        var sut = new MarkdownEditorService();

        var fromWrapper = sut.ReplaceSection(Sample, "Details", "Replaced.\n");

        Assert.Equal(
            MarkdownEditor.ReplaceSection(Sample, "Details", "Replaced.\n"),
            fromWrapper);
        Assert.Contains("Replaced.", fromWrapper);
        Assert.DoesNotContain("More text.", fromWrapper);
        // The OTHER section is untouched - a wrapper matching on the wrong heading would still
        // "replace something" but not leave this line alone.
        Assert.Contains("Intro text.", fromWrapper);
    }
}
