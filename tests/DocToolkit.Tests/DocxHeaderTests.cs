using Xunit;

namespace DocToolkit.Tests;

public class DocxHeaderTests
{
    [Fact]
    public void Text_ProducesOneSegmentAndDefaultsToLeft()
    {
        var header = DocxHeader.Text("Contoso Ltd");

        Assert.Equal(HeaderAlignment.Left, header.Alignment);
        var segment = Assert.Single(header.Segments);
        Assert.Equal("Contoso Ltd", segment.ToString());
    }

    [Fact]
    public void Of_KeepsSegmentOrder()
    {
        var header = DocxHeader.Of(
            HeaderAlignment.Right,
            DocxHeaderSegment.Text("Page "),
            DocxHeaderSegment.PageNumber,
            DocxHeaderSegment.Text(" of "),
            DocxHeaderSegment.PageCount);

        Assert.Equal(HeaderAlignment.Right, header.Alignment);
        Assert.Equal(4, header.Segments.Count);
        Assert.Equal("Page ", header.Segments[0].ToString());
        Assert.Equal("{PAGE}", header.Segments[1].ToString());
        Assert.Equal("{NUMPAGES}", header.Segments[3].ToString());
    }

    [Fact]
    public void NullArgumentsAreRejectedByName()
    {
        Assert.Equal("text",
            Assert.Throws<ArgumentNullException>(() => DocxHeaderSegment.Text(null!)).ParamName);
        Assert.Equal("segments",
            Assert.Throws<ArgumentNullException>(() => DocxHeader.Of(HeaderAlignment.Left, null!)).ParamName);
    }

    [Fact]
    public void AHeaderWithNoSegmentsIsRejected()
    {
        // An empty header would emit a header part containing an empty paragraph - a blank line
        // reserved at the top of every page for nothing. Callers who want no header omit it.
        Assert.Throws<ArgumentException>(() => DocxHeader.Of(HeaderAlignment.Left));
    }

    [Fact]
    public void SegmentsAreNotAliasedToTheCallersArray()
    {
        var segments = new[] { DocxHeaderSegment.Text("one") };
        var header = DocxHeader.Of(HeaderAlignment.Left, segments);

        segments[0] = DocxHeaderSegment.Text("mutated");

        Assert.Equal("one", header.Segments[0].ToString());
    }

    [Fact]
    public void AnUndefinedAlignmentIsRejected()
    {
        Assert.Throws<ArgumentOutOfRangeException>(
            () => DocxHeader.Text("x", (HeaderAlignment)42));
    }
}
