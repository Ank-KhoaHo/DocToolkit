namespace DocToolkit.Tests;

public class DocxBlockTests
{
    [Fact]
    public void Heading_CarriesItsTextAndLevel()
    {
        var block = DocxBlock.Heading("Quarterly Report", 1);

        var heading = Assert.IsType<HeadingBlock>(block);
        Assert.Equal("Quarterly Report", heading.Text);
        Assert.Equal(1, heading.Level);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(7)]
    [InlineData(-1)]
    public void Heading_RejectsALevelOutsideOneToSix(int level)
        => Assert.Throws<ArgumentOutOfRangeException>(() => DocxBlock.Heading("Title", level));

    [Fact]
    public void Heading_RejectsNullText()
        => Assert.Throws<ArgumentNullException>(() => DocxBlock.Heading(null!, 1));

    [Fact]
    public void Paragraph_AcceptsEmptyText()
    {
        // An empty paragraph is a blank line, which is a legitimate thing to want.
        var paragraph = Assert.IsType<ParagraphBlock>(DocxBlock.Paragraph(""));
        Assert.Equal("", paragraph.Text);
    }

    [Fact]
    public void Table_MaterialisesHeadersAndRowsEagerly()
    {
        var rows = new List<IEnumerable<object?>> { new object?[] { "North", 1200 } };

        var table = Assert.IsType<TableBlock>(DocxBlock.Table(new[] { "Region", "Total" }, rows));

        // Materialised, not deferred: mutating the caller's list afterwards must not change the block.
        rows.Add(new object?[] { "South", 5 });
        Assert.Single(table.Rows);
        Assert.Equal(new[] { "Region", "Total" }, table.Headers);
    }

    [Fact]
    public void Table_RejectsANullRow()
        => Assert.Throws<ArgumentException>(() =>
            DocxBlock.Table(new[] { "A" }, new IEnumerable<object?>[] { null! }));

    [Fact]
    public void Table_RejectsNoHeaders()
        => Assert.Throws<ArgumentException>(() =>
            DocxBlock.Table(Array.Empty<string>(), Array.Empty<IEnumerable<object?>>()));

    [Fact]
    public void Image_RejectsBytesThatAreNeitherPngNorJpeg()
        => Assert.Throws<ArgumentException>(() => DocxBlock.Image(new byte[] { 1, 2, 3 }));

    [Fact]
    public void Image_RejectsEmptyBytes()
        => Assert.Throws<ArgumentException>(() => DocxBlock.Image(Array.Empty<byte>()));

    [Fact]
    public void Image_RejectsANonPositiveSize()
        => Assert.Throws<ArgumentOutOfRangeException>(
            () => DocxBlock.Image(ImageFixtures.Png(), widthPoints: 0));

    /// <summary>
    /// The closed hierarchy is the whole reason an unrenderable block cannot exist, so it is
    /// asserted rather than assumed. Deriving from outside this assembly does not compile; from
    /// inside it does, which is why this inspects accessibility rather than attempting it.
    ///
    /// This checks EVERY constructor, with no filtering. An earlier version of this type was a
    /// record, and a non-sealed record's compiler-generated copy constructor is necessarily
    /// public or protected (CS8878) - protected reaches derived types in ANY assembly, so an
    /// external caller could derive through it with `: base(seed)`. That hole was missed because
    /// the test filtered compiler-generated constructors out before asserting. Do not reintroduce
    /// that filter, and do not turn DocxBlock back into a record: either change makes this test
    /// fail, which is the point.
    /// </summary>
    [Fact]
    public void DocxBlock_CannotBeDerivedFromOutsideTheAssembly()
    {
        var constructors = typeof(DocxBlock).GetConstructors(
            System.Reflection.BindingFlags.Instance |
            System.Reflection.BindingFlags.Public |
            System.Reflection.BindingFlags.NonPublic);

        Assert.NotEmpty(constructors);

        // private protected (IsFamilyAndAssembly), internal (IsAssembly) and private are all
        // unreachable from another assembly. public, protected (IsFamily) and protected internal
        // (IsFamilyOrAssembly) are all reachable, and any of the three reopens the hierarchy.
        Assert.All(constructors, c => Assert.False(
            c.IsPublic || c.IsFamily || c.IsFamilyOrAssembly,
            $"{c} is externally reachable; a public, protected or protected-internal constructor " +
            "lets an external assembly derive from DocxBlock"));
    }
}
