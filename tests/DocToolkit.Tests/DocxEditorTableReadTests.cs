using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using static DocToolkit.Tests.DocxFixtures;

namespace DocToolkit.Tests;

/// <summary>
/// Reading a DOCX table back as data (A23). Every assertion here is a LITERAL grid rather than a
/// count: a count assertion passes on any table at all, which is how a test ends up unable to fail.
/// </summary>
public class DocxEditorTableReadTests
{
    [Fact]
    public void ReadsTheLiteralGridIncludingTheHeaderRow()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Table(
                new[] { "Region", "Q1" },
                new[] { new object?[] { "EMEA", 1200 }, new object?[] { "APAC", 980 } }),
        });

        var table = DocxEditor.ReadTable(docx, 0);

        Assert.Equal(3, table.Count);
        Assert.Equal(new[] { "Region", "Q1" }, table[0]);
        Assert.Equal(new[] { "EMEA", "1200" }, table[1]);
        Assert.Equal(new[] { "APAC", "980" }, table[2]);
    }

    [Fact]
    public async Task RoundTripsWhatFillRowsWrote()
    {
        // The row's stated purpose: FillRows writes a table from data and nothing could read it
        // back, so the feature could not be verified from a consumer's side at all.
        var template = await HtmlToDocxConverter.ConvertAsync(
            """
            <table border="1">
              <tr><th>Desc</th><th>Qty</th></tr>
              <tr><td>{{item.Desc}}</td><td>{{item.Qty}}</td></tr>
            </table>
            """);

        var filled = DocxEditor.FillRows(template, "item", new[]
        {
            new Dictionary<string, string> { ["Desc"] = "Widget", ["Qty"] = "2" },
            new Dictionary<string, string> { ["Desc"] = "Gadget", ["Qty"] = "5" },
        });

        var table = DocxEditor.ReadTable(filled, 0);

        Assert.Equal(new[] { "Widget", "2" }, table[^2]);
        Assert.Equal(new[] { "Gadget", "5" }, table[^1]);
    }

    [Fact]
    public void CountsTablesAndIndexesThemInDocumentOrder()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Table(new[] { "first" }, new[] { new object?[] { "a" } }),
            DocxBlock.Paragraph("Between."),
            DocxBlock.Table(new[] { "second" }, new[] { new object?[] { "b" } }),
        });

        Assert.Equal(2, DocxEditor.TableCount(docx));

        // Index 1 must be the SECOND table. A test that only read index 0 would pass against an
        // implementation that ignored the index entirely.
        Assert.Equal(new[] { "second" }, DocxEditor.ReadTable(docx, 1)[0]);
        Assert.Equal(new[] { "first" }, DocxEditor.ReadTable(docx, 0)[0]);
    }

    [Fact]
    public void ANestedTableIsCellTextRatherThanATableOfItsOwn()
    {
        // Elements, not Descendants. Descendants would report this document as holding two tables
        // and let the inner one be indexed on its own, which is the TableRowFinder trap.
        var inner = Tbl(RowOf(P(R("inner"))));
        var outer = Tbl(new TableRow(new TableCell(P(R("outer")), inner)));
        var docx = DocxFixtures.Build(outer);

        Assert.Equal(1, DocxEditor.TableCount(docx));

        var cell = DocxEditor.ReadTable(docx, 0)[0][0];
        Assert.Contains("outer", cell, StringComparison.Ordinal);
        Assert.Contains("inner", cell, StringComparison.Ordinal);
    }

    [Fact]
    public void ARaggedRowComesBackShortRatherThanPadded()
    {
        // A horizontal merge is a row with fewer cells. Padding would invent a cell that is not in
        // the document, and a caller could not tell it from a real empty one.
        var docx = DocxFixtures.Build(Tbl(
            new TableRow(new TableCell(P(R("a"))), new TableCell(P(R("b")))),
            RowOf(P(R("only")))));

        var table = DocxEditor.ReadTable(docx, 0);

        Assert.Equal(2, table[0].Count);
        Assert.Single(table[1]);
    }

    [Fact]
    public void ACellHoldingSeveralParagraphsMatchesWhatExtractTextProduces()
    {
        var docx = DocxFixtures.Build(Tbl(
            RowOf(P(R("one")), P(R("two")))));

        var cell = DocxEditor.ReadTable(docx, 0)[0][0];

        // Same separator rule as ExtractText, because both go through BlockText.
        Assert.Equal("one\ntwo", cell);
    }

    [Fact]
    public void ANegativeIndexIsRejected()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Table(new[] { "only" }, new[] { new object?[] { "one" } }),
        });

        Assert.ThrowsAny<ArgumentOutOfRangeException>(() => DocxEditor.ReadTable(docx, -1));
    }

    [Fact]
    public void AnIndexPastTheLastTableIsRejectedAndTheMessageNamesTheCount()
    {
        // Split from the negative case deliberately: only this one can assert the count, because a
        // negative index is rejected by the guard before the document is opened. A single theory
        // with `if (index > 0)` around the message assertion would mean the -1 case silently
        // checked nothing but the exception type.
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Table(new[] { "only" }, new[] { new object?[] { "one" } }),
        });

        var ex = Assert.ThrowsAny<ArgumentOutOfRangeException>(() => DocxEditor.ReadTable(docx, 1));

        Assert.Contains("1 table", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ADocumentWithNoTablesCountsZeroAndCannotBeRead()
    {
        // Not an empty grid: an empty list is indistinguishable from a table that exists and has
        // no rows, and those mean different things.
        var docx = DocxEditor.Create(new[] { DocxBlock.Paragraph("No tables here.") });

        Assert.Equal(0, DocxEditor.TableCount(docx));
        Assert.ThrowsAny<ArgumentOutOfRangeException>(() => DocxEditor.ReadTable(docx, 0));
    }

    [Fact]
    public void UnreadableInputIsOneExceptionType()
    {
        var notADocx = System.Text.Encoding.UTF8.GetBytes("This is not a DOCX.");

        Assert.Throws<DocumentConversionException>(() => DocxEditor.TableCount(notADocx));
        Assert.Throws<DocumentConversionException>(() => DocxEditor.ReadTable(notADocx, 0));
    }

    [Fact]
    public async Task ReadsThroughAStream()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Table(new[] { "Region" }, new[] { new object?[] { "EMEA" } }),
        });

        using var forCount = new MemoryStream(docx, writable: false);
        Assert.Equal(1, await DocxEditor.TableCountAsync(forCount));

        using var forRead = new MemoryStream(docx, writable: false);
        var table = await DocxEditor.ReadTableAsync(forRead, 0);

        // The literal grid, not just a row count - the byte[] path is already proven, so what this
        // adds is that the Stream path produces the SAME data rather than merely succeeding.
        Assert.Equal(new[] { "Region" }, table[0]);
        Assert.Equal(new[] { "EMEA" }, table[1]);
    }

    [Fact]
    public async Task ReadsFromAPath()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Table(new[] { "Region" }, new[] { new object?[] { "EMEA" } }),
        });
        var path = Path.Join(Path.GetTempPath(), $"doctoolkit-{Guid.NewGuid():N}.docx");
        await File.WriteAllBytesAsync(path, docx);

        try
        {
            Assert.Equal(1, await DocxEditor.TableCountAsync(path));
            Assert.Equal(new[] { "EMEA" }, (await DocxEditor.ReadTableAsync(path, 0))[1]);
        }
        finally
        {
            File.Delete(path);
        }
    }
}
