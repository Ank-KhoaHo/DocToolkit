using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit.Tests;

public class DocxEditorFillRowsTests
{
    /// <summary>Row data keyed by bare field name, as <c>FillRows</c> expects.</summary>
    private static IReadOnlyDictionary<string, string> Rec(params (string Key, string Value)[] fields)
        => fields.ToDictionary(f => f.Key, f => f.Value, StringComparer.Ordinal);

    [Fact]
    public void FillRows_ProducesOneRowPerRecordInOrder()
    {
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.Row(DocxFixtures.R("Description")),
            DocxFixtures.Row(DocxFixtures.R("{{item.Desc}} x{{item.Qty}}"))));

        var filled = DocxEditor.FillRows(docx, "item", new[]
        {
            Rec(("Desc", "Widget"), ("Qty", "2")),
            Rec(("Desc", "Gadget"), ("Qty", "5")),
            Rec(("Desc", "Doohickey"), ("Qty", "1")),
        });

        var text = DocxEditor.ExtractText(filled);

        Assert.Contains("Widget x2", text);
        Assert.Contains("Gadget x5", text);
        Assert.Contains("Doohickey x1", text);
        Assert.DoesNotContain("{{item.", text);
        Assert.Contains("Description", text);   // the header row is untouched
        Assert.True(
            text.IndexOf("Widget", StringComparison.Ordinal) < text.IndexOf("Gadget", StringComparison.Ordinal),
            "records keep the order they were given in");
    }

    [Fact]
    public void FillRows_ProducesExactlyOneTableRowPerRecord()
    {
        var docx = DocxFixtures.Build(DocxFixtures.Tbl(
            DocxFixtures.Row(DocxFixtures.R("Description")),
            DocxFixtures.Row(DocxFixtures.R("{{item.Desc}}"))));

        var filled = DocxEditor.FillRows(docx, "item", new[]
        {
            Rec(("Desc", "Widget")),
            Rec(("Desc", "Gadget")),
        });

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var rows = doc.MainDocumentPart!.Document.Body!
            .ChildElements.OfType<Table>().Single()
            .ChildElements.OfType<TableRow>().ToList();

        Assert.Equal(3, rows.Count);          // header + two records; the template row is gone
        Assert.Contains("Description", rows[0].InnerText);
        Assert.Contains("Widget", rows[1].InnerText);
        Assert.Contains("Gadget", rows[2].InnerText);
    }
}
