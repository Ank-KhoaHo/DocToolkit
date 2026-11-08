using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>DOCX to Markdown repairs for what OfficeIMO 3.5 drops (A123).</summary>
public class DocxToMarkdownConverterTests
{
    private static readonly Regex DataImage = new(@"data:image/[a-z]+;base64,[A-Za-z0-9+/=]+", RegexOptions.Compiled);

    /// <summary>Two pictures of different sizes, both inside ONE run.</summary>
    private static byte[] TwoPicturesInOneRun()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Image(ImageFixtures.Png(10, 12), 40, 40, "first"),
            DocxBlock.Image(ImageFixtures.Png(30, 20), 40, 40, "second"),
        });

        using var ms = new MemoryStream();
        ms.Write(docx);
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            var drawings = doc.MainDocumentPart!.Document!.Body!.Descendants<Drawing>().ToList();
            Assert.Equal(2, drawings.Count);
            var second = drawings[1];
            var secondRun = second.Parent!;
            drawings[0].Parent!.AppendChild(second.CloneNode(true));
            secondRun.Remove();
            doc.MainDocumentPart.Document.Save();
        }

        return ms.ToArray();
    }

    private static string[] Images(string markdown) =>
        DataImage.Matches(markdown).Select(m => m.Value).ToArray();

    private static int MaxDrawingsInARun(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.Document!.Body!.Descendants<Run>()
            .Max(r => r.Elements<Drawing>().Count());
    }

    [Fact]
    public void Fixture_HoldsTwoDrawingsInOneRun()
    {
        Assert.Equal(2, MaxDrawingsInARun(TwoPicturesInOneRun()));
    }

    [Fact]
    public void OfficeImo35_KeepsOnlyTheFirstDrawingOfARun()
    {
        // The measured cause, pinned: raw 3.5 export of the unrepaired document.
        using var ms = new MemoryStream(TwoPicturesInOneRun());
        using var word = WordDocument.Load(ms);
        var raw = word.ToMarkdown(TextExportOptions.ForMarkdown());

        Assert.True(Images(raw).Length == 1, "OfficeIMO fixed this upstream; retire DrawingRunSplit");
    }

    [Fact]
    public void TwoImagesInOneRunBothReachTheMarkdown()
    {
        var images = Images(DocxToMarkdownConverter.Convert(TwoPicturesInOneRun()));

        Assert.Equal(2, images.Length);
        Assert.NotEqual(images[0], images[1]);
    }

    [Fact]
    public void TwoImagesInOneRunBothReachTheMarkdown_ThroughConvertWithReport()
    {
        var result = DocxToMarkdownConverter.ConvertWithReport(TwoPicturesInOneRun());

        var images = Images(result.Value);
        Assert.Equal(2, images.Length);
        Assert.NotEqual(images[0], images[1]);
    }

    [Fact]
    public void DrawingRunSplit_GivesEachDrawingItsOwnRun()
    {
        var repaired = DrawingRunSplit.Apply(TwoPicturesInOneRun());

        Assert.Equal(1, MaxDrawingsInARun(repaired));
    }

    [Fact]
    public void DrawingRunSplit_OneDrawingPerRun_ReturnsTheSameArray()
    {
        var docx = DocxEditor.Create(new[]
        {
            DocxBlock.Image(ImageFixtures.Png(10, 12), 40, 40, "first"),
            DocxBlock.Image(ImageFixtures.Png(30, 20), 40, 40, "second"),
        });

        Assert.Same(docx, DrawingRunSplit.Apply(docx));
    }
}
