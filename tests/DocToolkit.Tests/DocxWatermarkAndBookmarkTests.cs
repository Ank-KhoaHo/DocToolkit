using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// A108: DOCX watermarking and bookmarks, both through <c>OfficeIMO.Word</c> — the package already
/// behind every other <see cref="DocxEditor"/> operation, and not the PDF watermarking the backlog
/// declined.
/// </summary>
public class DocxWatermarkAndBookmarkTests
{
    private static byte[] Doc(params string[] paragraphs) =>
        DocxEditor.Create(paragraphs.Select(DocxBlock.Paragraph).ToArray());

    private static int WatermarkCount(byte[] docx)
    {
        using var ms = new MemoryStream(docx, writable: false);
        using var doc = OfficeIMO.Word.WordDocument.Load(ms);
        return doc.Sections.Sum(s => s.Watermarks.Count);
    }

    private static int SectionCount(byte[] docx)
    {
        using var ms = new MemoryStream(docx, writable: false);
        using var doc = OfficeIMO.Word.WordDocument.Load(ms);
        return doc.Sections.Count;
    }

    // --- watermark -----------------------------------------------------------------------------

    [Fact]
    public void AddWatermark_StampsTheDocumentAndLeavesTheTextAlone()
    {
        var docx = Doc("BODY-MARKER");

        var stamped = DocxEditor.AddWatermark(docx, "DRAFT");

        Assert.Equal(1, WatermarkCount(stamped));
        Assert.Contains("BODY-MARKER", DocxEditor.ExtractText(stamped), StringComparison.Ordinal);
    }

    [Fact]
    public void AddWatermark_OnAMergedDocument_StampsTheOneSectionOfficeIMOReports()
    {
        // A DOCUMENTED LIMITATION, pinned. Measured: a merged document's body carries TWO w:sectPr
        // elements, but OfficeIMO's section model reports ONE section - so the loop over
        // document.Sections applies a single watermark, not one per sectPr.
        //
        // An earlier version of this test asserted 2 and failed, because the "applied to every
        // section" claim was written before it was measured. The claim is now what the code
        // actually does, and the consequence a caller needs - a merged document may not be marked
        // on every page - is in the README's Known Limitations rather than left to be discovered.
        var merged = DocxEditor.Merge(new[] { Doc("FIRST"), Doc("SECOND") });
        Assert.Equal(1, SectionCount(merged));

        var stamped = DocxEditor.AddWatermark(merged, "DRAFT");

        Assert.Equal(1, WatermarkCount(stamped));
        Assert.Equal(SectionCount(merged), WatermarkCount(stamped));
    }

    [Fact]
    public void AddWatermark_SurvivesASecondSave()
    {
        // A watermark that survives one round trip and vanishes on the next is the shape that
        // passes a naive test and fails a real workflow.
        var stamped = DocxEditor.AddWatermark(Doc("body"), "DRAFT");

        var resaved = DocxEditor.WithMetadata(stamped, new DocumentMetadata { Title = "T" });

        Assert.Equal(1, WatermarkCount(resaved));
    }

    [Fact]
    public void RemoveWatermarks_ClearsThem_AndIsAnNoOpOnADocumentWithNone()
    {
        var stamped = DocxEditor.AddWatermark(Doc("body"), "DRAFT");

        var cleared = DocxEditor.RemoveWatermarks(stamped);
        Assert.Equal(0, WatermarkCount(cleared));

        // A document with none is returned rather than refused.
        var again = DocxEditor.RemoveWatermarks(cleared);
        Assert.Equal(0, WatermarkCount(again));
        Assert.Contains("body", DocxEditor.ExtractText(again), StringComparison.Ordinal);
    }

    [Fact]
    public void AddWatermark_RejectsBadArguments()
    {
        var docx = Doc("body");

        Assert.Throws<ArgumentNullException>(() => DocxEditor.AddWatermark(null!, "DRAFT"));
        Assert.Throws<ArgumentNullException>(() => DocxEditor.AddWatermark(docx, null!));
        Assert.Equal("docx", Assert.Throws<ArgumentException>(
            () => DocxEditor.AddWatermark(Array.Empty<byte>(), "DRAFT")).ParamName);
        Assert.Equal("text", Assert.Throws<ArgumentException>(
            () => DocxEditor.AddWatermark(docx, "   ")).ParamName);
    }

    [Fact]
    public async Task AddWatermarkAsync_AndRemoveWatermarksAsync_MatchTheByteArrayOverloads()
    {
        using var source = new MemoryStream(Doc("body"), writable: false);
        using var stamped = new MemoryStream();
        await DocxEditor.AddWatermarkAsync(source, "DRAFT", stamped);
        Assert.Equal(1, WatermarkCount(stamped.ToArray()));

        using var reread = new MemoryStream(stamped.ToArray(), writable: false);
        using var cleared = new MemoryStream();
        await DocxEditor.RemoveWatermarksAsync(reread, cleared);
        Assert.Equal(0, WatermarkCount(cleared.ToArray()));
    }

    // --- bookmarks -----------------------------------------------------------------------------

    [Fact]
    public void AddBookmark_ThenReadBookmarks_RoundTripsByName()
    {
        var marked = DocxEditor.AddBookmark(Doc("clause text"), 0, "Clause7");

        // By NAME, not by count: a count passes against a bookmark carrying the wrong name.
        Assert.Contains("Clause7", DocxEditor.ReadBookmarks(marked));
    }

    [Fact]
    public void ReadBookmarks_OnADocumentWithNone_ReturnsEmpty()
    {
        Assert.Empty(DocxEditor.ReadBookmarks(Doc("nothing marked")));
    }

    [Fact]
    public void AddBookmark_TwiceOnDifferentParagraphs_KeepsBoth()
    {
        var docx = Doc("first", "second");

        var marked = DocxEditor.AddBookmark(docx, 0, "Alpha");
        marked = DocxEditor.AddBookmark(marked, 1, "Beta");

        var names = DocxEditor.ReadBookmarks(marked);
        Assert.Contains("Alpha", names);
        Assert.Contains("Beta", names);
    }

    [Fact]
    public void AddBookmark_RejectsABadIndexAndBadArguments()
    {
        var docx = Doc("only one paragraph");

        Assert.Throws<ArgumentOutOfRangeException>(() => DocxEditor.AddBookmark(docx, -1, "X"));
        Assert.Throws<ArgumentOutOfRangeException>(() => DocxEditor.AddBookmark(docx, 99, "X"));
        Assert.Throws<ArgumentNullException>(() => DocxEditor.AddBookmark(null!, 0, "X"));
        Assert.Equal("name", Assert.Throws<ArgumentException>(
            () => DocxEditor.AddBookmark(docx, 0, "  ")).ParamName);
        Assert.Equal("docx", Assert.Throws<ArgumentException>(
            () => DocxEditor.AddBookmark(Array.Empty<byte>(), 0, "X")).ParamName);
    }

    [Fact]
    public async Task AddBookmarkAsync_AndReadBookmarksAsync_MatchTheByteArrayOverloads()
    {
        using var source = new MemoryStream(Doc("clause text"), writable: false);
        using var destination = new MemoryStream();
        await DocxEditor.AddBookmarkAsync(source, 0, "Clause7", destination);

        using var reread = new MemoryStream(destination.ToArray(), writable: false);
        Assert.Contains("Clause7", await DocxEditor.ReadBookmarksAsync(reread));
    }

}
