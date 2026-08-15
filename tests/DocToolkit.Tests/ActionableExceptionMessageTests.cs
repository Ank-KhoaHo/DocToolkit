using System.Text;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using WordDocument = DocumentFormat.OpenXml.Wordprocessing.Document;

namespace DocToolkit.Tests;

/// <summary>
/// Pins the FULL text of the <see cref="DocumentConversionException"/> messages that carry
/// hand-written, specific guidance (added by "make DocumentConversionException messages
/// actionable").
///
/// Every other test in this suite that touches these messages deliberately asserts
/// <c>Assert.Contains</c> against the ORIGINAL first sentence — that is the part a consumer might
/// grep their logs for, and preserving it verbatim was the whole point of that change. But a
/// substring check cannot see anything past the substring: a typo, a dropped word, or a stale API
/// name (<c>ReadCell</c>, <c>SheetNames</c>) in the guidance clause that follows would ship with
/// every one of those tests still green. This repo has already paid for exactly that shape of gap
/// once — see B11-FLAKE's <c>AssertOutcome</c> and 19 substring-only tests elsewhere in this
/// codebase's history. These tests close it for the messages added here: <c>Assert.Equal</c>
/// against the whole string, so a one-word slip fails loudly here even though nothing else notices.
/// </summary>
public class ActionableExceptionMessageTests
{
    // =============================================================================================
    // The truncated-image family (ImageInspector) — every guidance sentence in that file.
    // =============================================================================================

    [Fact]
    public void Png_TruncatedIhdrMessageIsExact()
    {
        // 23 bytes: enough to be recognised as PNG (8-byte signature), one short of a complete
        // 24-byte IHDR.
        var truncated = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A }
            .Concat(new byte[15])
            .ToArray();

        var ex = Assert.Throws<DocumentConversionException>(() => ImageInspector.Inspect(truncated));

        Assert.Equal(
            "PNG image is truncated: it has no complete IHDR chunk. The image bytes are "
            + "incomplete — check the file was fully read or uploaded before it was passed in.",
            ex.Message);
    }

    [Fact]
    public void Jpeg_TruncatedFrameHeaderMessageIsExact()
    {
        // SOI, then an SOF0 marker whose length/precision bytes are present but whose
        // height/width bytes never arrive.
        var truncated = new byte[] { 0xFF, 0xD8, 0xFF, 0xC0, 0x00, 0x11, 0x08, 0x00 };

        var ex = Assert.Throws<DocumentConversionException>(() => ImageInspector.Inspect(truncated));

        Assert.Equal(
            "JPEG image is truncated inside its frame header. The image bytes are "
            + "incomplete — check the file was fully read or uploaded before it was passed in.",
            ex.Message);
    }

    [Fact]
    public void Jpeg_TruncatedSegmentHeaderMessageIsExact()
    {
        // SOI, then a non-standalone marker whose two length bytes never arrive.
        var truncated = new byte[] { 0xFF, 0xD8, 0xFF, 0xE0, 0x00 };

        var ex = Assert.Throws<DocumentConversionException>(() => ImageInspector.Inspect(truncated));

        Assert.Equal(
            "JPEG image is truncated inside a segment header. The image bytes are "
            + "incomplete — check the file was fully read or uploaded before it was passed in.",
            ex.Message);
    }

    /// <summary>
    /// This is the message Finding 1 rewrote. The scan loop exhausting the buffer with no
    /// Start-Of-Frame does NOT prove truncation — a complete-but-malformed file (a segment length
    /// landing exactly on <c>image.Length</c>, or bytes that never carry a real SOF) reaches the
    /// same line with nothing cut short. The reworded guidance names both possibilities rather than
    /// asserting truncation as fact, which is exactly the mistake this repo's timeout-message
    /// history (see CLAUDE.md, "A timeout message must not name a cause the timeout cannot
    /// distinguish") already paid for once.
    /// </summary>
    [Fact]
    public void Jpeg_NoStartOfFrameMessageIsExact_AndNamesBothCandidatesWithoutAssertingEither()
    {
        // SOI, then a comment segment (COM, 0xFE) and nothing else: well-formed JPEG bytes that
        // simply never carry a frame header. Not truncated - complete and malformed.
        var wellFormedButNoFrame = new byte[] { 0xFF, 0xD8, 0xFF, 0xFE, 0x00, 0x04, 0x00, 0x00 };

        var ex = Assert.Throws<DocumentConversionException>(
            () => ImageInspector.Inspect(wellFormedButNoFrame));

        Assert.Equal(
            "JPEG image has no Start-Of-Frame segment, so its size cannot be determined. This "
            + "can mean the image bytes are truncated, or that the file is not actually a "
            + "well-formed JPEG — check both rather than assuming either.",
            ex.Message);
    }

    // =============================================================================================
    // The structural family (DocxEditor / PresentationEditor) — "no main part" / "no body" /
    // "no presentation part". None of these three throw sites had ANY test coverage before this
    // file: every existing DocxEditor/PresentationEditor test either supplies a well-formed package
    // or garbage bytes too broken to open at all (which is wrapped as "Failed to ... See the inner
    // exception for details." instead). Reaching these lines needs a package that OPENS
    // successfully via WordprocessingDocument.Open/PresentationDocument.Open but is missing the
    // part the code then dereferences - constructed directly with the OpenXml SDK below.
    // =============================================================================================

    [Fact]
    public void Docx_NoMainPartMessageIsExact()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            // Deliberately never call doc.AddMainDocumentPart() - a package that opens fine but
            // owns no main document part at all.
        }

        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxEditor.ReplaceText(ms.ToArray(), new Dictionary<string, string>()));

        Assert.Equal(
            "Document has no main part. This usually means the file is not really a .docx (for "
            + "example it was renamed from another format) or the upload is corrupt.",
            ex.Message);
    }

    /// <summary>
    /// The identical literal is duplicated at a second throw site, <c>ReplaceImageCore</c> - not
    /// shared through one constant. The two copies could drift independently, so both are pinned
    /// rather than trusting that testing one covers the other.
    /// </summary>
    [Fact]
    public void Docx_NoMainPartMessageIsExact_ViaReplaceImage()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
        }

        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxEditor.ReplaceImage(ms.ToArray(), "{{logo}}", ImageFixtures.Png()));

        Assert.Equal(
            "Document has no main part. This usually means the file is not really a .docx (for "
            + "example it was renamed from another format) or the upload is corrupt.",
            ex.Message);
    }

    [Fact]
    public void Docx_NoBodyMessageIsExact()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            // A w:document element with no w:body child - present, but empty of the one thing
            // ReplaceTextCore actually needs.
            main.Document = new WordDocument();
            main.Document.Save();
        }

        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxEditor.ReplaceText(ms.ToArray(), new Dictionary<string, string>()));

        Assert.Equal(
            "Document has no body. This usually means the file is not really a .docx (for "
            + "example it was renamed from another format) or the upload is corrupt.",
            ex.Message);
    }

    [Fact]
    public void Presentation_NoPresentationPartMessageIsExact()
    {
        using var ms = new MemoryStream();
        using (var doc = PresentationDocument.Create(ms, PresentationDocumentType.Presentation))
        {
            // Deliberately never call doc.AddPresentationPart().
        }

        var ex = Assert.Throws<DocumentConversionException>(
            () => PresentationEditor.ExtractText(ms.ToArray()));

        Assert.Equal(
            "Presentation has no presentation part. This usually means the file is not really "
            + "a .pptx (for example it was renamed from another format) or the upload is corrupt.",
            ex.Message);
    }

    // =============================================================================================
    // The WorkbookEditor family — ReadCell and SheetNames guidance. Named explicitly because these
    // are the two messages a future rename of either method would silently invalidate under a
    // Contains-only check.
    // =============================================================================================

    [Fact]
    public void Workbook_WorksheetNotFoundMessagePointsAtSheetNames()
    {
        var xlsx = WorkbookEditor.Create("Sales", new[] { new object?[] { "only" } });

        var ex = Assert.Throws<DocumentConversionException>(
            () => WorkbookEditor.ReadCell(xlsx, "Nope", "A1"));

        Assert.Equal(
            "Worksheet 'Nope' was not found. Call WorkbookEditor.SheetNames to see what is available.",
            ex.Message);
    }

    [Fact]
    public void Workbook_CellLimitExceededMessagePointsAtReadCell()
    {
        // Same construction as WorkbookEditorTests.ReadSheet_ThrowsRatherThanAllocateForAFarFlungStrayValue:
        // one stray value at Excel's own maximum address describes a 1,048,576 x 16,384 rectangle
        // without the file itself needing to be anywhere near that large.
        using var workbook = new XLWorkbook();
        var sheet = workbook.Worksheets.Add("Sales");
        sheet.Cell("A1").Value = "a";
        sheet.Cell("XFD1048576").Value = "x";
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);

        var ex = Assert.Throws<DocumentConversionException>(
            () => WorkbookEditor.ReadSheet(ms.ToArray(), "Sales"));

        Assert.Equal(
            "Sheet 'Sales' spans 1048576 rows x 16384 columns (17179869184 cells), which exceeds "
            + "the 2000000-cell limit ReadSheet will materialise. Read specific cells with "
            + "ReadCell instead.",
            ex.Message);
    }

    // =============================================================================================
    // The PDF-open family (PdfEditor.Open / PdfTextExtractor.Pages) — the identical literal is
    // duplicated in both files rather than shared, for the reason PdfTextExtractor's own class
    // comment gives (PdfDocument exists in both PdfSharp.Pdf and UglyToad.PdfPig, so merging the
    // two would force an alias onto every line). Both copies are pinned so they cannot drift.
    // =============================================================================================

    [Fact]
    public void Pdf_FailedToReadMessageIsExact_ViaPageCount()
    {
        var notAPdf = Encoding.UTF8.GetBytes("This is not a PDF.");

        var ex = Assert.Throws<DocumentConversionException>(() => PdfEditor.PageCount(notAPdf));

        Assert.Equal(
            "Failed to read the PDF. This usually means the PDF is password-protected, "
            + "truncated, or not actually a PDF — check the source bytes.",
            ex.Message);
    }

    [Fact]
    public void Pdf_FailedToReadMessageIsExact_ViaExtractText()
    {
        var notAPdf = Encoding.UTF8.GetBytes("This is not a PDF.");

        var ex = Assert.Throws<DocumentConversionException>(() => PdfEditor.ExtractText(notAPdf));

        Assert.Equal(
            "Failed to read the PDF. This usually means the PDF is password-protected, "
            + "truncated, or not actually a PDF — check the source bytes.",
            ex.Message);
    }
}
