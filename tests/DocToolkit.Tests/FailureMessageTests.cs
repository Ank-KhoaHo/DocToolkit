namespace DocToolkit.Tests;

/// <summary>
/// Two failure messages that named a cause they could not distinguish, and now name the real one.
///
/// <b>Both were found by running the library over real files rather than fixtures</b> - 62 legacy
/// `.xls` and 200 PDFs from a public crawl of the .gov domain, measured 2026-08-17. Neither is a
/// behaviour change: the same calls fail on the same inputs. Only what the caller is told changed.
///
/// This repository has now recorded this defect four times, which is why these are pinned by
/// <b>exact match</b> rather than by substring: the first sentence is what somebody greps their
/// logs for, and the guidance after it is what a substring check cannot see.
/// </summary>
public class FailureMessageTests
{
    private static byte[] LegacyXls() =>
        File.ReadAllBytes(Path.Join(AppContext.BaseDirectory, "assets", "legacy.xls"));

    // ---- A43: a valid .xls reported as corrupt ---------------------------------------------------

    [Fact]
    public void TheXlsFixture_IsARealExcel97Workbook_NotARenamedXlsx()
    {
        // Guards the premise. If this were regenerated as an .xlsx, the test below would pass while
        // exercising nothing - the message it pins only appears for a compound file.
        var b = LegacyXls();
        Assert.Equal(new byte[] { 0xD0, 0xCF, 0x11, 0xE0 }, b.Take(4));
        Assert.NotEqual(new byte[] { 0x50, 0x4B }, b.Take(2));
    }

    [Fact]
    public void ALegacyXls_IsNotCalledCorrupt_AndIsToldWhatToDo()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => WorkbookEditor.SheetNames(LegacyXls()));

        Assert.Equal(
            "This is not an .xlsx package. The bytes are a compound file, which means either a "
            + "legacy Excel 97-2003 .xls workbook - save it as .xlsx to read it here - or an "
            + "encrypted .xlsx, which WorkbookEditor.Unprotect will open with its password.",
            ex.Message);

        // The old message was "Failed to read XLSX" wrapping "File contains corrupted data", which
        // sent people to check a file that was fine.
        Assert.DoesNotContain("corrupt", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void EveryWorkbookEntryPoint_SaysTheSameThing_AboutTheSameFile()
    {
        // Ten call sites wrapped the old message. Centralising the check means they cannot disagree
        // - and this is what would catch it if one were later given its own copy.
        var xls = LegacyXls();
        var messages = new[]
        {
            Assert.Throws<DocumentConversionException>(() => WorkbookEditor.SheetNames(xls)).Message,
            Assert.Throws<DocumentConversionException>(() => WorkbookEditor.ReadCell(xls, "Sales", "A1")).Message,
            Assert.Throws<DocumentConversionException>(() => WorkbookEditor.ReadSheet(xls, "Sales")).Message,
            Assert.Throws<DocumentConversionException>(() => XlsxToCsvConverter.Convert(xls, "Sales")).Message,
        };

        Assert.Single(messages.Distinct(StringComparer.Ordinal));
    }

    [Fact]
    public void ARealXlsx_IsUnaffected()
    {
        // The control. Without it, a check that rejected every compound file would look identical
        // to one that rejected everything.
        var xlsx = WorkbookEditor.Create("Sales", new[] { new object?[] { "ok" } });

        Assert.Equal(new[] { "Sales" }, WorkbookEditor.SheetNames(xlsx));
    }

    // ---- A45: a readable PDF reported as unreadable ----------------------------------------------

    private static byte[] OwnerRestricted() => PdfEditor.Protect(
        DocxToPdfConverter.Convert(DocxEditor.Create(new[] { DocxBlock.Paragraph("RESTRICTED") })),
        new PdfProtection { OwnerPassword = "owner-pw", AllowPrinting = false });

    [Fact]
    public void APermissionRestrictedPdf_StillReads_WhichIsWhyTheOldMessageWasWrong()
    {
        // The premise of the whole row: reading is unaffected. If this ever starts throwing, the
        // message below becomes correct and the test above it becomes wrong.
        var restricted = OwnerRestricted();

        Assert.Equal(1, PdfEditor.PageCount(restricted));
        Assert.Contains("RESTRICTED", string.Concat(PdfEditor.ExtractText(restricted)),
            StringComparison.Ordinal);
    }

    [Fact]
    public void APermissionRestrictedPdf_SaysSo_AndNamesUnprotect()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => PdfEditor.WithMetadata(OwnerRestricted(), new PdfMetadata { Title = "t" }));

        Assert.Equal(
            "This PDF is permission-restricted: it carries an owner password, so it can be read "
            + "but not modified. Reading it works - PageCount and ExtractText are unaffected. To "
            + "change it, call PdfEditor.Unprotect with the owner password first.",
            ex.Message);
    }

    [Fact]
    public void AGenuinelyUnreadablePdf_KeepsTheOriginalMessage()
    {
        // The control that stops the new branch swallowing real failures: bytes that are not a PDF
        // must still be reported as unreadable, not as permission-restricted.
        var junk = System.Text.Encoding.UTF8.GetBytes(new string('x', 4096));

        var ex = Assert.Throws<DocumentConversionException>(
            () => PdfEditor.WithMetadata(junk, new PdfMetadata { Title = "t" }));

        Assert.StartsWith("Failed to read the PDF.", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnOrdinaryPdf_IsUnaffected()
    {
        var pdf = DocxToPdfConverter.Convert(DocxEditor.Create(new[] { DocxBlock.Paragraph("plain") }));

        Assert.Equal(1, PdfEditor.PageCount(PdfEditor.WithMetadata(pdf, new PdfMetadata { Title = "t" })));
    }
}
