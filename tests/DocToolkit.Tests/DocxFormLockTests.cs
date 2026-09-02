using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// <see cref="DocxForm.Fill"/> refuses a control the document locks against editing (A119).
///
/// <b>Measured before it was built.</b> The old behaviour did not merely ignore the lock — it wrote
/// through it and left the <c>w:lock</c> in place, so the document came back declaring a control
/// protected while its content had been replaced. Nothing in the file recorded that, and no caller
/// could detect it.
///
/// The lock is attached to a control OfficeIMO authored. <c>DocxFormFixtures</c> records why that
/// matters: a hand-built <c>SdtBlock</c> is not a typed control, and measurements taken against one
/// are measurements of the fixture.
/// </summary>
public class DocxFormLockTests
{
    private static byte[] Locked(byte[] docx, LockingValues how)
    {
        var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        ms.Position = 0;

        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            foreach (var properties in doc.MainDocumentPart!.Document!.Body!
                         .Descendants<SdtElement>()
                         .Where(c => c.Descendants<Tag>().Any(t => t.Val?.Value == "FullName"))
                         .Select(c => c.Elements<SdtProperties>().FirstOrDefault())
                         .OfType<SdtProperties>())
            {
                properties.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Lock { Val = how });
            }

            doc.MainDocumentPart.Document.Save();
        }

        return ms.ToArray();
    }

    private static Dictionary<string, DocxFormValue> Name(string value) =>
        new() { ["FullName"] = DocxFormValue.FromText(value) };

    /// <summary>
    /// Keyed by NAME rather than by value, because <c>LockingValues</c> is a struct in OpenXML SDK
    /// 3.x rather than a C# enum, so it cannot be an <c>InlineData</c> constant.
    /// </summary>
    [Theory]
    [InlineData("SdtContentLocked")]
    [InlineData("SdtLocked")]
    [InlineData("ContentLocked")]
    public void RefusesToFillALockedControl(string how)
    {
        var kind = how switch
        {
            "SdtContentLocked" => LockingValues.SdtContentLocked,
            "SdtLocked" => LockingValues.SdtLocked,
            _ => LockingValues.ContentLocked,
        };
        var docx = Locked(DocxFormFixtures.Form(), kind);

        var ex = Assert.Throws<InvalidOperationException>(
            () => DocxForm.Fill(docx, Name("OVERWRITTEN")));

        Assert.Contains("locks against editing", ex.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// The positive control, and it is load-bearing rather than decorative: a <c>Fill</c> that threw
    /// on EVERY input would satisfy every assertion above and nothing else here would notice.
    /// </summary>
    [Fact]
    public void PositiveControl_AnUnlockedControlIsStillFilled()
    {
        var filled = DocxForm.Fill(DocxFormFixtures.Form(), Name("Ada Lovelace"));

        Assert.Equal(
            "Ada Lovelace",
            Assert.Single(DocxForm.Inspect(filled).Fields, f => f.Key == "FullName").Value.Text);
    }

    /// <summary>
    /// The refusal is about the control the fill would CHANGE, not about the document containing a
    /// lock anywhere. A guard that refused on any locked control present would break every template
    /// that locks a heading, which is most of them.
    /// </summary>
    [Fact]
    public void FillsAnUnlockedControlInADocumentThatAlsoHasALockedOne()
    {
        var docx = Locked(DocxFormFixtures.Form(), LockingValues.SdtContentLocked);

        var filled = DocxForm.Fill(
            docx, new Dictionary<string, DocxFormValue> { ["Plan"] = DocxFormValue.FromChoice("Team") });

        Assert.Equal(
            "Team",
            Assert.Single(DocxForm.Inspect(filled).Fields, f => f.Key == "Plan").Value.Text);
    }

    [Fact]
    public void TheDocumentIsUnchangedWhenTheFillIsRefused()
    {
        // Not merely "it threw": the caller must not be handed a half-written document, and the
        // original bytes must still hold the original value.
        var docx = Locked(DocxFormFixtures.Form(), LockingValues.SdtContentLocked);

        Assert.Throws<InvalidOperationException>(() => DocxForm.Fill(docx, Name("OVERWRITTEN")));

        Assert.Equal(
            "Khoa Ho",
            Assert.Single(DocxForm.Inspect(docx).Fields, f => f.Key == "FullName").Value.Text);
    }

    [Fact]
    public async Task TheStreamOverloadRefusesToo_AndWritesNothing()
    {
        var docx = Locked(DocxFormFixtures.Form(), LockingValues.SdtContentLocked);
        using var source = new MemoryStream(docx, writable: false);
        using var destination = new MemoryStream();

        await Assert.ThrowsAsync<InvalidOperationException>(
            () => DocxForm.FillAsync(source, destination, Name("OVERWRITTEN")));

        Assert.Equal(0, destination.Length);
    }

    /// <summary>
    /// The lock survives a refused fill, which is the state the old behaviour corrupted: it kept
    /// the lock AND changed the content, so the file contradicted itself.
    /// </summary>
    [Fact]
    public void TheLockItselfIsNeverRemoved()
    {
        var docx = Locked(DocxFormFixtures.Form(), LockingValues.SdtContentLocked);

        Assert.Throws<InvalidOperationException>(() => DocxForm.Fill(docx, Name("OVERWRITTEN")));

        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Single(doc.MainDocumentPart!.Document!.Body!
            .Descendants<DocumentFormat.OpenXml.Wordprocessing.Lock>());
    }
}
