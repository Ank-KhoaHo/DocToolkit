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
    internal static byte[] Locked(byte[] docx, LockingValues how, string tag = "FullName")
    {
        var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        ms.Position = 0;

        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            foreach (var properties in doc.MainDocumentPart!.Document!.Body!
                         .Descendants<SdtElement>()
                         .Where(c => c.Descendants<Tag>().Any(t => t.Val?.Value == tag))
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
    /// <remarks>
    /// <c>SdtLocked</c> is not here since A120 (2026-11-13). It means the control cannot be
    /// DELETED; its content stays editable, in Word and in OfficeIMO 3.5's own lock check. A119
    /// refused all three kinds alike without weighing that; <see cref="FillsAnSdtLockedControl_AndKeepsTheLock"/>
    /// pins the corrected behaviour.
    /// </remarks>
    [Theory]
    [InlineData("SdtContentLocked")]
    [InlineData("ContentLocked")]
    public void RefusesToFillALockedControl(string how)
    {
        var docx = Locked(DocxFormFixtures.Form(), ByName(how));

        var ex = Assert.Throws<InvalidOperationException>(
            () => DocxForm.Fill(docx, Name("OVERWRITTEN")));

        Assert.Contains("locks against editing", ex.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// <c>SdtLocked</c> stops the control being deleted, not its content being edited — so the fill
    /// goes through, and the lock is still there afterwards.
    /// </summary>
    [Fact]
    public void FillsAnSdtLockedControl_AndKeepsTheLock()
    {
        var docx = Locked(DocxFormFixtures.Form(), LockingValues.SdtLocked);

        var filled = DocxForm.Fill(docx, Name("Ada Lovelace"));

        Assert.Equal(
            "Ada Lovelace",
            Assert.Single(DocxForm.Inspect(filled).Fields, f => f.Key == "FullName").Value.Text);
        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var survivor = Assert.Single(doc.MainDocumentPart!.Document!.Body!
            .Descendants<DocumentFormat.OpenXml.Wordprocessing.Lock>());
        Assert.Equal(LockingValues.SdtLocked, survivor.Val!.Value);
    }

    internal static LockingValues ByName(string how) => how switch
    {
        "SdtContentLocked" => LockingValues.SdtContentLocked,
        "SdtLocked" => LockingValues.SdtLocked,
        "ContentLocked" => LockingValues.ContentLocked,
        _ => throw new ArgumentOutOfRangeException(nameof(how), how, null),
    };

    /// <summary>
    /// On OfficeIMO 3.5 the library refuses first; DocToolkit's message replaces its text and the
    /// upstream exception stays reachable for diagnosis.
    /// </summary>
    [Fact]
    public void RefusalKeepsUpstreamAsInnerException()
    {
        var docx = Locked(DocxFormFixtures.Form(), LockingValues.SdtContentLocked);

        var ex = Assert.Throws<InvalidOperationException>(
            () => DocxForm.Fill(docx, Name("OVERWRITTEN")));

        Assert.NotNull(ex.InnerException);
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
