using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// A113: every <c>Stream</c> overload added for the metadata, protection and formula members
/// agrees with the <c>byte[]</c> sibling it delegates to.
/// </summary>
/// <remarks>
/// <b>StreamOverloadTests cannot make this assertion, and that is why this file exists.</b> That
/// suite proves the stream PLUMBING — that a source is guarded and drained, that a destination is
/// guarded and written, that a forward-only stream is never sought, that the token is observed. All
/// of it still passes if an overload delegates to the WRONG method, or drops an argument, as long
/// as whatever it calls reads a source and writes a destination.
///
/// So each test here pins the RESULT against the <c>byte[]</c> overload on the same input, which is
/// the only thing that discriminates a correct delegation from a plausible one. Same reasoning as
/// the DI mirror tests, one layer down.
/// </remarks>
public class StreamOverloadParityTests
{
    private static byte[] Docx() => DocxEditor.Create(new[] { DocxBlock.Paragraph("body") });

    private static byte[] Xlsx() => WorkbookEditor.Create("Sales", new object?[][]
    {
        new object?[] { "Region", "Total" },
        new object?[] { "North", XlsxFormula.From("1+2") },
    });

    private static byte[] Pptx() => PresentationEditor.Create(new[] { PptxSlide.Titled("Deck") });

    private static byte[] Pdf() => DocxToPdfConverter.Convert(Docx());

    private static async Task<T> ReadThrough<T>(byte[] input, System.Func<Stream, Task<T>> call)
    {
        using var source = new MemoryStream(input, writable: false);
        return await call(source);
    }

    private static async Task<byte[]> WriteThrough(byte[] input, System.Func<Stream, Stream, Task> call)
    {
        using var source = new MemoryStream(input, writable: false);
        using var destination = new MemoryStream();
        await call(source, destination);
        return destination.ToArray();
    }

    // --- read-only members: the Stream overload returns what the byte[] one returns -------------

    [Fact]
    public async Task DocxEditor_ReadMetadataAsync_MatchesTheByteArrayOverload()
    {
        var docx = DocxEditor.WithMetadata(Docx(), new DocumentMetadata { Title = "T", Creator = "C" });

        var expected = DocxEditor.ReadMetadata(docx);
        var actual = await ReadThrough(docx, s => DocxEditor.ReadMetadataAsync(s));

        Assert.Equal(expected.Title, actual.Title);
        Assert.Equal(expected.Creator, actual.Creator);
    }

    [Fact]
    public async Task DocxEditor_IsProtectedAsync_MatchesTheByteArrayOverload()
    {
        var plain = Docx();
        var locked = DocxEditor.Protect(plain, "pw");

        Assert.False(await ReadThrough(plain, s => DocxEditor.IsProtectedAsync(s)));
        Assert.True(await ReadThrough(locked, s => DocxEditor.IsProtectedAsync(s)));
    }

    [Fact]
    public async Task WorkbookEditor_ReadMetadataAsync_MatchesTheByteArrayOverload()
    {
        var xlsx = WorkbookEditor.WithMetadata(Xlsx(), new DocumentMetadata { Title = "T", Creator = "C" });

        var expected = WorkbookEditor.ReadMetadata(xlsx);
        var actual = await ReadThrough(xlsx, s => WorkbookEditor.ReadMetadataAsync(s));

        Assert.Equal(expected.Title, actual.Title);
        Assert.Equal(expected.Creator, actual.Creator);
    }

    [Fact]
    public async Task WorkbookEditor_IsProtectedAsync_MatchesTheByteArrayOverload()
    {
        var plain = Xlsx();
        var locked = WorkbookEditor.Protect(plain, "pw");

        Assert.False(await ReadThrough(plain, s => WorkbookEditor.IsProtectedAsync(s)));
        Assert.True(await ReadThrough(locked, s => WorkbookEditor.IsProtectedAsync(s)));
    }

    [Fact]
    public async Task WorkbookEditor_InspectFormulasAsync_MatchesTheByteArrayOverload()
    {
        var xlsx = Xlsx();

        var expected = WorkbookEditor.InspectFormulas(xlsx);
        var actual = await ReadThrough(xlsx, s => WorkbookEditor.InspectFormulasAsync(s));

        // The count, not just "it returned something" - the fixture deliberately carries a formula,
        // so a delegation that inspected an empty workbook would report 0 and fail here.
        Assert.Equal(expected.TotalFormulas, actual.TotalFormulas);
        Assert.True(actual.TotalFormulas > 0);
    }

    [Fact]
    public async Task PresentationEditor_ReadMetadataAsync_MatchesTheByteArrayOverload()
    {
        var pptx = PresentationEditor.WithMetadata(Pptx(), new DocumentMetadata { Title = "T", Creator = "C" });

        var expected = PresentationEditor.ReadMetadata(pptx);
        var actual = await ReadThrough(pptx, s => PresentationEditor.ReadMetadataAsync(s));

        Assert.Equal(expected.Title, actual.Title);
        Assert.Equal(expected.Creator, actual.Creator);
    }

    [Fact]
    public async Task PresentationEditor_IsProtectedAsync_MatchesTheByteArrayOverload()
    {
        var plain = Pptx();
        var locked = PresentationEditor.Protect(plain, "pw");

        Assert.False(await ReadThrough(plain, s => PresentationEditor.IsProtectedAsync(s)));
        Assert.True(await ReadThrough(locked, s => PresentationEditor.IsProtectedAsync(s)));
    }

    [Fact]
    public async Task PdfEditor_ReadMetadataAsync_MatchesTheByteArrayOverload()
    {
        var pdf = PdfEditor.WithMetadata(Pdf(), new PdfMetadata { Title = "T", Author = "A" });

        var expected = PdfEditor.ReadMetadata(pdf);
        var actual = await ReadThrough(pdf, s => PdfEditor.ReadMetadataAsync(s));

        Assert.Equal(expected.Title, actual.Title);
        Assert.Equal(expected.Author, actual.Author);
    }

    // --- writing members: the bytes written match the bytes returned ---------------------------

    [Fact]
    public async Task DocxEditor_WithMetadataAsync_MatchesTheByteArrayOverload()
    {
        var docx = Docx();
        var metadata = new DocumentMetadata { Title = "Stamped" };

        var written = await WriteThrough(docx, (s, d) => DocxEditor.WithMetadataAsync(s, metadata, d));

        Assert.Equal("Stamped", DocxEditor.ReadMetadata(written).Title);
    }

    [Fact]
    public async Task WorkbookEditor_WithMetadataAsync_MatchesTheByteArrayOverload()
    {
        var xlsx = Xlsx();
        var metadata = new DocumentMetadata { Title = "Stamped" };

        var written = await WriteThrough(xlsx, (s, d) => WorkbookEditor.WithMetadataAsync(s, metadata, d));

        Assert.Equal("Stamped", WorkbookEditor.ReadMetadata(written).Title);
    }

    [Fact]
    public async Task WorkbookEditor_EvaluateFormulasAsync_MatchesTheByteArrayOverload()
    {
        var xlsx = Xlsx();

        var written = await WriteThrough(xlsx, (s, d) => WorkbookEditor.EvaluateFormulasAsync(s, d));

        // The cached value is the point of EvaluateFormulas, so assert the VALUE rather than that
        // bytes came out: a delegation that copied the source through would still produce a
        // readable workbook, and would still be wrong.
        Assert.Equal("3", WorkbookEditor.ReadCell(written, "Sales", "B2"));
    }

    [Fact]
    public async Task PresentationEditor_WithMetadataAsync_MatchesTheByteArrayOverload()
    {
        var pptx = Pptx();
        var metadata = new DocumentMetadata { Title = "Stamped" };

        var written = await WriteThrough(pptx, (s, d) => PresentationEditor.WithMetadataAsync(s, metadata, d));

        Assert.Equal("Stamped", PresentationEditor.ReadMetadata(written).Title);
    }

    [Fact]
    public async Task PdfEditor_WithMetadataAsync_MatchesTheByteArrayOverload()
    {
        var pdf = Pdf();
        var metadata = new PdfMetadata { Title = "Stamped" };

        var written = await WriteThrough(pdf, (s, d) => PdfEditor.WithMetadataAsync(s, metadata, d));

        Assert.Equal("Stamped", PdfEditor.ReadMetadata(written).Title);
    }
}
