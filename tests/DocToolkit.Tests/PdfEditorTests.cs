using System.Text;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// The PDF utilities: page count, merge, page extraction and metadata.
///
/// Every input here is a PDF this library produced. That is deliberate — it makes the suite
/// self-contained, and it exercises the combination a consumer actually hits, which is "I rendered
/// something, now I want to do one more thing to it".
/// </summary>
public class PdfEditorTests
{
    private static async Task<byte[]> PdfAsync(string heading) =>
        await HtmlToPdfConverter.ConvertAsync($"<h1>{heading}</h1><p>Body text.</p>");

    // =============================================================================================
    // Page count
    // =============================================================================================

    [Fact]
    public async Task APdfThisLibraryProducedHasOnePage()
    {
        Assert.Equal(1, PdfEditor.PageCount(await PdfAsync("One")));
    }

    [Fact]
    public async Task PageCountReadsFromAStream()
    {
        await using var source = new MemoryStream(await PdfAsync("One"));

        Assert.Equal(1, await PdfEditor.PageCountAsync(source));
    }

    [Fact]
    public async Task InputThatIsNotAPdfIsRejectedAsOneExceptionType()
    {
        var notAPdf = Encoding.UTF8.GetBytes("This is not a PDF.");

        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => Task.FromResult(PdfEditor.PageCount(notAPdf)));

        Assert.NotNull(ex.InnerException);
    }

    // =============================================================================================
    // Merge
    // =============================================================================================

    [Fact]
    public async Task MergingTwoDocumentsAddsTheirPages()
    {
        var merged = PdfEditor.Merge([await PdfAsync("First"), await PdfAsync("Second")]);

        Assert.Equal(2, PdfEditor.PageCount(merged));
    }

    [Fact]
    public async Task MergingPreservesOrder()
    {
        // Page one of the merge must be page one of the FIRST document. Comparing extracted single
        // pages by byte length is unreliable, so this compares the text each page carries.
        var merged = PdfEditor.Merge([await PdfAsync("Alpha"), await PdfAsync("Beta")]);

        Assert.Contains("Alpha", PdfProbe.ExtractText(PdfEditor.ExtractPages(merged, 1, 1)), StringComparison.Ordinal);
        Assert.Contains("Beta", PdfProbe.ExtractText(PdfEditor.ExtractPages(merged, 2, 1)), StringComparison.Ordinal);
    }

    [Fact]
    public void MergingNothingIsRejectedRatherThanProducingAnEmptyDocument()
    {
        // A zero-page PDF is not a useful artefact and several readers refuse to open one, so this
        // fails loudly instead of handing back something that looks like a document.
        Assert.Throws<ArgumentException>(() => PdfEditor.Merge([]));
    }

    // =============================================================================================
    // Extract pages — the split half
    // =============================================================================================

    [Fact]
    public async Task ExtractingASinglePageYieldsASinglePageDocument()
    {
        var merged = PdfEditor.Merge([await PdfAsync("A"), await PdfAsync("B"), await PdfAsync("C")]);

        Assert.Equal(1, PdfEditor.PageCount(PdfEditor.ExtractPages(merged, 2, 1)));
        Assert.Equal(2, PdfEditor.PageCount(PdfEditor.ExtractPages(merged, 2, 2)));
    }

    [Theory]
    [InlineData(0, 1)]      // pages are 1-based; 0 is not a page
    [InlineData(1, 0)]      // asking for no pages
    [InlineData(4, 1)]      // past the end
    [InlineData(3, 2)]      // starts inside, runs off the end
    public async Task AnExtractionRangeOutsideTheDocumentIsRejected(int firstPage, int count)
    {
        var merged = PdfEditor.Merge([await PdfAsync("A"), await PdfAsync("B"), await PdfAsync("C")]);

        Assert.ThrowsAny<ArgumentException>(() => PdfEditor.ExtractPages(merged, firstPage, count));
    }

    [Fact]
    public async Task TheWholeDocumentIsAValidRange()
    {
        var merged = PdfEditor.Merge([await PdfAsync("A"), await PdfAsync("B"), await PdfAsync("C")]);

        // The boundary: firstPage + count - 1 lands exactly on the last page. Off by one here and
        // extracting an entire document fails, which is the least surprising thing to ask for.
        Assert.Equal(3, PdfEditor.PageCount(PdfEditor.ExtractPages(merged, 1, 3)));
    }

    // =============================================================================================
    // Metadata
    // =============================================================================================

    [Fact]
    public async Task MetadataSurvivesARoundTrip()
    {
        var stamped = PdfEditor.WithMetadata(
            await PdfAsync("Report"),
            new PdfMetadata
            {
                Title = "Quarterly report",
                Author = "Contoso Ltd",
                Subject = "Revenue",
                Keywords = "revenue, quarterly",
            });

        var read = PdfEditor.ReadMetadata(stamped);

        Assert.Equal("Quarterly report", read.Title);
        Assert.Equal("Contoso Ltd", read.Author);
        Assert.Equal("Revenue", read.Subject);
        Assert.Equal("revenue, quarterly", read.Keywords);
    }

    [Fact]
    public async Task StampingMetadataDoesNotDisturbThePages()
    {
        var merged = PdfEditor.Merge([await PdfAsync("A"), await PdfAsync("B")]);

        var stamped = PdfEditor.WithMetadata(merged, new PdfMetadata { Title = "Two pages" });

        Assert.Equal(2, PdfEditor.PageCount(stamped));
    }

    [Fact]
    public async Task MetadataNotSetReadsBackAsNullRatherThanEmpty()
    {
        // Distinguishing "absent" from "present but blank" matters to anything that merges metadata
        // from more than one source, so the absent case must not arrive as "".
        var read = PdfEditor.ReadMetadata(await PdfAsync("Plain"));

        Assert.Null(read.Title);
        Assert.Null(read.Subject);
    }


    // =============================================================================================
    // The Stream and path overloads
    // =============================================================================================

    [Fact]
    public async Task MergeAndExtractWorkThroughStreams()
    {
        await using var merged = new MemoryStream();
        await PdfEditor.MergeAsync(
            [new MemoryStream(await PdfAsync("A")), new MemoryStream(await PdfAsync("B"))],
            merged);

        Assert.Equal(2, await PdfEditor.PageCountAsync(new MemoryStream(merged.ToArray())));

        await using var extracted = new MemoryStream();
        await PdfEditor.ExtractPagesAsync(new MemoryStream(merged.ToArray()), 1, 1, extracted);

        Assert.Equal(1, PdfEditor.PageCount(extracted.ToArray()));
    }

    [Fact]
    public async Task PageCountReadsFromAPath()
    {
        var path = Path.Join(Path.GetTempPath(), $"doctoolkit-{Guid.NewGuid():N}.pdf");
        await File.WriteAllBytesAsync(path, await PdfAsync("One"));

        try
        {
            Assert.Equal(1, await PdfEditor.PageCountAsync(path));
        }
        finally
        {
            File.Delete(path);
        }
    }

    [Fact]
    public async Task AStreamThatCannotSeekIsReadThrough()
    {
        // The same guarantee the rest of the library makes: a request body or a pipe is a valid
        // source. A MemoryStream would hide a rewind, so this one refuses to seek.
        await using var source = new ForwardOnly(await PdfAsync("One"));

        Assert.Equal(1, await PdfEditor.PageCountAsync(source));
    }

    [Fact]
    public void NullArgumentsAreRejectedByName()
    {
        Assert.Equal("pdf", Assert.Throws<ArgumentNullException>(() => PdfEditor.PageCount(null!)).ParamName);
        Assert.Equal("pdfs", Assert.Throws<ArgumentNullException>(() => PdfEditor.Merge(null!)).ParamName);
        Assert.Equal("metadata", Assert.Throws<ArgumentNullException>(
            () => PdfEditor.WithMetadata([1, 2, 3], null!)).ParamName);
    }

    private sealed class ForwardOnly(byte[] data) : Stream
    {
        private readonly MemoryStream _inner = new(data);

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override void Flush() { }
    }
}
