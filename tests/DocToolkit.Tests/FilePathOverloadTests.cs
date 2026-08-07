using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// The file-path overloads, tested as one surface rather than one class at a time.
///
/// Each one reads its input completely, calls the byte[] overload, and only then opens the output.
/// That ordering is not decoration: it is what makes editing a document in place safe, and what
/// stops a failed conversion truncating a file that was already there. Both properties survive
/// exactly as long as nobody "optimises" the implementation into a stream copy, which is why they
/// are pinned here instead of being left as a comment.
///
/// These are NOT Stream overloads, so StreamOverloadTests does not cover them — none of its
/// properties (forward-only sources, not disposing caller streams) apply.
/// </summary>
public class FilePathOverloadTests
{
    private static byte[] Docx() =>
        DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("Dear {{customer}}")));

    private static Dictionary<string, string> Replacements() =>
        new() { ["{{customer}}"] = "Contoso Ltd" };

    /// <summary>
    /// Editing a document in place: the same path as input and output. An implementation that
    /// streamed input to output would truncate the file before it finished reading it.
    /// </summary>
    [Fact]
    public async Task EditingInPlace_Works_WhenInputAndOutputAreTheSameFile()
    {
        using var file = new TempFile();
        await File.WriteAllBytesAsync(file.Path, Docx());

        await DocxEditor.ReplaceTextAsync(file.Path, file.Path, Replacements());

        Assert.Contains("Contoso Ltd", await DocxEditor.ExtractTextAsync(file.Path));
    }

    /// <summary>
    /// A document that cannot be processed must leave an existing output file exactly as it was.
    /// The bytes are computed before the destination is opened, so there is no window in which a
    /// half-written file exists.
    /// </summary>
    [Fact]
    public async Task AFailedConversion_LeavesAnExistingOutputFileUntouched()
    {
        var sentinel = "do not overwrite me"u8.ToArray();

        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, new byte[] { 1, 2, 3, 4 });   // not a .docx
        await File.WriteAllBytesAsync(output.Path, sentinel);

        await Assert.ThrowsAsync<DocumentConversionException>(
            () => DocxEditor.ReplaceTextAsync(input.Path, output.Path, Replacements()));

        Assert.Equal(sentinel, await File.ReadAllBytesAsync(output.Path));
    }

    /// <summary>
    /// The same three properties, for <c>CreateToFileAsync</c>. It has no input file, so the
    /// read-completely half does not apply — but "a failed build must not touch an existing output"
    /// does, and it held for none of them until these tests existed: the plan specified only a
    /// happy-path test and a blank-path test, so this overload would have shipped carrying none of
    /// the guarantees this class exists to pin. Structurally the same escape as an overload missing
    /// from <c>StreamOverloadTests</c>' name lists, one layer out.
    /// </summary>
    [Fact]
    public async Task AFailedBuild_LeavesAnExistingOutputFileUntouched()
    {
        var sentinel = "do not overwrite me"u8.ToArray();

        using var output = new TempFile();
        await File.WriteAllBytesAsync(output.Path, sentinel);

        // A null block fails ValidateBlocks, before any byte of the document is built.
        await Assert.ThrowsAsync<ArgumentException>(
            () => DocxEditor.CreateToFileAsync(new DocxBlock[] { null! }, output.Path));

        Assert.Equal(sentinel, await File.ReadAllBytesAsync(output.Path));
    }

    /// <summary>A successful build overwrites, and the result is a real document.</summary>
    [Fact]
    public async Task CreateToFileAsync_WritesAValidDocument_OverwritingWhatWasThere()
    {
        using var output = new TempFile();
        await File.WriteAllBytesAsync(output.Path, "stale"u8.ToArray());

        await DocxEditor.CreateToFileAsync(
            new[] { DocxBlock.Paragraph("Written to disk.") }, output.Path);

        Assert.Contains("Written to disk.", await DocxEditor.ExtractTextAsync(output.Path));
    }

    /// <summary>
    /// The <c>ParamName</c> assertion is what makes this non-vacuous on Windows. Deleting the
    /// <c>ThrowIfNullOrWhiteSpace</c> guard entirely still throws <see cref="ArgumentException"/>
    /// here, because <c>File.WriteAllBytesAsync(" ", …)</c> rejects the path itself — but with
    /// <c>ParamName = "path"</c>, not <c>"outputPath"</c>. On Linux <c>" "</c> is a legal filename
    /// and the guard is the only thing rejecting it, so without this assertion the test proves
    /// nothing locally and something entirely different in CI.
    /// </summary>
    [Fact]
    public async Task CreateToFileAsync_RejectsABlankPath()
    {
        var ex = await Assert.ThrowsAsync<ArgumentException>(
            () => DocxEditor.CreateToFileAsync(Array.Empty<DocxBlock>(), " "));

        Assert.Equal("outputPath", ex.ParamName);
    }

    [Fact]
    public async Task CreateToFileAsync_AnAlreadyCancelledToken_PreventsAnyOutputBeingWritten()
    {
        using var output = new TempFile();
        using var cts = new CancellationTokenSource();
        await cts.CancelAsync();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => DocxEditor.CreateToFileAsync(
                new[] { DocxBlock.Paragraph("x") }, output.Path, cts.Token));

        Assert.False(File.Exists(output.Path), "a cancelled call wrote an output file");
    }

    /// <summary>A successful conversion does overwrite — the previous contents are gone.</summary>
    [Fact]
    public async Task ASuccessfulConversion_OverwritesAnExistingOutputFile()
    {
        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, Docx());
        await File.WriteAllBytesAsync(output.Path, "stale"u8.ToArray());

        await DocxEditor.ReplaceTextAsync(input.Path, output.Path, Replacements());

        Assert.Contains("Contoso Ltd", await DocxEditor.ExtractTextAsync(output.Path));
    }

    /// <summary>An already-cancelled token stops the work before anything is written.</summary>
    [Fact]
    public async Task AnAlreadyCancelledToken_PreventsAnyOutputBeingWritten()
    {
        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, Docx());

        using var cts = new CancellationTokenSource();
        await cts.CancelAsync();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => DocxEditor.ReplaceTextAsync(input.Path, output.Path, Replacements(), cts.Token));

        Assert.False(File.Exists(output.Path), "a cancelled call wrote an output file");
    }

    /// <summary>
    /// A zero-byte input file (an interrupted download, a pre-created placeholder) is not a blank
    /// path, so it does not fail the path validation the docs lead with — it reaches the byte[]
    /// overload's own empty-content check instead, and surfaces that overload's
    /// <see cref="ArgumentException.ParamName"/> ("docx"), not the file-path parameter the caller
    /// actually passed. This is documented behaviour, not a bug, so it is pinned here rather than
    /// left to be "fixed" into a different exception later.
    /// </summary>
    [Fact]
    public async Task AZeroByteInputFile_ThrowsArgumentException_OnAReadOverload()
    {
        using var input = new TempFile();
        await File.WriteAllBytesAsync(input.Path, Array.Empty<byte>());

        var ex = await Assert.ThrowsAsync<ArgumentException>(
            () => DocxEditor.ExtractTextAsync(input.Path));

        Assert.Equal("docx", ex.ParamName);
    }

    /// <summary>Same pin as above, for a transform rather than a pure read.</summary>
    [Fact]
    public async Task AZeroByteInputFile_ThrowsArgumentException_OnATransform()
    {
        using var input = new TempFile();
        using var output = new TempFile();
        await File.WriteAllBytesAsync(input.Path, Array.Empty<byte>());

        var ex = await Assert.ThrowsAsync<ArgumentException>(
            () => DocxEditor.ReplaceTextAsync(input.Path, output.Path, Replacements()));

        Assert.Equal("docx", ex.ParamName);
    }
}
