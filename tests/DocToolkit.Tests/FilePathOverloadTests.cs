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
}
