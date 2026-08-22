using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace DocToolkit;

/// <summary>
/// Reads a file for a public file-path overload, raising argument failures against the parameter
/// the <b>caller</b> passed.
/// </summary>
/// <remarks>
/// <b>Without this, a zero-byte file reached the <c>byte[]</c> implementation underneath and that
/// overload raised the exception naming ITS OWN parameter</b> - so
/// <c>DocxEditor.ExtractTextAsync(path)</c> told a caller to go and check an argument called
/// <c>docx</c>, which its signature does not have. Nineteen overloads across six types did this,
/// found by walking the public surface rather than by reading, and the review that filed it had
/// spotted two.
///
/// <para><b>The shipped documentation already said otherwise, which is what settled it.</b>
/// Thirteen of those overloads carry an
/// <c>&lt;exception cref="ArgumentException"&gt;</c> tag attributing the empty-file case to
/// <c>path</c> or <c>inputPath</c> - so the behaviour disagreed with the XML docs that render on
/// the API site, rather than implementing them. A test had pinned the old behaviour as
/// "documented behaviour, not a bug"; the documentation says the opposite, and the comment on
/// <c>PdfEditor.Open</c> had already called an exception naming the wrong parameter "its own
/// defect".</para>
///
/// <para>The blank-path guard here is deliberate belt-and-braces: every public overload already
/// checks its own path first, and this makes the helper correct for any that forgets.</para>
/// </remarks>
internal static class FilePipeline
{
    internal static byte[] Read(string path, string paramName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path, paramName);
        return NonEmpty(File.ReadAllBytes(path), path, paramName);
    }

    internal static async Task<byte[]> ReadAsync(string path, string paramName, CancellationToken ct)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path, paramName);
        return NonEmpty(await File.ReadAllBytesAsync(path, ct).ConfigureAwait(false), path, paramName);
    }

    private static byte[] NonEmpty(byte[] bytes, string path, string paramName)
    {
        // Named after the FILE rather than after the content, because the caller handed over a
        // path and never saw any bytes. "DOCX content was empty" describes something they did not
        // pass.
        if (bytes.Length == 0)
            throw new ArgumentException($"The file at '{path}' is empty.", paramName);

        return bytes;
    }
}
