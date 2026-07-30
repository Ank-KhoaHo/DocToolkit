using System.Diagnostics.CodeAnalysis;

namespace DocToolkit;

/// <summary>
/// The shared plumbing behind every <see cref="Stream"/> overload: guard the caller's streams,
/// drain a source asynchronously, and emit a finished package asynchronously.
///
/// Two constraints shape all of it.
///
/// <b>The caller owns its streams.</b> Nothing here disposes, closes or seeks a destination, and
/// nothing disposes a source. A destination may be an HTTP response body, which is write-only and
/// forward-only; rewinding it to patch a header is not an option, so a package is assembled in a
/// scratch buffer and then copied out in one direction.
///
/// <b>The packaging layer is not incremental.</b> An OOXML package is a ZIP, and a ZIP's central
/// directory lives at the end of the file, so neither <c>System.IO.Packaging</c> nor ClosedXML can
/// begin work until the whole thing has arrived. The source is therefore drained once, in chunks,
/// with <c>ReadAsync</c> — which is where the <c>async</c> in these overloads is earned and where
/// the <see cref="CancellationToken"/> actually bites. What the overloads remove is not the buffer
/// but the <i>duplication</i>: a <c>byte[]</c>-in/<c>byte[]</c>-out call holds the caller's array,
/// an internal copy of it, and a <c>ToArray()</c> of the result at the same time.
/// </summary>
internal static class StreamPipeline
{
    /// <summary>
    /// The framework's own default for <see cref="Stream.CopyToAsync(Stream)"/>. Large enough that
    /// a multi-megabyte document is not copied a page at a time, small enough to stay off the
    /// large object heap.
    /// </summary>
    private const int CopyBufferSize = 81920;

    /// <summary>
    /// Rejects a source that is missing or cannot be read, by name and in a sentence.
    ///
    /// Without this a caller who passes a write-only half of a pipe gets a
    /// <see cref="NotSupportedException"/> raised several libraries down, naming neither the
    /// argument nor the problem.
    /// </summary>
    public static void RequireReadable([NotNull] Stream? source, string paramName)
    {
        ArgumentNullException.ThrowIfNull(source, paramName);

        if (!source.CanRead)
        {
            throw new ArgumentException(
                "Source stream is not readable. Pass a stream opened with read access.", paramName);
        }
    }

    /// <summary>Rejects a destination that is missing or cannot be written, by name.</summary>
    public static void RequireWritable([NotNull] Stream? destination, string paramName)
    {
        ArgumentNullException.ThrowIfNull(destination, paramName);

        if (!destination.CanWrite)
        {
            throw new ArgumentException(
                "Destination stream is not writable. Pass a stream opened with write access.", paramName);
        }
    }

    /// <summary>
    /// Reads <paramref name="source"/> to its end into a seekable scratch buffer, positioned at 0.
    ///
    /// The read is chunked and awaited, so a source that is a socket or an HTTP request body does
    /// not pin the calling thread, and <paramref name="ct"/> is observed between chunks rather than
    /// only before the first one. The caller's stream is never sought and never disposed.
    /// </summary>
    /// <exception cref="ArgumentException"><paramref name="source"/> held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The source could not be read.</exception>
    public static async Task<MemoryStream> DrainAsync(
        Stream source,
        string emptyMessage,
        string paramName,
        string failureMessage,
        CancellationToken ct)
    {
        var scratch = NewScratch(source);

        try
        {
            await source.CopyToAsync(scratch, CopyBufferSize, ct).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            scratch.Dispose();
            throw;
        }
        catch (Exception ex)
        {
            scratch.Dispose();
            throw new DocumentConversionException(failureMessage, ex);
        }

        // A zero-byte "document" is what a truncated upload looks like. The byte[] overloads
        // already reject it by name rather than letting the package layer report a corrupt ZIP.
        if (scratch.Length == 0)
        {
            scratch.Dispose();
            throw new ArgumentException(emptyMessage, paramName);
        }

        scratch.Position = 0;
        return scratch;
    }

    /// <summary>
    /// Copies <paramref name="scratch"/> from its start to <paramref name="destination"/>.
    ///
    /// The destination is written with <c>WriteAsync</c> and never sought, never read and never
    /// disposed, so an HTTP response body is a valid destination.
    /// </summary>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The destination could not be written.</exception>
    public static async Task EmitAsync(
        MemoryStream scratch, Stream destination, string failureMessage, CancellationToken ct)
    {
        scratch.Position = 0;

        try
        {
            await scratch.CopyToAsync(destination, CopyBufferSize, ct).ConfigureAwait(false);
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException(failureMessage, ex);
        }
    }

    /// <summary>
    /// A scratch buffer sized to the source where that is knowable.
    ///
    /// A <see cref="MemoryStream"/> that has to grow reallocates and copies its whole contents on
    /// every doubling, so a 50 MB document arriving into a default-sized buffer is copied a dozen
    /// times on the way in. A forward-only source cannot say how long it is, and that is fine — it
    /// just gets the growing buffer.
    /// </summary>
    private static MemoryStream NewScratch(Stream source)
    {
        if (!source.CanSeek) return new MemoryStream();

        var remaining = source.Length - source.Position;
        return remaining is > 0 and <= int.MaxValue ? new MemoryStream((int)remaining) : new MemoryStream();
    }
}
