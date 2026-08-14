namespace DocToolkit.Tests;

/// <summary>
/// Stream doubles for the <see cref="Stream"/> overloads.
///
/// The point of those overloads is that a caller's stream may be a file, a socket or an HTTP
/// message body — things that are forward-only, that cannot be sought, that are write-only, and
/// that the caller, not DocToolkit, owns. A <c>MemoryStream</c> is none of those, so it cannot
/// tell a genuine streaming implementation apart from a <c>byte[]</c> round trip wearing a
/// <c>Stream</c> parameter. These can.
/// </summary>
internal static class StreamDoubles
{
    /// <summary>A readable, seekable stream over <paramref name="bytes"/>, as a file would be.</summary>
    public static MemoryStream Seekable(byte[] bytes)
    {
        var ms = new MemoryStream();
        ms.Write(bytes, 0, bytes.Length);
        ms.Position = 0;
        return ms;
    }
}

/// <summary>
/// Records what was done to a stream and — deliberately — refuses to close.
///
/// <see cref="IsDisposed"/> flips when someone calls <c>Dispose</c>/<c>DisposeAsync</c>, but the
/// underlying buffer stays usable so the test can still assert on the bytes afterwards. That is
/// what makes "the caller's stream is left open" checkable: a stream that really closed would fail
/// the later assertions for the wrong reason.
/// </summary>
internal sealed class TrackingStream : Stream
{
    private readonly Stream _inner;

    public TrackingStream(Stream inner) => _inner = inner;

    public int SyncReads { get; private set; }

    public int AsyncReads { get; private set; }

    public int SyncWrites { get; private set; }

    public int AsyncWrites { get; private set; }

    public int Seeks { get; private set; }

    public bool IsDisposed { get; private set; }

    public long BytesWritten { get; private set; }

    public override bool CanRead => _inner.CanRead;

    public override bool CanSeek => _inner.CanSeek;

    public override bool CanWrite => _inner.CanWrite;

    public override long Length => _inner.Length;

    public override long Position
    {
        get => _inner.Position;
        set { Seeks++; _inner.Position = value; }
    }

    public override void Flush() => _inner.Flush();

    public override Task FlushAsync(CancellationToken ct) => _inner.FlushAsync(ct);

    public override int Read(byte[] buffer, int offset, int count)
    {
        SyncReads++;
        return _inner.Read(buffer, offset, count);
    }

    public override int Read(Span<byte> buffer)
    {
        SyncReads++;
        return _inner.Read(buffer);
    }

    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken ct)
    {
        AsyncReads++;
        return _inner.ReadAsync(buffer, offset, count, ct);
    }

    public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken ct = default)
    {
        AsyncReads++;
        return _inner.ReadAsync(buffer, ct);
    }

    public override long Seek(long offset, SeekOrigin origin)
    {
        Seeks++;
        return _inner.Seek(offset, origin);
    }

    public override void SetLength(long value) => _inner.SetLength(value);

    public override void Write(byte[] buffer, int offset, int count)
    {
        SyncWrites++;
        BytesWritten += count;
        _inner.Write(buffer, offset, count);
    }

    public override void Write(ReadOnlySpan<byte> buffer)
    {
        SyncWrites++;
        BytesWritten += buffer.Length;
        _inner.Write(buffer);
    }

    public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken ct)
    {
        AsyncWrites++;
        BytesWritten += count;
        return _inner.WriteAsync(buffer, offset, count, ct);
    }

    public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken ct = default)
    {
        AsyncWrites++;
        BytesWritten += buffer.Length;
        return _inner.WriteAsync(buffer, ct);
    }

    protected override void Dispose(bool disposing) => IsDisposed = true;

    public override ValueTask DisposeAsync()
    {
        IsDisposed = true;
        return ValueTask.CompletedTask;
    }
}

/// <summary>
/// A readable, <b>non-seekable</b> stream: an HTTP request body or a socket. Anything that reaches
/// for <c>Length</c>, <c>Position</c> or <c>Seek</c> gets a <see cref="NotSupportedException"/>.
/// </summary>
internal sealed class ForwardOnlySource : Stream
{
    private readonly Stream _inner;

    public ForwardOnlySource(byte[] bytes) => _inner = new MemoryStream(bytes, writable: false);

    public bool IsDisposed { get; private set; }

    public override bool CanRead => true;

    public override bool CanSeek => false;

    public override bool CanWrite => false;

    public override long Length => throw new NotSupportedException("Forward-only stream.");

    public override long Position
    {
        get => throw new NotSupportedException("Forward-only stream.");
        set => throw new NotSupportedException("Forward-only stream.");
    }

    public override void Flush() { }

    public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

    public override int Read(Span<byte> buffer) => _inner.Read(buffer);

    public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken ct = default)
        => _inner.ReadAsync(buffer, ct);

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    protected override void Dispose(bool disposing) => IsDisposed = true;
}

/// <summary>
/// A write-only, non-seekable, non-readable sink: an HTTP response body. An implementation that
/// tries to rewind the destination, or to read back what it just wrote, fails here.
/// </summary>
internal sealed class ForwardOnlySink : Stream
{
    private readonly MemoryStream _written = new();

    public bool IsDisposed { get; private set; }

    public int Writes { get; private set; }

    public byte[] ToArray() => _written.ToArray();

    public override bool CanRead => false;

    public override bool CanSeek => false;

    public override bool CanWrite => true;

    public override long Length => throw new NotSupportedException("Forward-only stream.");

    public override long Position
    {
        get => throw new NotSupportedException("Forward-only stream.");
        set => throw new NotSupportedException("Forward-only stream.");
    }

    public override void Flush() { }

    public override Task FlushAsync(CancellationToken ct) => Task.CompletedTask;

    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count)
    {
        Writes++;
        _written.Write(buffer, offset, count);
    }

    public override void Write(ReadOnlySpan<byte> buffer)
    {
        Writes++;
        _written.Write(buffer);
    }

    public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken ct)
    {
        Writes++;
        return _written.WriteAsync(buffer, offset, count, ct);
    }

    public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken ct = default)
    {
        Writes++;
        return _written.WriteAsync(buffer, ct);
    }

    protected override void Dispose(bool disposing) => IsDisposed = true;
}

/// <summary>A stream that claims it cannot be read — the wrong end of a pipe handed in by mistake.</summary>
internal sealed class NonReadableStream : Stream
{
    public override bool CanRead => false;

    public override bool CanSeek => false;

    public override bool CanWrite => true;

    public override long Length => 0;

    public override long Position { get => 0; set { } }

    public override void Flush() { }

    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) { }

    public override void Write(byte[] buffer, int offset, int count) { }
}

/// <summary>A stream that claims it cannot be written — a file opened for reading, say.</summary>
internal sealed class NonWritableStream : Stream
{
    public override bool CanRead => true;

    public override bool CanSeek => false;

    public override bool CanWrite => false;

    public override long Length => 0;

    public override long Position { get => 0; set { } }

    public override void Flush() { }

    public override int Read(byte[] buffer, int offset, int count) => 0;

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) { }

    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
}

/// <summary>
/// Cancels the supplied token from <i>inside</i> the first read and then keeps serving bytes.
///
/// This is what separates "the token is checked on the way in" from "the token is honoured while
/// the source is being consumed": the guard at the top of the method has already passed by the
/// time this fires, so only an implementation that keeps observing the token during the read can
/// throw.
/// </summary>
internal sealed class CancelsOnFirstReadSource : Stream
{
    private readonly MemoryStream _inner;
    private readonly CancellationTokenSource _cts;

    public CancelsOnFirstReadSource(byte[] bytes, CancellationTokenSource cts)
    {
        _inner = new MemoryStream(bytes, writable: false);
        _cts = cts;
    }

    public override bool CanRead => true;

    public override bool CanSeek => false;

    public override bool CanWrite => false;

    public override long Length => throw new NotSupportedException();

    public override long Position
    {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override void Flush() { }

    public override int Read(byte[] buffer, int offset, int count)
    {
        _cts.Cancel();
        return _inner.Read(buffer, offset, count);
    }

    public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken ct = default)
    {
        _cts.Cancel();
        return _inner.ReadAsync(buffer, ct);
    }

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
}

/// <summary>
/// Cancels the caller's token on its first write, so cancellation lands <b>during</b> the copy
/// rather than before it.
///
/// The mirror of <see cref="CancelsOnFirstReadSource"/>, and it exists for the same reason that
/// one does: <c>StreamPipeline.EmitAsync</c>'s <c>catch (OperationCanceledException)</c> rethrow
/// was executed by <b>no test at all</b> — reported as <c>NoCoverage</c> when B14 widened the
/// mutation scope. An already-cancelled token never reaches it, because the guard at the top of
/// each overload refuses first; only a token cancelled mid-write does.
/// </summary>
internal sealed class CancelsOnFirstWriteSink : Stream
{
    private readonly CancellationTokenSource _cts;

    public CancelsOnFirstWriteSink(CancellationTokenSource cts) => _cts = cts;

    public override bool CanRead => false;

    public override bool CanSeek => false;

    public override bool CanWrite => true;

    public override long Length => throw new NotSupportedException();

    public override long Position
    {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override void Flush() { }

    public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

    public override void Write(byte[] buffer, int offset, int count)
    {
        _cts.Cancel();
        _cts.Token.ThrowIfCancellationRequested();
    }

    public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken ct = default)
    {
        _cts.Cancel();
        ct.ThrowIfCancellationRequested();
        return ValueTask.CompletedTask;
    }

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();
}

/// <summary>
/// Fails mid-transfer with an <see cref="IOException"/> — a socket dying part-way, not a
/// misused API.
///
/// It exists because <c>StreamPipeline</c>'s generic <c>catch (Exception)</c> arms, the ones that
/// wrap a failure in <see cref="DocumentConversionException"/>, were reached by <b>no test at
/// all</b> — four mutants there came back <c>NoCoverage</c> when B14 widened the mutation scope.
/// Every existing double refuses an operation up front (<c>CanRead</c> false, <c>Seek</c> throws),
/// which is a different path: those are caught by the <c>RequireReadable</c>/<c>RequireWritable</c>
/// guards before any transfer starts.
///
/// <paramref name="failAfter"/> bytes are moved successfully first, so the failure lands in the
/// middle of a copy rather than on its first call — otherwise a pipeline that validated eagerly and
/// never actually transferred would pass.
/// </summary>
internal sealed class FailsPartWayStream : Stream
{
    private readonly MemoryStream _inner;
    private readonly int _failAfter;
    private int _moved;

    public FailsPartWayStream(byte[]? bytes, int failAfter)
    {
        _inner = bytes is null ? new MemoryStream() : new MemoryStream(bytes, writable: false);
        _failAfter = failAfter;
    }

    public override bool CanRead => true;

    public override bool CanSeek => false;

    public override bool CanWrite => true;

    public override long Length => throw new NotSupportedException();

    public override long Position
    {
        get => throw new NotSupportedException();
        set => throw new NotSupportedException();
    }

    public override void Flush() { }

    public override int Read(byte[] buffer, int offset, int count)
    {
        Advance(count);
        return _inner.Read(buffer, offset, count);
    }

    public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken ct = default)
    {
        Advance(buffer.Length);
        return _inner.ReadAsync(buffer, ct);
    }

    public override void Write(byte[] buffer, int offset, int count) => Advance(count);

    public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken ct = default)
    {
        Advance(buffer.Length);
        return ValueTask.CompletedTask;
    }

    private void Advance(int count)
    {
        _moved += count;
        if (_moved > _failAfter)
            throw new IOException("the underlying connection was reset");
    }

    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

    public override void SetLength(long value) => throw new NotSupportedException();
}
