using System.Diagnostics;
using DocToolkit;

Console.WriteLine("Large-file streaming");
Console.WriteLine("====================");

// WHY THE STREAM OVERLOADS EXIST, and what they do NOT do.
//
// They are not a memory optimisation. Measured: the same edit costs about the same either way,
// because reading a package requires the whole thing to be present - a .docx or .xlsx is a ZIP,
// and a ZIP's central directory is at the END. You cannot process the first entry until you have
// seen the last byte.
//
// What they buy is that `source` may be FORWARD-ONLY and non-seekable: an HTTP request body, a
// network stream, a pipe. Those cannot be handed to a ZIP reader directly, and the alternative -
// ReadAllBytes into a byte[] first - is exactly what the Stream overload does for you, correctly,
// including honouring your CancellationToken while it does it.
//
// This sample proves that with a stream that REFUSES to seek, so a rewind would throw rather than
// quietly work on a MemoryStream and mislead you.

// --- A workbook big enough that the numbers mean something -----------------------------------

byte[] source = WorkbookEditor.Create("Sales",
    Enumerable.Range(1, 50_000).Select(i => new object?[] { $"Region {i}", i, i * 1.5 }));

Console.WriteLine($"\nWorkbook: {source.Length / 1024.0 / 1024:0.0} MB, 50,000 rows");

// --- Edit it, reading from something that cannot seek ----------------------------------------

var sw = Stopwatch.StartNew();

// Scoped so the file handle is closed before it is read back below. DocToolkit never closes a
// stream you hand it - it is yours - which is what lets `output` be a response body you go on
// writing to, and equally what makes closing it your job here.
using (var input = new ForwardOnlyStream(source))
using (var output = File.Create("edited.xlsx"))
{
    await WorkbookEditor.SetCellAsync(input, "Sales", "B2", 999_999, output);
}

sw.Stop();

Console.WriteLine($"Edited in {sw.ElapsedMilliseconds:N0} ms, straight from a non-seekable source");
Console.WriteLine($"Wrote     : edited.xlsx, {new FileInfo("edited.xlsx").Length / 1024.0 / 1024:0.0} MB");

// The destination was written but never read back or sought by DocToolkit either, so it could
// equally have been a response body. Reading it here is this sample verifying its own output.
Console.WriteLine($"Cell B2   : {WorkbookEditor.ReadCell(File.ReadAllBytes("edited.xlsx"), "Sales", "B2")}");

// --- The peak-memory caveat, stated rather than implied --------------------------------------

Console.WriteLine(
    $"\nPeak managed memory: {GC.GetTotalAllocatedBytes() / 1024.0 / 1024:N0} MB allocated in total.");
Console.WriteLine(
    "Note this is NOT lower than the byte[] overload would be - see the comment at the top of\n" +
    "this file. Memory is dominated by the OOXML object model, not by how the bytes arrive.");

File.Delete("edited.xlsx");
Console.WriteLine("\nDone.");

// --- A stream that behaves like a socket, not a file ------------------------------------------

/// <summary>Forward-only: no Length, no Position, no Seek. Exactly what a request body gives you.</summary>
sealed class ForwardOnlyStream(byte[] data) : Stream
{
    private readonly MemoryStream _inner = new(data);

    public override bool CanRead => true;
    public override bool CanSeek => false;
    public override bool CanWrite => false;

    public override long Length => throw new NotSupportedException("A forward-only source has no length.");
    public override long Position
    {
        get => throw new NotSupportedException("A forward-only source has no position.");
        set => throw new NotSupportedException("A forward-only source cannot be positioned.");
    }

    public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
    public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException("Cannot seek.");
    public override void SetLength(long value) => throw new NotSupportedException();
    public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    public override void Flush() { }
}
