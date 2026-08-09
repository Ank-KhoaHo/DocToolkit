using Xunit;
using Xunit.Abstractions;

namespace DocToolkit.Tests;

/// <summary>
/// Pins the one performance claim this repository makes about the <see cref="Stream"/> overloads:
/// <b>they are not a memory optimisation.</b>
///
/// `CLAUDE.md` records the measurement — the same edit costs 238 MB through `SetCellAsync` against
/// 233 MB through `SetCell`, because `DrainAsync` buffers the whole source anyway — and warns
/// against "drifting into an efficiency claim the numbers do not support". That warning exists
/// because the drift already happened once. A sentence in a file cannot stop it happening again;
/// this can.
///
/// <b>Allocation, not elapsed time.</b> Wall-clock assertions on a shared CI runner have flapped
/// twice in this suite already and needed best-of-N defences. Allocated bytes are close to
/// deterministic for the same input, which is what makes this gateable where a timing assertion
/// would not be.
///
/// A failure here does not necessarily mean the code is wrong — it means the <i>documentation</i> is
/// now wrong, whichever direction the number moved.
/// </summary>
public class StreamAllocationParityTests
{
    private readonly ITestOutputHelper _output;

    public StreamAllocationParityTests(ITestOutputHelper output) => _output = output;

    /// <summary>
    /// Parity band. Measured 2026-08-09 at a ratio of <b>1.000</b>; this allows a 30% drift either
    /// way before failing, which is far tighter than the claim being protected ("not cheaper") and
    /// far looser than GC noise.
    /// </summary>
    private const double LowerBound = 0.70;
    private const double UpperBound = 1.40;

    private static byte[] Workbook(int rows) => WorkbookEditor.Create("Sales",
        Enumerable.Range(1, rows).Select(i => new object?[] { "Region " + i, i }));

    /// <summary>
    /// Runs the action once untimed — JIT and first-touch statics are not the subject — then
    /// collects and measures the second run.
    /// </summary>
    private static long AllocatedBy(Action action)
    {
        action();
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        long before = GC.GetAllocatedBytesForCurrentThread();
        action();
        return GC.GetAllocatedBytesForCurrentThread() - before;
    }

    [Fact]
    public void TheStreamOverloadAllocatesTheSameAsTheByteArrayOverload()
    {
        byte[] xlsx = Workbook(2000);

        long viaArray = AllocatedBy(() => WorkbookEditor.SetCell(xlsx, "Sales", "B2", 1500));
        long viaStream = AllocatedBy(() =>
        {
            using var source = new MemoryStream(xlsx);
            using var destination = new MemoryStream();
            WorkbookEditor.SetCellAsync(source, "Sales", "B2", 1500, destination)
                .GetAwaiter().GetResult();
        });

        double ratio = (double)viaStream / viaArray;
        _output.WriteLine(
            $"byte[] {viaArray / 1048576.0:0.0} MB, Stream {viaStream / 1048576.0:0.0} MB, "
            + $"ratio {ratio:0.000} (band {LowerBound}-{UpperBound})");

        Assert.True(
            ratio > LowerBound && ratio < UpperBound,
            $"The Stream overload now allocates {ratio:0.000}x the byte[] overload. README.md and "
            + "CLAUDE.md both state the Stream overloads exist for forward-only, non-seekable "
            + "sources and are NOT cheaper. Whichever way this moved, that documentation is now "
            + "wrong - update it, or explain why the change is not what it appears to be.");
    }
}
