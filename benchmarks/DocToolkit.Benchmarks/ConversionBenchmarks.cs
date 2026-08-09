using BenchmarkDotNet.Attributes;
using DocToolkit;

namespace DocToolkit.Benchmarks;

/// <summary>
/// A baseline for the conversions consumers actually run, so a future change has something to be
/// compared against.
///
/// <b>Nothing gates on these numbers.</b> Elapsed time on a shared runner is noisy enough that two
/// wall-clock assertions in this repository's test suite have already flapped and needed best-of-N
/// defences; a CI gate on benchmark timings would be the same mistake at greater cost. This exists
/// to be run deliberately - locally, or via the Benchmarks workflow - and read by a person.
///
/// The one performance claim that IS gated is allocation parity between the byte[] and Stream
/// overloads, and it lives in the ordinary test suite as
/// <c>StreamAllocationParityTests</c>, because allocated bytes are stable enough to assert where
/// time is not.
/// </summary>
[MemoryDiagnoser]
public class ConversionBenchmarks
{
    private const string Html =
        "<h1>Quarterly Report</h1><p>Revenue was up <strong>12%</strong>.</p>"
        + "<table><tr><th>Region</th><th>Total</th></tr><tr><td>North</td><td>1200</td></tr></table>";

    private byte[] _docx = [];
    private byte[] _xlsx = [];

    [GlobalSetup]
    public void Setup()
    {
        _docx = DocxEditor.Create(new[]
        {
            DocxBlock.Heading("Quarterly Report", 1),
            DocxBlock.Paragraph("Revenue was up 12%."),
        });

        _xlsx = WorkbookEditor.Create("Sales",
            Enumerable.Range(1, 5_000).Select(i => new object?[] { "Region " + i, i }));
    }

    [Benchmark] public Task<byte[]> HtmlToDocx() => HtmlToDocxConverter.ConvertAsync(Html);

    [Benchmark] public Task<byte[]> HtmlToPdf() => HtmlToPdfConverter.ConvertAsync(Html);

    [Benchmark] public byte[] DocxToPdf() => DocxToPdfConverter.Convert(_docx);

    [Benchmark] public string DocxToMarkdown() => DocxToMarkdownConverter.Convert(_docx);

    /// <summary>5,000 rows: large enough that the OOXML object model dominates, as it does in production.</summary>
    [Benchmark] public IReadOnlyList<IReadOnlyList<string>> ReadSheet()
        => WorkbookEditor.ReadSheet(_xlsx, "Sales");

    [Benchmark] public byte[] SetCell() => WorkbookEditor.SetCell(_xlsx, "Sales", "B2", 1500);
}
