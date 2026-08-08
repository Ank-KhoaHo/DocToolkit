using Xunit;
using Xunit.Abstractions;

namespace DocToolkit.Tests;

/// <summary>
/// The public API is documented as "stateless and safe to call concurrently" — in `README.md`, in
/// the package README and in `CLAUDE.md`. Nothing tested it.
///
/// <para><b>What these tests can and cannot prove.</b> A passing concurrency test does not prove
/// thread-safety; it fails to disprove it. Interleavings are the scheduler's to choose, and the one
/// that breaks may never be picked here. That is the same honesty this repository already applies
/// to `OpenXmlValidator` — it proves a package parses, never that it is correct — and it is why
/// these tests are worth having anyway: the claim currently rests on nothing at all, and a race in
/// the dependency stack is the kind of defect that reaches production precisely because no
/// single-threaded test can see it.</para>
///
/// <para><b>Why the risk is in the dependencies, not in this code.</b> The library holds no mutable
/// static state. Surveyed 2026-08-08, the statics are: a readonly `int[]` of heading sizes, a
/// readonly `char[]` of forbidden sheet-name characters, a span-returning property over a literal,
/// `GuardedResourceLoader`'s `HttpClient` (fixed at construction, every per-request value on the
/// request — see `CLAUDE.md`), and `OfflineResourceLoader.Instance`, which has no fields and
/// returns constants. So a race here would come from ClosedXML, DocumentFormat.OpenXml,
/// HtmlToOpenXml or OfficeIMO — checking our own API and not our dependencies is the exact mistake
/// that let ShapeCrawler in.</para>
///
/// <para><b>No test here asserts on time.</b> Every flake this repository has had — three of them —
/// was a wall-clock assertion. These assert only that concurrent results equal the single-threaded
/// result, so a slow or loaded machine cannot fail them; only a genuine race can.</para>
/// </summary>
public class ConcurrencyGuardTests
{
    private readonly ITestOutputHelper _output;

    public ConcurrencyGuardTests(ITestOutputHelper output) => _output = output;

    /// <summary>
    /// Enough concurrency to interleave without making the suite slow. Not a stress level — the
    /// point is overlapping execution, and past four or so threads more parallelism mostly queues.
    /// </summary>
    private const int Concurrency = 8;

    /// <summary>
    /// Runs <paramref name="operation"/> once per worker, <b>each with a different input</b>, all
    /// released together, and asserts every worker got the answer for its own input.
    ///
    /// <para><b>The distinct inputs are the whole point, and that is measured rather than
    /// assumed.</b> An earlier version of this file ran the same input on every worker. Injecting a
    /// real race into <c>WorkbookEditor.CreateCore</c> — a static field written, yielded on, then
    /// read back — and running it both ways, 2026-08-08:</para>
    ///
    /// <list type="bullet">
    /// <item><description>worker <c>i</c> gets an input only it uses → <b>test fails</b>, naming
    /// the worker that received another's answer</description></item>
    /// <item><description>every worker gets byte-for-byte identical input → <b>test passes</b>,
    /// because the two calls swap indistinguishable values</description></item>
    /// </list>
    ///
    /// <para>So the same-input version would have been a test that cannot fail. Do not "simplify"
    /// these back to a shared fixture.</para>
    ///
    /// <para>Results are compared, never timings. An exception from any worker propagates through
    /// <see cref="Task.WhenAll(Task{TResult}[])"/> and fails the test, which is the other shape a
    /// race takes.</para>
    /// </summary>
    private static async Task EachWorkerGetsItsOwnResultAsync<T>(Func<int, Task<T>> operation)
    {
        // Every worker's correct answer, established with nothing else running.
        var expected = new T[Concurrency];
        for (var i = 0; i < Concurrency; i++)
            expected[i] = await operation(i);

        var started = new TaskCompletionSource();

        // The gate releases them together rather than letting them trickle out as each is
        // scheduled - without it the first often finishes before the last begins, and nothing
        // actually overlaps.
        var workers = Enumerable.Range(0, Concurrency)
            .Select(async i =>
            {
                await started.Task;
                return (Worker: i, Result: await operation(i));
            })
            .ToArray();

        started.SetResult();
        var results = await Task.WhenAll(workers);

        foreach (var (worker, result) in results)
            Assert.True(
                EqualityComparer<T>.Default.Equals(expected[worker], result),
                $"worker {worker} got another worker's answer under concurrency. " +
                $"expected '{expected[worker]}' but got '{result}'");
    }

    private static Task EachWorkerGetsItsOwnResultAsync<T>(Func<int, T> operation)
        => EachWorkerGetsItsOwnResultAsync(i => Task.Run(() => operation(i)));

    // =============================================================================================
    // HTML conversion - HtmlToOpenXml, and the OfflineResourceLoader singleton shared by every call
    // =============================================================================================

    /// <summary>Markup unique to one worker, so another worker's result is recognisably wrong.</summary>
    private static string HtmlFor(int worker) =>
        $"<h1>Report {worker}</h1><p>Revenue rose {worker} percent.</p>" +
        $"<table><tr><td>Region{worker}</td><td>{worker * 100}</td></tr></table>";

    [Fact]
    public Task HtmlToDocx_ConvertsCorrectlyUnderConcurrency()
        // Compared by extracted text rather than bytes: a .docx is a zip, and two conversions a
        // moment apart legitimately differ byte-for-byte.
        => EachWorkerGetsItsOwnResultAsync(async worker =>
            DocxEditor.ExtractText(await HtmlToDocxConverter.ConvertAsync(HtmlFor(worker))));

    [Fact]
    public Task HtmlToPdf_ConvertsCorrectlyUnderConcurrency()
        // The heaviest path: HTML through DOCX through OfficeIMO's PDF writer, the most
        // reflection- and font-dependent code in the graph.
        => EachWorkerGetsItsOwnResultAsync(async worker =>
            PdfProbe.ExtractText(await HtmlToPdfConverter.ConvertAsync(HtmlFor(worker))));

    [Fact]
    public async Task DocxToPdf_ConvertsCorrectlyUnderConcurrency()
    {
        // One distinct source document per worker, built up front so the conversion under test is
        // the only thing running concurrently.
        var sources = new byte[Concurrency][];
        for (var i = 0; i < Concurrency; i++)
            sources[i] = await HtmlToDocxConverter.ConvertAsync(HtmlFor(i));

        await EachWorkerGetsItsOwnResultAsync(
            worker => PdfProbe.ExtractText(DocxToPdfConverter.Convert(sources[worker])));
    }

    // =============================================================================================
    // DOCX - DocumentFormat.OpenXml
    // =============================================================================================

    [Fact]
    public Task DocxEditor_CreatesCorrectlyUnderConcurrency()
        => EachWorkerGetsItsOwnResultAsync(worker => DocxEditor.ExtractText(DocxEditor.Create(new[]
        {
            DocxBlock.Heading($"Report {worker}", 1),
            DocxBlock.Paragraph($"Revenue rose {worker} percent."),
            DocxBlock.Table(
                new[] { "Region", "Revenue" },
                new[] { new object?[] { $"Region{worker}", worker * 100 } }),
        })));

    [Fact]
    public Task DocxEditor_ReplacesTextCorrectlyUnderConcurrency()
        // Exercises RunTextSplicer, which the DOCX and PPTX editors share - so a static in it
        // would corrupt both.
        => EachWorkerGetsItsOwnResultAsync(worker =>
        {
            var docx = DocxEditor.Create(new[] { DocxBlock.Paragraph("Hello {{who}}, from {{where}}.") });
            return DocxEditor.ExtractText(DocxEditor.ReplaceText(docx, new Dictionary<string, string>
            {
                ["{{who}}"] = $"World{worker}",
                ["{{where}}"] = $"Region{worker}",
            }));
        });

    // =============================================================================================
    // XLSX - ClosedXML, which carries the most internal caching of the four
    // =============================================================================================

    [Fact]
    public Task WorkbookEditor_CreatesAndReadsCorrectlyUnderConcurrency()
        => EachWorkerGetsItsOwnResultAsync(worker =>
        {
            var xlsx = WorkbookEditor.Create(new[]
            {
                XlsxSheet.Named($"Sales{worker}", new[]
                {
                    new object?[] { "Region", "Q1", "Q2" },
                    new object?[] { $"Region{worker}", worker * 10, worker * 20 },
                }),
                XlsxSheet.Named("Summary", new[]
                {
                    new object?[] { "Total", XlsxFormula.From($"SUM(Sales{worker}!B2:C2)") },
                }),
            });

            return string.Join(
                "|",
                string.Join(",", WorkbookEditor.SheetNames(xlsx)),
                WorkbookEditor.ReadCell(xlsx, $"Sales{worker}", "A2"),
                WorkbookEditor.ReadCell(xlsx, "Summary", "B1"));
        });

    [Fact]
    public Task WorkbookEditor_AppendsCorrectlyUnderConcurrency()
        => EachWorkerGetsItsOwnResultAsync(worker =>
        {
            var xlsx = WorkbookEditor.Create("Log", new[] { new object?[] { $"start{worker}" } });
            var appended = WorkbookEditor.AppendRows(xlsx, "Log", new[]
            {
                new object?[] { $"a{worker}" },
                new object?[] { $"b{worker}" },
            });

            return string.Join("|", WorkbookEditor.ReadSheet(appended, "Log").Select(r => string.Join(",", r)));
        });

    // =============================================================================================
    // PPTX
    // =============================================================================================

    [Fact]
    public Task PresentationEditor_CreatesAndEditsCorrectlyUnderConcurrency()
        => EachWorkerGetsItsOwnResultAsync(worker =>
        {
            var pptx = PresentationEditor.Create(new[]
            {
                PptxSlide.Titled($"Results {worker}", $"Revenue up {worker}"),
                PptxSlide.Titled("Outlook {{when}}", "Hiring"),
            });

            var edited = PresentationEditor.ReplaceText(
                pptx, new Dictionary<string, string> { ["{{when}}"] = $"20{worker}7" });

            return $"{PresentationEditor.SlideCount(edited)}|{string.Join("~", PresentationEditor.ExtractText(edited))}";
        });

    // =============================================================================================
    // The mixed workload: all six static classes at once, which is how a real service uses them
    // =============================================================================================

    /// <summary>
    /// The tests above run one capability at a time. This runs all of them together, which is the
    /// arrangement a web application actually produces and the only one that can surface state
    /// shared <i>between</i> two different capabilities — a static inside DocumentFormat.OpenXml
    /// touched by both the DOCX and PPTX paths, for instance.
    ///
    /// <para>Every operation is worker-distinct for the same reason as above: identical inputs
    /// would make cross-contamination invisible.</para>
    /// </summary>
    [Fact]
    public async Task EveryCapabilityAtOnceProducesCorrectResults()
    {
        var operations = new (string Name, Func<int, Task<string>> Run)[]
        {
            ("html->docx", async w => DocxEditor.ExtractText(await HtmlToDocxConverter.ConvertAsync(HtmlFor(w)))),
            ("docx create", w => Task.Run(() => DocxEditor.ExtractText(
                DocxEditor.Create(new[] { DocxBlock.Paragraph($"Body {w}") })))),
            ("docx edit", w => Task.Run(() => DocxEditor.ExtractText(DocxEditor.ReplaceText(
                DocxEditor.Create(new[] { DocxBlock.Paragraph("Hello {{who}}.") }),
                new Dictionary<string, string> { ["{{who}}"] = $"World{w}" })))),
            ("xlsx write", w => Task.Run(() => WorkbookEditor.ReadCell(
                WorkbookEditor.Create("S", new[] { new object?[] { $"cell{w}" } }), "S", "A1"))),
            ("xlsx edit", w => Task.Run(() => WorkbookEditor.ReadCell(WorkbookEditor.SetCell(
                WorkbookEditor.Create("S", new[] { new object?[] { "before" } }), "S", "A1", $"after{w}"),
                "S", "A1"))),
            ("pptx", w => Task.Run(() => string.Join("~", PresentationEditor.ExtractText(
                PresentationEditor.Create(new[] { PptxSlide.Titled($"Title {w}", $"bullet {w}") }))))),
        };

        // Each (operation, worker) pair's correct answer, established with nothing else running.
        var expected = new string[operations.Length, Concurrency];
        for (var op = 0; op < operations.Length; op++)
            for (var w = 0; w < Concurrency; w++)
                expected[op, w] = await operations[op].Run(w);

        var started = new TaskCompletionSource();
        var workers = Enumerable.Range(0, Concurrency)
            .SelectMany(w => operations.Select((op, index) => (Op: op, Index: index, Worker: w)))
            .Select(async pair =>
            {
                await started.Task;
                return (pair.Index, pair.Worker, pair.Op.Name, Result: await pair.Op.Run(pair.Worker));
            })
            .ToArray();

        started.SetResult();
        var results = await Task.WhenAll(workers);

        foreach (var (index, worker, name, result) in results)
            Assert.True(
                string.Equals(expected[index, worker], result, StringComparison.Ordinal),
                $"'{name}' worker {worker} produced the wrong result under concurrency.\n" +
                $"  expected: {expected[index, worker]}\n" +
                $"  actual:   {result}");

        _output.WriteLine(
            $"{results.Length} operations across {operations.Length} capabilities, " +
            $"{Concurrency} distinct inputs each, all correct");
    }
}
