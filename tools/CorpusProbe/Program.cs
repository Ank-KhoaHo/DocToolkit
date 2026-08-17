using System.Diagnostics;
using System.Text;
using DocToolkit;

// Runs every file in a directory through the conversion it is eligible for, and reports the rate
// and the failure frames. Reports; never judges. See .github/workflows/corpus.yml for why.

Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

var root = args.FirstOrDefault();
if (root is null || !Directory.Exists(root))
{
    Console.Error.WriteLine("usage: corpus-probe <directory> [--limit N]");
    return 2;
}

var limit = int.MaxValue;
var limitIndex = Array.IndexOf(args, "--limit");
if (limitIndex >= 0 && limitIndex + 1 < args.Length && int.TryParse(args[limitIndex + 1], out var parsed))
    limit = parsed;

// Reading a legacy corpus as UTF-8 unconditionally is how the first measurement of it went wrong:
// most of govdocs1 is windows-1252, so invalid bytes became U+FFFD and the renderer then correctly
// refused to encode them - reported at the time as nine library failures that did not exist.
// Strict decoding distinguishes "is UTF-8" from "decodes to something"; the lenient default cannot,
// because it substitutes silently.
static string ReadText(string path)
{
    var bytes = File.ReadAllBytes(path);
    try
    {
        return new UTF8Encoding(false, throwOnInvalidBytes: true).GetString(bytes);
    }
    catch (DecoderFallbackException)
    {
        return Encoding.GetEncoding(1252).GetString(bytes);
    }
}

static string Frame(Exception ex)
{
    var inner = ex.InnerException ?? ex;
    var line = (inner.StackTrace ?? string.Empty)
        .Split('\n')
        .Select(l => l.Trim())
        .FirstOrDefault(l => l.StartsWith("at ", StringComparison.Ordinal)) ?? "(no frame)";
    var paren = line.IndexOf('(');
    if (paren > 0) line = line[..paren];
    return $"{inner.GetType().Name} @ {line.Replace("at ", string.Empty)}";
}

var conversions = new (string Name, string[] Extensions, Func<string, Task> Run)[]
{
    ("HTML -> DOCX", [".html", ".htm"], async p => await HtmlToDocxConverter.ConvertAsync(ReadText(p))),
    ("HTML -> PDF",  [".html", ".htm"], async p => await HtmlToPdfConverter.ConvertAsync(ReadText(p))),
    ("DOC  -> DOCX", [".doc"],  p => Task.Run(() => DocToDocxConverter.Convert(File.ReadAllBytes(p), new LegacyDocOptions { AllowContentLoss = true }))),
    ("PPT  -> PDF",  [".ppt"],  p => Task.Run(() => PptxToPdfConverter.Convert(File.ReadAllBytes(p)))),
    ("PPTX -> PDF",  [".pptx"], p => Task.Run(() => PptxToPdfConverter.Convert(File.ReadAllBytes(p)))),
    ("XLSX -> PDF",  [".xlsx"], p => Task.Run(() => XlsxToPdfConverter.Convert(File.ReadAllBytes(p)))),
};

var exit = 0;
var summary = new StringBuilder();
summary.AppendLine("| conversion | files | converted | rate |");
summary.AppendLine("|---|---:|---:|---:|");

var detail = new StringBuilder();

foreach (var (name, extensions, run) in conversions)
{
    var files = extensions
        .SelectMany(e => Directory.EnumerateFiles(root, "*" + e, SearchOption.AllDirectories))
        .Distinct()
        .OrderBy(f => f, StringComparer.Ordinal)
        .Take(limit)
        .ToArray();

    if (files.Length == 0) continue;

    var ok = 0;
    var frames = new Dictionary<string, int>(StringComparer.Ordinal);
    var stopwatch = Stopwatch.StartNew();

    foreach (var file in files)
    {
        try
        {
            await run(file);
            ok++;
        }
        catch (Exception ex)
        {
            var key = Frame(ex);
            frames[key] = frames.TryGetValue(key, out var n) ? n + 1 : 1;
        }
    }

    stopwatch.Stop();
    var rate = 100.0 * ok / files.Length;
    Console.WriteLine($"{name}: {ok}/{files.Length} ({rate:F1}%) in {stopwatch.Elapsed.TotalSeconds:F0}s");
    summary.AppendLine($"| {name} | {files.Length} | {ok} | **{rate:F1}%** |");

    if (frames.Count > 0)
    {
        detail.AppendLine($"### {name}");
        detail.AppendLine();
        foreach (var kv in frames.OrderByDescending(k => k.Value))
        {
            Console.WriteLine($"    {kv.Value,4}x  {kv.Key}");
            detail.AppendLine($"- **{kv.Value}x** `{kv.Key}`");
        }
        detail.AppendLine();
    }
}

var summaryPath = Environment.GetEnvironmentVariable("GITHUB_STEP_SUMMARY");
if (!string.IsNullOrEmpty(summaryPath))
{
    // REPORTS, NEVER FAILS. A rate is not a pass/fail signal: it moves with the corpus slice as
    // much as with the code, and a threshold here would be switched off the first time a chunk
    // happened to contain more legacy files than the last one.
    await File.AppendAllTextAsync(summaryPath,
        "## Real-world corpus\n\n" + summary + "\n" + detail
        + "\nRates, not a gate. They move with the corpus slice as much as with the code.\n");
}

return exit;
