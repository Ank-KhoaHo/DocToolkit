using System.Reflection;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// Guards the licence and platform constraints from the research spec. These are the two
/// mistakes that are cheap to make and expensive to discover in production.
/// </summary>
public class DependencyGuardTests
{
    private static readonly string[] BannedAssemblies =
    {
        "System.Drawing.Common",  // throws PlatformNotSupportedException on Linux (.NET 7+)
        "SkiaSharp",              // pulls native binaries
        "EPPlus",                 // Polyform Noncommercial - not free for commercial use
        "NPOI",                   // >= 2.8.0 requires a paid maintenance fee
        "Magick.NET-Q16-AnyCPU",  // pulls native binaries; arrived transitively via ShapeCrawler
        "ShapeCrawler",           // removed - pulled SkiaSharp + Magick.NET, 664 MB of native runtimes
    };

    [Fact]
    public void DocToolkit_DoesNotReferenceBannedAssemblies()
    {
        var toolkit = typeof(HtmlToDocxConverter).Assembly;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var queue = new Queue<Assembly>();
        queue.Enqueue(toolkit);

        while (queue.Count > 0)
        {
            foreach (var reference in queue.Dequeue().GetReferencedAssemblies())
            {
                if (!seen.Add(reference.Name!)) continue;
                try { queue.Enqueue(Assembly.Load(reference)); }
                catch { /* not all references load standalone; the name check below still applies */ }
            }
        }

        var violations = seen.Where(name =>
            BannedAssemblies.Any(b => name.Equals(b, StringComparison.OrdinalIgnoreCase))).ToList();

        Assert.True(violations.Count == 0,
            $"Banned assemblies referenced: {string.Join(", ", violations)}");
    }

    [Fact]
    public void NoNativeBinariesAreCopiedToOutput()
    {
        var outputDir = Path.GetDirectoryName(typeof(DependencyGuardTests).Assembly.Location)!;
        var native = Directory.EnumerateFiles(outputDir, "*", SearchOption.AllDirectories)
                              .Where(f => f.EndsWith(".so", StringComparison.OrdinalIgnoreCase)
                                       || f.EndsWith(".dylib", StringComparison.OrdinalIgnoreCase))
                              .ToList();

        Assert.True(native.Count == 0,
            $"Unexpected native binaries in output: {string.Join(", ", native.Select(Path.GetFileName))}");
    }
}
