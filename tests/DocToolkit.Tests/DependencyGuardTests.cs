using System.Reflection;
using System.Text.RegularExpressions;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// Guards the licence and platform constraints from the research spec. These are the two
/// mistakes that are cheap to make and expensive to discover in production.
/// </summary>
public class DependencyGuardTests
{
    /// <summary>
    /// Matched as an exact assembly name or as a namespace prefix, so "Spire" also catches
    /// Spire.Doc, Spire.Xls and the rest of the family.
    /// </summary>
    private static readonly string[] BannedAssemblies =
    {
        "System.Drawing.Common",  // throws PlatformNotSupportedException on Linux (.NET 7+)
        "SkiaSharp",              // pulls native binaries
        "EPPlus",                 // Polyform Noncommercial - not free for commercial use
        "NPOI",                   // >= 2.8.0 requires a paid maintenance fee
        "Magick.NET-Q16-AnyCPU",  // pulls native binaries; arrived transitively via ShapeCrawler
        "ShapeCrawler",           // removed - pulled SkiaSharp + Magick.NET, 664 MB of native runtimes
        "Spire",                  // feature-capped free editions
        "Syncfusion",             // revenue-gated community licence
        "QuestPDF",               // revenue-gated community licence
        "IronPdf",                // commercial
    };

    /// <summary>Shared-object filenames, including versioned ones such as libfoo.so.1.2.</summary>
    private static readonly Regex NativeLibrary =
        new(@"(\.so(\.\d+)*|\.dylib)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

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

        var violations = seen.Where(name => BannedAssemblies.Any(banned =>
            name.Equals(banned, StringComparison.OrdinalIgnoreCase) ||
            name.StartsWith(banned + ".", StringComparison.OrdinalIgnoreCase))).ToList();

        Assert.True(violations.Count == 0,
            $"Banned assemblies referenced: {string.Join(", ", violations)}");
    }

    [Fact]
    public void NoNativeBinariesAreCopiedToOutput()
    {
        var outputDir = Path.GetDirectoryName(typeof(DependencyGuardTests).Assembly.Location)!;
        var native = Directory.EnumerateFiles(outputDir, "*", SearchOption.AllDirectories)
                              .Where(f => NativeLibrary.IsMatch(f))
                              .ToList();

        Assert.True(native.Count == 0,
            $"Unexpected native binaries in output: {string.Join(", ", native.Select(Path.GetFileName))}");
    }

    [Fact]
    public void SixLaborsFontsStaysOnTheApacheLicensedOneX()
    {
        // OfficeIMO.Word asks for [1.0.0, 3.0.0). SixLabors.Fonts 1.x is Apache-2.0; 2.x switched
        // to the Six Labors Split License 1.0, which is Apache-2.0 only below $1M annual revenue
        // and commercial above. Floating into 2.x would quietly take DocToolkit off permissive
        // licensing, so DocToolkit.csproj pins an exact version on the 1.x line and this catches
        // the pin being dropped. Asserted on the MAJOR, so a patch bump of the pin does not have
        // to touch this test - which is why it kept passing while five comments still said 1.0.0.
        var loaded = Assembly.Load(new AssemblyName("SixLabors.Fonts"));
        var version = loaded.GetName().Version;

        Assert.NotNull(version);
        Assert.True(version!.Major == 1,
            $"SixLabors.Fonts resolved to {version} - 2.x is not Apache-2.0. See THIRD-PARTY-NOTICES.txt.");
    }
}
