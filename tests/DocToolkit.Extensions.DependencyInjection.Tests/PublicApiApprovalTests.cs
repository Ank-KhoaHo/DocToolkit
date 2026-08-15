using System.Reflection;
using PublicApiGenerator;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

/// <summary>
/// Pins the DI extensions package's public surface, for the same reason the core project does.
///
/// This one has extra history behind it: the 0.3.0 changelog records adding members to these six
/// shipped interfaces, which is source-breaking for anyone who implemented one by hand. That change
/// was fine and deliberate — the point is that nothing flagged it at the time.
/// </summary>
public class PublicApiApprovalTests
{
    [Fact]
    public void PublicApi_MatchesTheApprovedSurface()
        => ApiApproval.Verify(
            typeof(ServiceCollectionExtensions).Assembly,
            "DocToolkit.Extensions.DependencyInjection");
}

internal static class ApiApproval
{
    public static void Verify(Assembly assembly, string name)
    {
        var actual = Normalise(assembly.GeneratePublicApi(new ApiGeneratorOptions
        {
            ExcludeAttributes = new[]
            {
                "System.Runtime.Versioning.TargetFrameworkAttribute",
                "System.Reflection.AssemblyMetadataAttribute",
                "System.Runtime.CompilerServices.InternalsVisibleToAttribute",
            },
        }));

        var approvedPath = Path.Join(AppContext.BaseDirectory, "PublicApi", $"{name}.approved.txt");

        if (!File.Exists(approvedPath))
        {
            throw new InvalidOperationException(
                $"No approved API file at '{approvedPath}'. Create " +
                $"tests/.../PublicApi/{name}.approved.txt with the surface below, then re-run.\n\n" +
                actual);
        }

        var approved = Normalise(File.ReadAllText(approvedPath));
        if (string.Equals(approved, actual, StringComparison.Ordinal)) return;

        var receivedPath = Path.Join(AppContext.BaseDirectory, "PublicApi", $"{name}.received.txt");
        Directory.CreateDirectory(Path.GetDirectoryName(receivedPath)!);
        File.WriteAllText(receivedPath, actual);

        var removed = approved.Split('\n').Except(actual.Split('\n'), StringComparer.Ordinal).ToList();
        var added = actual.Split('\n').Except(approved.Split('\n'), StringComparer.Ordinal).ToList();

        var report = new System.Text.StringBuilder()
            .AppendLine($"The public API of {name} changed.")
            .AppendLine()
            .AppendLine("REMOVED or CHANGED — these break existing callers:");
        foreach (var line in removed.Where(l => l.Length > 0).Take(25))
            report.AppendLine("  - " + line.Trim());

        report.AppendLine().AppendLine("ADDED — additive for callers, but SOURCE-BREAKING for anyone");
        report.AppendLine("implementing one of these interfaces by hand:");
        foreach (var line in added.Where(l => l.Length > 0).Take(25))
            report.AppendLine("  + " + line.Trim());

        report.AppendLine()
              .AppendLine("If intended, copy the full new surface over the approved file:")
              .AppendLine($"  from: {receivedPath}")
              .AppendLine($"  to:   tests/.../PublicApi/{name}.approved.txt");

        Assert.Fail(report.ToString());
    }

    private static string Normalise(string text) =>
        text.Replace("\r\n", "\n", StringComparison.Ordinal).TrimEnd();
}
