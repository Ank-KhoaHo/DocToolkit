using System.Reflection;
using PublicApiGenerator;

namespace DocToolkit.Tests;

/// <summary>
/// Pins the shipped public API to an approved text file.
///
/// `CLAUDE.md` says changing an existing name or signature is a breaking change for consumers, and
/// until now nothing checked that mechanically — the 0.3.0 changelog already records one accidental
/// source-breaking change that reached a release. The argument is sharper than it looks, because
/// this package stays below 1.0.0 permanently: semver will never start protecting consumers on its
/// own here, so a mechanical guard is the only thing that would.
///
/// Deliberately no Verify/ApprovalTests package. Those pull in Argon and DiffEngine, and DiffEngine
/// launches external diff tools — a lot of machinery for comparing a string to a file. Same call
/// this repo made writing its own PNG/JPEG header parsers rather than taking an image library.
/// </summary>
public class PublicApiApprovalTests
{
    // ONE CALL PER SHIPPED ASSEMBLY. The package contains seven, since the per-format project
    // split - but they are all packed into the one nupkg, so together these pin exactly the
    // surface a consumer receives.
    //
    // EACH LOCATOR TYPE MUST BE ONE ITS OWN PROJECT GENUINELY OWNS. `Verify` finds the assembly
    // through the type, so a locator that later moves to another project silently repoints the
    // whole file - the approved diff would then lose everything that stayed behind and gain only
    // what moved, which is a large diff on a change that altered no API at all. Check
    // artifacts/split-assignment.txt before changing one.
    //
    // There is deliberately NO test for DocToolkit.Html: all five of its files are internal, so it
    // has no public surface. An approved file parsing to zero types is a case
    // check-readme-coverage.py fails on by design, rather than passing vacuously.

    // The locator was `typeof(DocxEditor)` until DocxEditor moved to DocToolkit.Docx, at which
    // point this line silently began pinning a DIFFERENT assembly - the approved file would have
    // lost everything that stayed behind and gained only the DOCX types, a huge diff on a change
    // that altered no API at all. HtmlToPdfConverter is a converter, and converters stay here by
    // design, so it is a stable locator.
    [Fact]
    public void PublicApi_MatchesTheApprovedSurface()
        => ApiApproval.Verify(typeof(HtmlToPdfConverter).Assembly, "DocToolkit");

    [Fact]
    public void PublicApi_Primitives_MatchesTheApprovedSurface()
        => ApiApproval.Verify(typeof(PageSetup).Assembly, "DocToolkit.Primitives");

    [Fact]
    public void PublicApi_Docx_MatchesTheApprovedSurface()
        => ApiApproval.Verify(typeof(DocxEditor).Assembly, "DocToolkit.Docx");

    [Fact]
    public void PublicApi_Pdf_MatchesTheApprovedSurface()
        => ApiApproval.Verify(typeof(PdfEditor).Assembly, "DocToolkit.Pdf");

    [Fact]
    public void PublicApi_Pptx_MatchesTheApprovedSurface()
        => ApiApproval.Verify(typeof(PresentationEditor).Assembly, "DocToolkit.Pptx");

    [Fact]
    public void PublicApi_Xlsx_MatchesTheApprovedSurface()
        => ApiApproval.Verify(typeof(WorkbookEditor).Assembly, "DocToolkit.Xlsx");
}

/// <summary>
/// Shared by both test projects — the core assembly and the DI extensions assembly each pin their
/// own surface.
/// </summary>
internal static class ApiApproval
{
    public static void Verify(Assembly assembly, string name)
    {
        var actual = Normalise(assembly.GeneratePublicApi(new ApiGeneratorOptions
        {
            // Excluded because they are not part of the surface a consumer codes against, and they
            // change with the build rather than with the API: including them would make the
            // approved file churn on every release for no signal.
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

        // Written next to the test output so the whole new surface can be copied over the approved
        // file in one step, rather than reconstructed from a diff.
        var receivedPath = Path.Join(AppContext.BaseDirectory, "PublicApi", $"{name}.received.txt");
        Directory.CreateDirectory(Path.GetDirectoryName(receivedPath)!);
        File.WriteAllText(receivedPath, actual);

        var approvedLines = approved.Split('\n');
        var actualLines = actual.Split('\n');
        var removed = approvedLines.Except(actualLines, StringComparer.Ordinal).ToList();
        var added = actualLines.Except(approvedLines, StringComparer.Ordinal).ToList();

        var report = new System.Text.StringBuilder()
            .AppendLine($"The public API of {name} changed.")
            .AppendLine()
            .AppendLine("REMOVED or CHANGED — these break existing callers:");
        foreach (var line in removed.Where(l => l.Length > 0).Take(25))
            report.AppendLine("  - " + line.Trim());

        report.AppendLine().AppendLine("ADDED — additive, safe for existing callers:");
        foreach (var line in added.Where(l => l.Length > 0).Take(25))
            report.AppendLine("  + " + line.Trim());

        report.AppendLine()
              .AppendLine("If the change is intended, copy the full new surface over the approved file:")
              .AppendLine($"  from: {receivedPath}")
              .AppendLine($"  to:   tests/.../PublicApi/{name}.approved.txt")
              .AppendLine("and include that diff in the pull request, so the API change is reviewed rather than noticed later.");

        Assert.Fail(report.ToString());
    }

    /// <summary>Line endings only — the file is committed and checked out on both Windows and Linux.</summary>
    private static string Normalise(string text) =>
        text.Replace("\r\n", "\n", StringComparison.Ordinal).TrimEnd();
}
