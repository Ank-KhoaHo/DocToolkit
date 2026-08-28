using System.Reflection;
using System.Runtime.Loader;
using DocToolkit.TestSupport;
using Mono.Cecil;
using Mono.Cecil.Cil;
using PublicApiGenerator;

namespace DocToolkit.Tests;

/// <summary>
/// Pins the two properties of <see cref="ApiSurface"/> that are not visible from an ordinary run:
/// that it still describes an assembly correctly when a STALE symbol file sits beside it, and that
/// it does not report a type STRYKER ITSELF injected as part of the shipped surface.
///
/// Both are things Stryker's instrumentation does to an assembly, and until 2026-08-27/28 both made
/// <see cref="PublicApiApprovalTests"/> fail in every mutation run for reasons having nothing to do
/// with the mutant under test — handing every static mutant in the affected assembly a free kill.
/// See <see cref="ApiSurface"/> for the measurement and B30 for the cost. **The two are independent
/// defects, not one**: fixing the symbol mismatch alone left the second one live, and a run that
/// only checked "did an exception fire" missed it, because the second one is a genuine, correctly
/// reported <c>Assert.Fail</c> — the failure is real, it is just about the wrong thing.
///
/// Both conditions are built here rather than by running Stryker: costs milliseconds instead of a
/// mutation run, and runs on every pull request rather than weekly.
///
/// A THIRD thing came from building this file, not from the original defect: both fixtures below
/// load a second, same-named copy of the currently-mutated assembly, and doing so under a real
/// Stryker run corrupted coverage attribution for OTHER, unrelated tests — see
/// <see cref="ApiSurface.WithADistinctIdentity"/>. There is no cheap unit-level positive control
/// for that one; its only observable effect is on Stryker's own coverage bookkeeping, which nothing
/// short of a real mutation run exercises. It was verified by a scoped Stryker diagnostic instead:
/// 154 corrupted mutants across ten unrelated files with the rename removed, 0 with it restored,
/// scoped to the same two files both times.
/// </summary>
public class ApiSurfaceTests
{
    /// <summary>
    /// The assembly under test is the DOCX one throughout this file, for no reason other than that
    /// it is one this project already loads. The approved file is looked up from the assembly's own
    /// name rather than a second constant beside it, so changing which assembly is used cannot leave
    /// the two disagreeing.
    /// </summary>
    private static readonly Assembly Subject = typeof(DocxEditor).Assembly;

    [Fact]
    public void Generate_DescribesTheAssembly_WhenAMismatchedSymbolFileSitsBesideIt()
    {
        using var fixture = WithAMismatchedSymbolFile();
        var generated = ApiSurface.Normalise(ApiSurface.Generate(fixture.Assembly));

        // The literal the approval test itself compares against, so this cannot pass by generating
        // something merely non-empty.
        Assert.Equal(Approved(), generated);
    }

    /// <summary>
    /// The positive control for the symbol mismatch. <see cref="ApiSurface"/> copies the assembly
    /// only because the unprotected call fails on this input; if a future PublicApiGenerator
    /// tolerates a mismatched symbol file, this test fails and says the copying can be removed —
    /// rather than the test above passing vacuously for a reason that has nothing to do with the
    /// guard.
    /// </summary>
    [Fact]
    public void TheUnprotectedCall_ThrowsOnThatSameAssembly_SoTheCopyingIsLoadBearing()
    {
        using var fixture = WithAMismatchedSymbolFile();

        Assert.Throws<SymbolsNotMatchingException>(
            () => fixture.Assembly.GeneratePublicApi(ApiSurface.Options));
    }

    [Fact]
    public void Generate_DescribesTheAssembly_WhenAStrykerLikeTypeHasBeenInjectedIntoIt()
    {
        using var fixture = WithAnInjectedStrykerType();
        var generated = ApiSurface.Normalise(ApiSurface.Generate(fixture.Assembly));

        Assert.Equal(Approved(), generated);
    }

    /// <summary>
    /// The positive control for the namespace filter. Confirms the injected type really is public
    /// API by PublicApiGenerator's own unfiltered judgement — so the test above passing is evidence
    /// the filter did the excluding, not evidence the injection built nothing observable.
    /// </summary>
    [Fact]
    public void TheUnfilteredCall_DoesReportTheInjectedType_SoTheDenyListIsLoadBearing()
    {
        using var fixture = WithAnInjectedStrykerType();

        var unfiltered = fixture.Assembly.GeneratePublicApi(new ApiGeneratorOptions());

        Assert.Contains("FreeKillIfNotDenied", unfiltered, StringComparison.Ordinal);
    }

    /// <summary>
    /// A copy of the subject assembly, under a DIFFERENT assembly identity (see
    /// <see cref="ApiSurface.WithADistinctIdentity"/>), with the PPTX assembly's symbols renamed to
    /// sit beside it. The copy is loaded in its own context so that <see cref="Assembly.Location"/>
    /// is the temp path — which is what Cecil opens, and therefore what decides which <c>.pdb</c>
    /// it finds.
    /// </summary>
    private static LoadedFixture WithAMismatchedSymbolFile()
    {
        var subject = Subject.Location;
        var alienSymbols = Path.ChangeExtension(typeof(PresentationEditor).Assembly.Location, ".pdb");

        // Without symbols on disk there is no mismatch to build, and both tests would report
        // something untrue about the guard rather than failing.
        Assert.True(
            File.Exists(alienSymbols),
            $"No symbol file at '{alienSymbols}'. This suite needs one to build the mismatch it " +
            "exists to test; a build that emits no .pdb cannot exercise it.");

        var directory = Directory.CreateTempSubdirectory("doctoolkit-stale-pdb-");
        var copy = Path.Join(directory.FullName, Path.GetFileName(subject));
        ApiSurface.WithADistinctIdentity(subject, copy);
        File.Copy(alienSymbols, Path.ChangeExtension(copy, ".pdb"));

        return LoadedFixture.Load("stale-pdb", directory, copy);
    }

    /// <summary>
    /// A copy of the subject assembly, under a DIFFERENT assembly identity, carrying one extra
    /// public type added with Mono.Cecil the way Stryker adds <c>MutantControl</c> — a namespace it
    /// did not have before, containing a type a consumer never wrote. The name is deliberately NOT
    /// "Stryker" alone: proves the filter matches by PREFIX the way Stryker's randomised-suffix
    /// namespace needs, rather than by an exact string that would only ever match a fixture built to
    /// spell it out.
    /// </summary>
    private static LoadedFixture WithAnInjectedStrykerType()
    {
        var directory = Directory.CreateTempSubdirectory("doctoolkit-injected-type-");
        var copy = Path.Join(directory.FullName, Path.GetFileName(Subject.Location));

        // Adds the type in the SAME Cecil session WithADistinctIdentity already opens for the
        // rename, via its mutate hook, rather than duplicating the rename logic here.
        ApiSurface.WithADistinctIdentity(Subject.Location, copy, module =>
        {
            var injected = new TypeDefinition(
                "StrykerXYZ123",
                "FreeKillIfNotDenied",
                Mono.Cecil.TypeAttributes.Public | Mono.Cecil.TypeAttributes.Abstract | Mono.Cecil.TypeAttributes.Sealed,
                module.TypeSystem.Object);
            module.Types.Add(injected);
        });

        return LoadedFixture.Load("injected-type", directory, copy);
    }

    private static string Approved() =>
        ApiSurface.Normalise(File.ReadAllText(Path.Join(
            AppContext.BaseDirectory, "PublicApi", $"{Subject.GetName().Name}.approved.txt")));

    /// <summary>
    /// A loaded fixture assembly plus the temp directory and collectible
    /// <see cref="AssemblyLoadContext"/> that back it, released together on <see cref="Dispose"/>.
    /// <see cref="ApiSurface.Generate"/> can do this cleanup inside one method call because its
    /// copy never needs to outlive that call; these fixtures hand the loaded <see cref="Assembly"/>
    /// back to a caller that is still using it, so the same best-effort cleanup has to wait for
    /// that caller to finish — hence <c>using var fixture = ...</c> at each call site instead.
    /// </summary>
    private readonly struct LoadedFixture : IDisposable
    {
        public Assembly Assembly { get; }
        private readonly DirectoryInfo _directory;
        private readonly AssemblyLoadContext _context;

        private LoadedFixture(Assembly assembly, DirectoryInfo directory, AssemblyLoadContext context)
        {
            Assembly = assembly;
            _directory = directory;
            _context = context;
        }

        public static LoadedFixture Load(string contextName, DirectoryInfo directory, string assemblyPath)
        {
            var context = new AssemblyLoadContext(contextName, isCollectible: true);
            return new LoadedFixture(context.LoadFromAssemblyPath(assemblyPath), directory, context);
        }

        public void Dispose()
        {
            _context.Unload();

            // Best effort, same reasoning as ApiSurface.Generate's own cleanup: unload is
            // asynchronous, so the copy is usually still mapped here.
            try
            {
                _directory.Delete(recursive: true);
            }
            catch (Exception)
            {
            }
        }
    }
}
