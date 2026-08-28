using System.Reflection;
using System.Runtime.Loader;
using Mono.Cecil;
using PublicApiGenerator;

namespace DocToolkit.TestSupport;

/// <summary>
/// Generates the public surface of a shipped assembly, for the approval tests in BOTH test
/// projects. The file is linked into <c>DocToolkit.Extensions.DependencyInjection.Tests</c> rather
/// than copied — the two projects' <c>ApiApproval</c> classes already carry two copies of the
/// comparison and report wording, and one of them once claimed in a doc comment to be shared when
/// it was not.
///
/// <para>
/// WHY THE ASSEMBLY IS COPIED BEFORE IT IS READ, which is the whole reason this class exists.
/// <c>GeneratePublicApi</c> does not reflect over the <see cref="Assembly"/> it is handed; it opens
/// <see cref="Assembly.Location"/> with Mono.Cecil, and Cecil reads the <c>.pdb</c> sitting beside
/// that file. A tool that rewrites the assembly and leaves the original symbols in place therefore
/// makes the read throw <c>SymbolsNotMatchingException</c> — a failure that says nothing whatever
/// about the public API.
/// </para>
///
/// <para>
/// Measured 2026-08-27, and it was not hypothetical: Stryker writes its instrumented build over
/// <c>tests/DocToolkit.Tests/bin/…/DocToolkit.*.dll</c> and does not touch the <c>.pdb</c>, so all
/// six approval tests threw in every mutation run. Stryker cannot attribute a <b>static</b> mutant
/// to a test, so it runs the whole suite and counts any failure as a kill — and those six handed
/// every static mutant a free kill. 26 mutants on the 2026-08-27 run, 4 of them in
/// <c>GuardedResourceLoader</c>, the network path the suite exists to protect. The reported score
/// was 96.51% against an honest 92.88%, under the <c>break: 95</c> gate. See B30.
/// </para>
///
/// <para>
/// Copying the assembly into a directory with no symbols beside it removes the <c>.pdb</c> from the
/// question entirely, rather than tolerating a mismatched one. It is done on EVERY call, not as a
/// fallback: a second path taken only under instrumentation is a path CI never executes.
/// Measured cost for the six core assemblies: 494 ms against 180 ms, once per test run.
/// </para>
/// </summary>
internal static class ApiSurface
{
    /// <summary>
    /// Excluded because they are not part of the surface a consumer codes against, and they change
    /// with the build rather than with the API: including them would make the approved file churn
    /// on every release for no signal.
    /// </summary>
    public static readonly ApiGeneratorOptions Options = new()
    {
        ExcludeAttributes = new[]
        {
            "System.Runtime.Versioning.TargetFrameworkAttribute",
            "System.Reflection.AssemblyMetadataAttribute",
            "System.Runtime.CompilerServices.InternalsVisibleToAttribute",
        },

        // A SECOND, independent defect from the same cause (B30): Stryker injects a genuinely
        // PUBLIC `MutantControl` type into whatever assembly it mutates, in a namespace named
        // "Stryker" plus a random suffix per build. That is a real difference in the assembly's
        // surface, not an artefact of a stale symbol file - the symbol-free copy above does
        // nothing about it, because Cecil reads it correctly. Denying the "Stryker" prefix is
        // harmless on an un-instrumented build, where no such namespace exists.
        DenyNamespacePrefixes = new[] { "Stryker" },
    };

    /// <summary>
    /// The public API of <paramref name="assembly"/>, read from a copy that has no symbol file
    /// beside it and a DIFFERENT assembly identity. See the class summary for the symbol half, and
    /// <see cref="WithADistinctIdentity"/> for the identity half — both are load-bearing.
    /// </summary>
    /// <param name="assembly">The shipped assembly to describe.</param>
    public static string Generate(Assembly assembly)
    {
        var directory = Directory.CreateTempSubdirectory("doctoolkit-api-");
        try
        {
            var copy = Path.Join(directory.FullName, Path.GetFileName(assembly.Location));
            WithADistinctIdentity(assembly.Location, copy);

            // Its own collectible context: LoadFrom would hand back the ALREADY-LOADED assembly,
            // whose Location is the original path, and the copy would do nothing.
            var context = new AssemblyLoadContext("api-surface", isCollectible: true);
            try
            {
                return context.LoadFromAssemblyPath(copy).GeneratePublicApi(Options);
            }
            finally
            {
                context.Unload();
            }
        }
        finally
        {
            // Unloading is asynchronous, so the copy is usually still mapped here and the delete
            // fails with UnauthorizedAccessException. Best effort: the directory is under TEMP and
            // holds one managed assembly.
            //
            // UNFILTERED on purpose, and this is the one place in this repository where that is the
            // safer choice rather than the lazy one: this runs in a `finally`, so anything it lets
            // escape REPLACES the exception the caller actually needs to see. A filtered catch that
            // does not match never runs its body - the same mechanism that once leaked a
            // MemoryStream out of a converter, recorded under CodeQL in CLAUDE.md.
            try
            {
                directory.Delete(recursive: true);
            }
            catch (Exception)
            {
            }
        }
    }

    /// <summary>
    /// Writes a copy of <paramref name="source"/> to <paramref name="destination"/> with a
    /// DIFFERENT, per-call-unique simple assembly name — a THIRD thing this class exists to work
    /// around, found while re-measuring B30's own fix under a real Stryker run.
    ///
    /// <para>
    /// A second, differently-named copy of an assembly Stryker is currently mutating is harmless.
    /// A second copy sharing a name already loaded is not: measured 2026-08-28, loading one — via
    /// <c>File.Copy</c> plus <c>AssemblyLoadContext.LoadFromAssemblyPath</c>, with no symbol read,
    /// no reflection past <c>GeneratePublicApi</c> and nothing that should touch Stryker's own
    /// coverage machinery — corrupted coverage attribution for OTHER, unrelated tests running in
    /// the same process. One scoped run over two files went from 4 legitimate `NoCoverage`
    /// mutants (present even in the symbol-only fix, tolerated as a small residual) to 154 across
    /// ten unrelated files once a second pair of tests started loading a second same-named copy of
    /// their own; skipping only those two tests, nothing else changed, returned it to 4.
    /// </para>
    ///
    /// <para>
    /// A FIXED renamed suffix — one shared by every call — closed the scoped case (0
    /// `NoCoverage` on the same two files) but NOT the full 16-file scope: 110 of 565 mutants
    /// still corrupted, essentially unchanged. Stryker reuses one test host process across many
    /// mutant iterations and does not promptly unload each call's own collectible
    /// <see cref="AssemblyLoadContext"/>, so over enough iterations several copies alive at once
    /// and sharing ONE identity collide with EACH OTHER the same way the original bug collided
    /// with the real assembly. A per-call GUID — a name that cannot repeat — cut that to 18 of
    /// 657 at full scope. <b>That is an improvement, not a fix</b>: hand-verifying a sample of
    /// what still regresses at full scope finds it is STILL happening, at roughly a sixth of the
    /// prior rate. See B30's "third cause" section for the full account, including two untried
    /// directions for closing the rest of it, and why the remaining size and (pessimistic)
    /// direction of the gap made this an acceptable place to stop rather than spend a further
    /// 26-minute re-measurement chasing it.
    /// </para>
    ///
    /// <para>
    /// The exact mechanism inside Stryker's data collector was never identified precisely, in
    /// either version. Renaming is a Mono.Cecil metadata edit, not an IL rewrite: it changes
    /// <see cref="AssemblyNameDefinition.Name"/> and the module's own name, and nothing a mutant or
    /// `MutantControl` reference resolves by. It also does not re-introduce the symbol problem -
    /// Cecil's default <see cref="WriterParameters"/> writes no symbols, same as the plain file
    /// copy it replaces, so the copy still has no `.pdb` for the FIRST fix to worry about.
    /// </para>
    /// </summary>
    /// <param name="source">The assembly to copy.</param>
    /// <param name="destination">Where to write the renamed copy.</param>
    /// <param name="mutate">
    /// Applied to the in-memory module before the identity rename and write — the hook
    /// <see cref="ApiSurfaceTests"/> uses to add its injected-type fixture in the same Cecil
    /// session, rather than duplicating the rename afterward.
    /// </param>
    internal static void WithADistinctIdentity(
        string source, string destination, Action<ModuleDefinition>? mutate = null)
    {
        using var module = ModuleDefinition.ReadModule(source);

        mutate?.Invoke(module);

        // A per-call GUID, not a fixed suffix. A FIXED one only fixed the small case: re-measured
        // at full scope, 110 of 565 tested mutants still regressed to NoCoverage/Survived, because
        // Stryker reuses the test host process across many mutant runs and does not promptly
        // unload each call's own collectible AssemblyLoadContext - so, over enough iterations,
        // multiple copies sharing ONE fixed renamed identity collide with EACH OTHER the same way
        // the original bug collided with the real one. A GUID per call cannot repeat.
        module.Assembly.Name.Name += "." + Guid.NewGuid().ToString("N");
        module.Name = Path.GetFileName(destination);
        module.Write(destination);
    }

    /// <summary>Line endings only — approved files are committed and checked out on both Windows
    /// and Linux. Shared here so the three call sites (both approval-test projects and this
    /// class's own regression tests) cannot drift onto different rules.</summary>
    internal static string Normalise(string text) =>
        text.Replace("\r\n", "\n", StringComparison.Ordinal).TrimEnd();
}
