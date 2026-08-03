# Design: Stream/async parity for the DI extensions

**Date:** 2026-08-03
**Status:** Approved design — ready for implementation plan

## 1. Context and goal

`Ank.DocToolkit.Extensions.DependencyInjection` wraps the six static classes in `Ank.DocToolkit`
behind six injectable interfaces (`IHtmlToDocxConverter`, `IDocxToPdfConverter`,
`IHtmlToPdfConverter`, `IDocxEditor`, `IWorkbookEditor`, `IPresentationEditor`), per
[2026-07-30-di-extensions-design.md](2026-07-30-di-extensions-design.md). It currently references
`Ank.DocToolkit` at `Version="[0.1.0, )"` — a floor, not the package's current release (both
packages have since moved past 0.1.0; see `CHANGELOG.md`).

Since that design was approved, the core `Ank.DocToolkit` static classes have grown a `Stream`-in
/ `Stream`-out async overload for nearly every operation — `ExtractTextAsync`, `ReplaceTextAsync`,
`SlideCountAsync`, `CreateAsync`, `ReadCellAsync`, `SetCellAsync`, and
`DocxToPdfConverter.ConvertAsync(Stream, Stream, ct)` — specifically so a caller (an ASP.NET Core
action, a worker service) can read from and write straight to a request body, a response body, or
a file, without buffering the whole document into a `byte[]`. None of those overloads were carried
into the DI interfaces; the DI package only ever wrapped the original `byte[]`-in/`byte[]`-out
surface. That's the gap this design closes: the DI package's stated goal — serve ASP.NET
Core/worker-service consumers — is undercut by the one API shape those consumers most want
(streaming into/out of a request or response body) being unavailable through DI.

## 2. Non-goals

- Not adding the file-path convenience methods (`ConvertToFileAsync`, `ConvertFile`). The original
  design excluded these deliberately — a DI consumer already holds a `byte[]`/`Stream` and can
  write it wherever it needs to — and that reasoning still holds. Confirmed with the repo owner.
- Not adding a per-call override for `AllowRemoteImageDownload`. It stays a registration-time
  `DocToolkitOptions` setting, per the original design's non-goals; the new `Stream`-destination
  overloads on `IHtmlToDocxConverter`/`IHtmlToPdfConverter` thread the same options value the
  existing `byte[]` overload already uses.
- Not a rewrite: every new interface method is a one-line delegate to an already-existing core
  static method, identical in spirit to every method the DI package already has.
- Not a core (`Ank.DocToolkit`) change. Every method being wrapped already exists there — as of
  the published `0.2.0` release, which added the `Stream` overloads (see CHANGELOG). The DI
  project's own `PackageReference` floor does need to move to `[0.2.0, )` for that surface to be
  visible to compile against; see §6.

## 3. Interface changes

One new method group per interface, matching its core static counterpart's signature exactly
(same convention the existing methods already follow: no `allowRemoteImageDownload` parameter,
no file-path methods):

```csharp
public interface IDocxEditor
{
    // existing: ReplaceText(byte[], ...), ExtractText(byte[]), ExtractText(byte[], bool)
    Task ReplaceTextAsync(
        Stream source, IReadOnlyDictionary<string, string> replacements, Stream destination,
        CancellationToken ct = default);
    Task<string> ExtractTextAsync(Stream source, CancellationToken ct = default);
    Task<string> ExtractTextAsync(Stream source, bool includeHeadersAndFooters, CancellationToken ct = default);
}

public interface IPresentationEditor
{
    // existing: SlideCount(byte[]), ExtractText(byte[]), ReplaceText(byte[], ...)
    Task<int> SlideCountAsync(Stream source, CancellationToken ct = default);
    Task<IReadOnlyList<string>> ExtractTextAsync(Stream source, CancellationToken ct = default);
    Task ReplaceTextAsync(
        Stream source, IReadOnlyDictionary<string, string> replacements, Stream destination,
        CancellationToken ct = default);
}

public interface IWorkbookEditor
{
    // existing: Create(string, rows), ReadCell(byte[], ...), SetCell(byte[], ...)
    Task CreateAsync(
        string sheetName, IEnumerable<IEnumerable<object?>> rows, Stream destination,
        CancellationToken ct = default);
    Task<string> ReadCellAsync(
        Stream source, string sheetName, string cellRef, CancellationToken ct = default);
    Task SetCellAsync(
        Stream source, string sheetName, string cellRef, object? value, Stream destination,
        CancellationToken ct = default);
}

public interface IHtmlToDocxConverter
{
    // existing: ConvertAsync(string, ct) -> byte[]
    Task ConvertAsync(string html, Stream destination, CancellationToken ct = default);
}

public interface IHtmlToPdfConverter
{
    // existing: ConvertAsync(string, ct) -> byte[]
    Task ConvertAsync(string html, Stream destination, CancellationToken ct = default);
}

public interface IDocxToPdfConverter
{
    // existing: Convert(byte[]) -> byte[]
    Task ConvertAsync(Stream source, Stream destination, CancellationToken ct = default);
}
```

Each `*Service` implementation is a direct pass-through, exactly like every existing method:

```csharp
// DocxEditorService
public Task ReplaceTextAsync(
    Stream source, IReadOnlyDictionary<string, string> replacements, Stream destination,
    CancellationToken ct = default)
    => DocToolkit.DocxEditor.ReplaceTextAsync(source, replacements, destination, ct);
```

```csharp
// HtmlToDocxConverterService — threads the registration-time option, same as the byte[] overload
public Task ConvertAsync(string html, Stream destination, CancellationToken ct = default)
    => DocToolkit.HtmlToDocxConverter.ConvertAsync(html, _options.AllowRemoteImageDownload, destination, ct);
```

No logic is duplicated between the static class and the service wrapper — same rule as the
original design.

**Compatibility note:** adding members to a shipped public interface is source/binary-breaking for
any external type that implements `IDocxEditor` etc. directly (mocking frameworks such as
Moq/NSubstitute are unaffected, since they generate implementations dynamically). These interfaces
are documented as wrapping `internal sealed` services and the package is pre-1.0 (`0.2.2`, heading
to `0.3.0`), so this is treated as an acceptable, deliberate additive change rather than a reason to
introduce versioned interfaces or default-interface-method shims.

## 4. Testing

Extend the existing per-interface parity-test files
(`DocxEditorServiceTests.cs`, `PresentationEditorServiceTests.cs`, `WorkbookEditorServiceTests.cs`,
`HtmlToDocxConverterServiceTests.cs`, `HtmlToPdfConverterServiceTests.cs`,
`DocxToPdfConverterServiceTests.cs`) — no new test file. One parity test per new method: given the
same input, the DI-resolved service's output (bytes written to the destination stream, or the
returned string/int/list) is identical to calling the core static method directly.

Plain `MemoryStream` is sufficient for these. `DocToolkit.Tests` has an elaborate set of stream
doubles (`ForwardOnlySource`, `ForwardOnlySink`, `TrackingStream`, ...) that prove the *core*
methods genuinely stream — read to end without seeking, write without reading back, honor
cancellation mid-read. The DI test project has no reference to `DocToolkit.Tests` and doesn't need
one: a parity test here is only proving "the wrapper delegates its arguments and return value
correctly," which a `MemoryStream` proves just as well as a forward-only one. Re-proving the
streaming contract itself would be duplicate coverage of what `DocToolkit.Tests` already owns.

## 5. Documentation

`src/DocToolkit.Extensions.DependencyInjection/README.md` already states the interfaces "mirror
[`Ank.DocToolkit`]'s static API one-for-one" — true after this change, false before it. Add one
short example showing a `Stream`-destination call (e.g. writing a converted PDF straight to an
ASP.NET Core response body) alongside the existing `byte[]`-returning example.

## 6. What does not change — and the one thing that does

- The core `Ank.DocToolkit` package: no source changes, no version bump required — the `Stream`
  overloads this design wraps already shipped in the published `0.2.0` release.
- The registration shape in `ServiceCollectionExtensions.AddDocToolkit` — same six `TryAddSingleton`
  calls, same `DocToolkitOptions`.
- The file-path convenience methods and the per-call `allowRemoteImageDownload` override stay out
  of scope, per §2.

**One dependency change is required:** `DocToolkit.Extensions.DependencyInjection.csproj`'s
`PackageReference Include="Ank.DocToolkit"` must move its version floor from `[0.1.0, )` to
`[0.2.0, )`. NuGet resolves a minimum-version range to the *lowest* satisfying version, not the
latest, so the existing `[0.1.0, )` floor pins the DI package to a core release that predates the
`Stream` overloads — restore silently succeeds against 0.1.0 and the new code fails to compile
with "does not contain a definition for ...Async". The floor must name the version that actually
introduced the API being wrapped. (Discovered during Task 1 of the implementation plan; a
`ProjectReference` is not an acceptable substitute — it would ship a package with no declared
`Ank.DocToolkit` dependency, breaking every consumer who installs the extensions package alone.)

## 7. Rollout

Conventional-commit type `feat` on the implementing commit(s), so release-please's existing
changelog/versioning pipeline for the extensions package (per
[2026-07-31-changelog-design.md](2026-07-31-changelog-design.md)) picks this up as a minor version
bump without any manual version editing.
