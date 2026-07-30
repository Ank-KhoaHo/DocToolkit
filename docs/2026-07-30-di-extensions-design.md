# Design: Dependency-injection extensions for DocToolkit

**Date:** 2026-07-30
**Status:** Approved design — ready for implementation plan

## 1. Context and goal

`Ank.DocToolkit` 0.1.0 is published on nuget.org. Its public API is six static classes
(`HtmlToDocxConverter`, `DocxToPdfConverter`, `HtmlToPdfConverter`, `DocxEditor`,
`WorkbookEditor`, `PresentationEditor`), each `byte[]` in / `byte[]` out, stateless, safe to
call concurrently.

That shape is fine for scripts and simple console apps, but it's awkward for ASP.NET Core and
worker-service consumers, who expect to `services.AddX()` and inject an interface rather than
call a static method directly — the static shape can't be mocked via DI, can't have its
configuration (e.g. whether to fetch remote images) centralised, and doesn't compose with the
rest of a typical .NET host's service graph.

The goal is to add that shape **without touching what's already shipped.** 0.1.0 has real
consumers now (however few); nothing about its public surface should change.

## 2. Non-goals

- Not a rewrite of the conversion logic. Every interface method is a thin wrapper delegating to
  the existing static method.
- Not a replacement for the static API. Both remain fully supported, indefinitely.
- Not adding per-call configuration overrides in this pass. `AllowRemoteImageDownload` becomes a
  registration-time setting (see §5); a future per-call override is a separate, additive change
  if it turns out to be needed.

## 3. Package structure

A new package and project, alongside the existing one:

```
DocToolkit/
├── src/
│   ├── DocToolkit/                                    (existing, UNCHANGED)
│   └── DocToolkit.Extensions.DependencyInjection/      (new)
│       ├── DocToolkit.Extensions.DependencyInjection.csproj
│       ├── DocToolkitOptions.cs
│       ├── ServiceCollectionExtensions.cs
│       ├── IHtmlToDocxConverter.cs + HtmlToDocxConverterService.cs
│       ├── IDocxToPdfConverter.cs + DocxToPdfConverterService.cs
│       ├── IHtmlToPdfConverter.cs + HtmlToPdfConverterService.cs
│       ├── IDocxEditor.cs + DocxEditorService.cs
│       ├── IWorkbookEditor.cs + WorkbookEditorService.cs
│       └── IPresentationEditor.cs + PresentationEditorService.cs
└── tests/
    ├── DocToolkit.Tests/                                          (existing, UNCHANGED)
    └── DocToolkit.Extensions.DependencyInjection.Tests/            (new)
```

**Package id:** `Ank.DocToolkit.Extensions.DependencyInjection`, versioned independently of the
core package, starting at `0.1.0`.

**Dependencies:** `Ank.DocToolkit` (as a real `PackageReference` with a version floor, e.g.
`[0.1.0, )`, so a consumer installing the extensions package transitively gets the core one — not
a project reference masquerading as one), `Microsoft.Extensions.DependencyInjection.Abstractions`,
`Microsoft.Extensions.Options`. Both are MIT and already present indirectly in the dependency
closure via `Microsoft.Extensions.Logging.Abstractions`, so this doesn't introduce a new licence
to track.

**Target frameworks:** `net8.0;net10.0`, matching the core package exactly — a consumer must never
be able to install the extensions package into a project the core package doesn't support.

## 4. Interfaces

One interface per capability, mirroring the six static classes 1:1. Each method signature matches
its static counterpart exactly, with two exceptions: the `allowRemoteImageDownload` parameter is
dropped (moved to `DocToolkitOptions`, §5), and static-only convenience overloads
(`ConvertToFileAsync`, `ConvertFile`) are not carried over — a DI consumer receiving a `byte[]`
result is expected to write it wherever it needs to go itself.

```csharp
namespace DocToolkit.Extensions.DependencyInjection;

public interface IHtmlToDocxConverter
{
    Task<byte[]> ConvertAsync(string html, CancellationToken ct = default);
}

public interface IDocxToPdfConverter
{
    byte[] Convert(byte[] docx);
}

public interface IHtmlToPdfConverter
{
    Task<byte[]> ConvertAsync(string html, CancellationToken ct = default);
}

public interface IDocxEditor
{
    byte[] ReplaceText(byte[] docx, IReadOnlyDictionary<string, string> replacements);
    string ExtractText(byte[] docx);
}

public interface IWorkbookEditor
{
    byte[] Create(string sheetName, IEnumerable<IEnumerable<object?>> rows);
    string ReadCell(byte[] xlsx, string sheetName, string cellRef);
    byte[] SetCell(byte[] xlsx, string sheetName, string cellRef, object? value);
}

public interface IPresentationEditor
{
    int SlideCount(byte[] pptx);
    IReadOnlyList<string> ExtractText(byte[] pptx);
    byte[] ReplaceText(byte[] pptx, IReadOnlyDictionary<string, string> replacements);
}
```

Sync methods stay sync, async methods stay async — each mirrors whether the underlying static
method actually awaits anything. No method is wrapped in `Task.Run` to appear asynchronous.

Each `*Service` implementation is a direct pass-through: e.g.
`HtmlToDocxConverterService.ConvertAsync` calls
`HtmlToDocxConverter.ConvertAsync(html, options.Value.AllowRemoteImageDownload, ct)` and returns
the result. No logic is duplicated between the static class and its service wrapper.

## 5. Configuration

```csharp
namespace DocToolkit.Extensions.DependencyInjection;

public sealed class DocToolkitOptions
{
    /// <summary>
    /// When true, HTML-to-DOCX/PDF conversion downloads images referenced by absolute
    /// http/https URLs. Issues outbound network requests - do not enable in an air-gapped
    /// environment. Default: false.
    /// </summary>
    public bool AllowRemoteImageDownload { get; set; } = false;
}
```

Registered and consumed through the standard `IOptions<T>` pattern, so it composes with
configuration binding, named options, and options validation if a consumer wants those later.

## 6. Registration

```csharp
namespace DocToolkit.Extensions.DependencyInjection;

public static class ServiceCollectionExtensions
{
    public static IServiceCollection AddDocToolkit(
        this IServiceCollection services,
        Action<DocToolkitOptions>? configure = null)
    {
        services.AddOptions<DocToolkitOptions>();
        if (configure is not null) services.Configure(configure);

        services.AddSingleton<IHtmlToDocxConverter, HtmlToDocxConverterService>();
        services.AddSingleton<IDocxToPdfConverter, DocxToPdfConverterService>();
        services.AddSingleton<IHtmlToPdfConverter, HtmlToPdfConverterService>();
        services.AddSingleton<IDocxEditor, DocxEditorService>();
        services.AddSingleton<IWorkbookEditor, WorkbookEditorService>();
        services.AddSingleton<IPresentationEditor, PresentationEditorService>();

        return services;
    }
}
```

```csharp
// usage
services.AddDocToolkit();                                     // default: no remote images
services.AddDocToolkit(o => o.AllowRemoteImageDownload = true);
```

`Singleton` is correct, not a shortcut: the services hold no mutable state, same as the static
classes they wrap, and the existing `AirGapGuardTests`-style proof already establishes that the
underlying calls are safe under concurrent use.

## 7. Testing

New test project, one test class per interface plus one cross-cutting class:

- **Parity tests** (one per interface): given the same input, the DI-resolved service produces
  byte-for-byte (or exact string) identical output to calling the static method directly. This is
  the test that actually matters — it proves the wrapper adds nothing and loses nothing.
- **Registration tests**: `AddDocToolkit()` resolves all six interfaces from a built
  `ServiceProvider`; resolving twice returns the same instance (`Singleton`); `AddDocToolkit(configure)`
  makes the configured value observable via injected `IOptions<DocToolkitOptions>`.
- **Options-actually-takes-effect test**: reuse the loopback-listener technique from
  `AirGapGuardTests` — register with `AllowRemoteImageDownload = false` (confirm zero connections)
  and `= true` (confirm the listener is hit). A test that only checked the option value without
  proving it changes runtime behaviour would be exactly the kind of vacuous test the existing
  `AirGapGuardTests` design was built to avoid.

## 8. What does not change

- The core `Ank.DocToolkit` package: no source changes, no version bump required by this work.
- The design premise (permissive licences, no native binaries, Linux, no runtime network) and its
  three CI guards — the new project must pass all of them too, so CI is extended to build and
  test it, not exempted from the checks.
- The release workflow's shape — the new package gets tagged and published the same way, most
  likely as its own tag (`extensions-v0.1.0` or similar) so its version can move independently of
  core. The exact tagging scheme is an implementation-plan decision, not a design one.

## 9. Open questions carried into the implementation plan

- Exact CI changes needed to build/test/pack a second project (likely: extend the existing
  matrix rather than duplicate it).
- Whether the release workflow needs a second, parameterised job, or whether tagging conventions
  alone are sufficient to distinguish "release core" from "release extensions."
