# Samples

Every project here is runnable and answers one question. Each references the **published** NuGet
packages rather than this repo's source — the same restore an external consumer gets.

**The count is deliberately not written down.** This line said "Nine", then "Thirteen", and while it
said thirteen the table below listed **eleven** — `Container` and `LargeFileStreaming` were missing
from it entirely, so two samples were invisible to anyone reading this page. A number and a list are
two claims that drift apart. `ls -d samples/*/` is the answer, and the table below is now the list.

| Sample | Answers |
|---|---|
| [HtmlConversion](HtmlConversion/) | How do I turn HTML into a DOCX or a PDF, on the page size I want — and what does a failure look like? |
| [MarkdownConversion](MarkdownConversion/) | How do I turn Markdown into a DOCX or a PDF — and how do I find out what the conversion could not carry across? |
| [DocxTemplating](DocxTemplating/) | How do I fill a Word template, including one row per record — or build a document with no template at all? |
| [DocxImages](DocxImages/) | How do I drop a logo or signature into a placeholder, and what happens to an image the HTML only points at? |
| [Spreadsheets](Spreadsheets/) | How do I create, edit and read an XLSX — make it look like a report, add a chart or a pivot table, and hand one sheet to something that is not Excel? |
| [Presentations](Presentations/) | How do I read text out of a PowerPoint file, add a chart to a slide, or read a SmartArt diagram's text? |
| [Signatures](Signatures/) | How do I tell whether a document is signed, and what does a self-signed certificate or a tampered document actually look like when I validate it? |
| [MinimalApi](MinimalApi/) | How do I wire this into ASP.NET Core dependency injection? |
| [RazorPdf](RazorPdf/) | How do I turn a Razor view I already maintain into a PDF or a Word document? |
| [WorkerService](WorkerService/) | How do I generate documents from a background job? |
| [PdfUtilities](PdfUtilities/) | How do I count, merge, split or label a PDF I already have? |
| [Telemetry](Telemetry/) | How do I find out whether my remote images are actually arriving, and why one didn't? |
| [Protection](Protection/) | How do I put a password on a document — and why is my "protected" file still opening for everyone? |
| [LegacyDoc](LegacyDoc/) | I have a share drive full of Word 97 `.doc` files. How do I read them, and turn them into `.docx`? |
| [Container](Container/) | Does this really work in a container with no fonts, no ICU and no native libraries — on `runtime-deps`, the smallest official .NET image? |
| [LargeFileStreaming](LargeFileStreaming/) | How do I convert from a forward-only, non-seekable source — an HTTP request body or a pipe? (The `Stream` overloads are for that, **not** for saving memory.) |

Each folder has its own `README.md` with the command to run it and the one thing about that
capability that is not obvious.

## The `#region` markers are load-bearing

These files are also the source of every code block in the
[guides](https://ank-khoaho.github.io/DocToolkit/guides/getting-started.html). The guides contain
no code of their own — each block is a `#region` pulled out of a project here, so a documented
snippet cannot drift from the API: it *is* the sample, and the sample stops compiling.

**Renaming or removing a region breaks a guide silently.** DocFX renders a reference to a region
that no longer exists as an empty code block and still exits 0 — measured, not assumed. That is
why `scripts/check-doc-snippets.py` runs in the `formatting` job and fails the build instead.

## Why they reference the published packages

Every sample uses `<PackageReference ... Version="*" />`, never a `ProjectReference`. That makes
them a canary: they exercise the artifact a consumer would actually restore, not whatever happens
to be on `main`.

`*` resolves to the newest published stable release, and that non-determinism is the point. A
version **floor** does the opposite — NuGet resolves a minimum-version range to the **lowest**
satisfying version, so `[0.2.1, )` once kept the samples building against 0.2.1 while three
releases shipped past them unexercised. Bumping the floor fixed it for exactly one release and
went stale on the next.

The trade-off: a capability that has been merged but not yet released **cannot appear here** until
it ships. That is not a gap to work around — it is the guarantee working.

Shared build settings live in `Directory.Build.props`. It deliberately declares no package
reference: `MinimalApi` references only the extensions package, which proves the core package
arrives transitively.

## `LargeFileStreaming`

What the `Stream` overloads are actually for, and what they are not.

It reads a 50,000-row workbook through a stream that **refuses to seek** - no `Length`, no
`Position`, `Seek` throws - which is what an HTTP request body or a socket gives you. A
`MemoryStream` would have hidden the point by quietly allowing a rewind.

It also prints the allocation total and says plainly that it is **not** lower than the `byte[]`
overload's. Memory is dominated by the OOXML object model, not by how the bytes arrive.

## `Container`

Every capability, inside a plain `dotnet/runtime` image - no SDK, no LibreOffice, no browser, no
fonts, and invariant globalization so not even ICU is present.

```bash
docker build -f samples/Container/Dockerfile -t doctoolkit-container .
docker run --rm --network none doctoolkit-container
```

`--network none` is worth using: the offline guarantee is then something you watched happen
rather than something you were told. CI builds and runs this image on every push, for the same
reason - a Dockerfile nobody builds is a claim, not a sample.

## `Telemetry`

Only `GuardedResourceLoader.FetchAsync` - the opt-in remote-image fetch - is instrumented, so
showing telemetry meaningfully means actually triggering a fetch. That needs somewhere to fetch
from, and reaching the real internet was rejected: it would make the sample non-deterministic,
occasionally slow, and a quiet argument that network access is normal here, which is exactly what
the rest of this library goes out of its way not to be.

Instead the sample brings its own loopback HTTP server - reachable only at `127.0.0.1`, only for
the life of the process, and only because the sample opts in with `AllowPrivateAddresses = true`
the same way `AirGapGuardTests` and `TelemetryTests` do. Nothing here weakens the "no socket opens
by default" guarantee: it demonstrates a real `ok` outcome against a server this process owns, and
two refused outcomes - `blocked_address` and `host_not_allowed` - that need no server at all,
because the guard refuses them before opening a connection. It also prints the request URL, which
carries a fake signed token in its query string, next to the span's recorded host, so "only the
host is ever recorded" is something you can check against the output rather than take on faith.
