# Samples

Six runnable projects, each answering one question. Every one references the **published** NuGet
packages rather than this repo's source — the same restore an external consumer gets.

| Sample | Answers |
|---|---|
| [HtmlConversion](HtmlConversion/) | How do I turn HTML into a DOCX or a PDF? |
| [DocxTemplating](DocxTemplating/) | How do I fill a Word template, including one row per record? |
| [DocxImages](DocxImages/) | How do I drop a logo or signature into a placeholder? |
| [Spreadsheets](Spreadsheets/) | How do I create, edit and read an XLSX? |
| [Presentations](Presentations/) | How do I read text out of a PowerPoint file? |
| [MinimalApi](MinimalApi/) | How do I wire this into ASP.NET Core dependency injection? |

Each folder has its own `README.md` with the command to run it and the one thing about that
capability that is not obvious.

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
