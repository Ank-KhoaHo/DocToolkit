# Docs/samples site — design

## Why

Ank.DocToolkit and Ank.DocToolkit.Extensions.DependencyInjection are live on nuget.org, but an
adopter's only reference material is the two READMEs. A generated API-reference site (from XML
doc comments both projects already produce — `GenerateDocumentationFile` is already `true` in
both csprojs) and runnable sample projects give adopters something more concrete than reading
source, and cost little given the doc comments already exist.

## Scope

Two related but separable components, combined here because the site's landing/getting-started
content wants to point at real sample code:

1. Sample projects (`samples/`), added to `DocToolkit.sln`, built by the existing CI.
2. A DocFX-generated API-reference site (`docfx/`), published to GitHub Pages on release.

Not in scope: hand-written conceptual/tutorial pages beyond a landing page adapted from the
READMEs; a documentation *versioning* scheme (multiple doc versions for multiple package
versions) — the site always reflects the latest successful release, matching how the READMEs
already work.

## Samples

```
samples/
├── ConsoleSample/
│   ├── ConsoleSample.csproj
│   └── Program.cs
└── MinimalApiSample/
    ├── MinimalApiSample.csproj
    └── Program.cs
```

**`ConsoleSample`** exercises all five core capabilities in one `Program.cs`, mirroring the
top-level README's existing usage snippet: HTML→DOCX, HTML→PDF, DOCX edit + extract, XLSX
create/read/edit, PPTX read/edit. Plain console output describing each step, so it's readable
top-to-bottom without a debugger.

**`MinimalApiSample`** is an ASP.NET Core minimal API: `services.AddDocToolkit()` at startup, one
`MapPost`/`MapGet` endpoint per injected interface (`IHtmlToDocxConverter`, `IHtmlToPdfConverter`,
`IDocxEditor`, `IWorkbookEditor`, `IPresentationEditor`, `IDocxToPdfConverter`), each a thin
wrapper returning the resulting file. Demonstrates the DI registration path the console sample
can't.

Both `.csproj` files reference the packages via:

```xml
<PackageReference Include="Ank.DocToolkit" Version="[0.2.1, )" />
<PackageReference Include="Ank.DocToolkit.Extensions.DependencyInjection" Version="[0.2.1, )" />
```

An open floor range, not a pin — matches how the extensions package already references core.
Both projects are `<IsPackable>false</IsPackable>` (never shipped as packages themselves) and
added to `DocToolkit.sln` via `dotnet sln add`. Because they join the solution, the existing
`ci.yml` build step (`dotnet build DocToolkit.sln ...`) picks them up with **no CI file changes
required** — a breaking change in a future release fails the very next sample build instead of
silently drifting unnoticed.

## Docs site

New top-level `docfx/` folder — deliberately separate from the existing `docs/` folder, which
holds dated planning/spec markdown (this project's own history), not published site source. The
two serve different readers (maintainer vs. adopter) and different lifecycles (append-only history
vs. always-current site).

Scaffolded via `docfx init` (not hand-written from scratch — this generates a working default
`docfx.json` and template to edit, which is the standard workflow rather than authoring the config
from memory). The generated `docfx.json`'s `metadata` section is pointed at both csproj files:

```
src/DocToolkit/DocToolkit.csproj
src/DocToolkit.Extensions.DependencyInjection/DocToolkit.Extensions.DependencyInjection.csproj
```

so the API reference covers both packages. The landing page (`docfx/index.md`) is adapted from
the top-level README's introduction and usage sections rather than written fresh — the content
already exists and is already accurate.

Output directory is DocFX's default, `docfx/_site/`, gitignored (generated, not committed).

## Publish pipeline

New `.github/workflows/docs.yml`:

```yaml
name: Docs

on:
  workflow_run:
    workflows: ["Release"]
    types: [completed]

permissions:
  contents: read
  pages: write
  id-token: write

concurrency:
  group: pages
  cancel-in-progress: false

jobs:
  build-and-deploy:
    if: github.event.workflow_run.conclusion == 'success'
    runs-on: ubuntu-latest
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    steps:
      - uses: actions/checkout@v4
        with:
          ref: ${{ github.event.workflow_run.head_sha }}
      - uses: actions/setup-dotnet@v4
        with:
          dotnet-version: |
            8.0.x
            10.0.x
      - run: dotnet restore DocToolkit.sln
      - run: dotnet tool update -g docfx
      - run: docfx docfx/docfx.json
      - uses: actions/configure-pages@v5
      - uses: actions/upload-pages-artifact@v3
        with:
          path: docfx/_site
      - id: deployment
        uses: actions/deploy-pages@v4
```

**Triggered by `workflow_run` on `Release` completing, gated on `conclusion == 'success'`** — not
independently on the same tag push. This is deliberate: if `release.yml`'s guards catch something
(missing CHANGELOG entry, a banned dependency, broken tests) and refuse to publish, the docs site
must not publish either, describing a version that never actually shipped. `ref:
${{ github.event.workflow_run.head_sha }}` checks out the exact commit the release ran against, so
the generated API reference matches what's on nuget.org, not whatever `main` has moved to since.

**One manual step, once** (same shape as the nuget Trusted Publishing policy and the Codecov
token earlier): enable GitHub Pages in the repo's Settings → Pages, source "GitHub Actions." If
the first post-release run's deploy step fails, check that setting before assuming the workflow
is broken.

**Known gotcha, not a bug if hit:** `workflow_run` triggers off the workflow file as it exists on
the default branch at the time the triggering workflow (`release.yml`) completes. `docs.yml` must
already be merged to `main` *before* the next `v*` tag is pushed — it cannot retroactively fire
for a release whose triggering commit predates `docs.yml` existing on `main`. The first real test
of this pipeline is therefore the *next* release after this work merges, not the current one.

## Docs updates

- `README.md` gets a "Documentation" link once the site is live (URL will be
  `https://ank-khoaho.github.io/DocToolkit/` — GitHub Pages' default for a `github-actions`
  source deployment on this repo).
- `CLAUDE.md` gets a short section: the `samples/` convention (PackageReference, not
  ProjectReference — same reasoning as the extensions package), and a pointer to this design's
  publish-pipeline section so the `workflow_run` dependency on `release.yml` isn't mistaken for a
  redundant/independent trigger and "simplified" away later.

## Testing / validation

- Samples: `dotnet build DocToolkit.sln` must succeed with both new projects included; running
  `ConsoleSample` locally should produce the expected output files (docx/pdf/xlsx/pptx) without
  throwing.
- Docs site: run `docfx docfx/docfx.json` locally and confirm `docfx/_site/api/` contains pages
  for the public types in both projects (not an empty/near-empty API section — the actual
  discriminating check, since a misconfigured `metadata.src` glob would still "succeed" while
  generating nothing).
- Publish pipeline: after the next real `v*` tag release succeeds, confirm `docs.yml` actually
  triggered (via `gh run list --workflow=docs.yml`), watch it to completion, and load the deployed
  URL to confirm real content renders — not just that the job reports success.
