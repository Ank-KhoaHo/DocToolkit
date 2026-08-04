# Samples split per capability — design

## Why

`samples/ConsoleSample/Program.cs` runs eight numbered sections in one file: HTML→DOCX, HTML→PDF,
DOCX→PDF, template fill and text extraction, repeating table rows, image placeholders, XLSX
create/read/edit, and PPTX read.

That worked when there were three sections. At eight it fails the job a sample exists to do. Someone
who wants to know "how do I fill a Word template?" has to read past four unrelated conversions to
find it, and everything shares one scope because the file uses top-level statements — so variable
names collide across capabilities that have nothing to do with each other. The file is a
demonstration of the whole package, when what a reader wants is a demonstration of one thing.

## Structure

```
samples/
  README.md                  index: which sample answers which question
  HtmlConversion/            HTML→DOCX, HTML→PDF, DOCX→PDF        (was sections 1, 2, 3)
  DocxTemplating/            fill + extract, repeating table rows  (was sections 4, 5)
  DocxImages/                image placeholder                     (was section 6)
  Spreadsheets/              create/read/edit, plus bulk reading    (was section 7, extended)
  Presentations/             PPTX read                             (was section 8)
  MinimalApi/                unchanged apart from the folder name
```

`ConsoleSample` is **retired**, not kept alongside. Keeping it would put every capability in two
places, and duplicated demonstrations drift — the repo has been bitten by that before.

**Names drop the `Sample` suffix.** `samples/HtmlConversionSample` says "sample" twice. For
consistency that also renames `MinimalApiSample` to `MinimalApi`.

### Why three sections group rather than split

Splitting strictly one-per-section would lose the thing each group teaches.

**`HtmlConversion` keeps all three conversions together** because HTML→PDF *pivots through DOCX
internally*. Seeing the three side by side is what makes that visible; in three folders it is
invisible.

**`DocxTemplating` keeps fill and repeating rows together** because the invoice example only teaches
anything as a pair: `FillRows` must run before `ReplaceText`, since expanding clones the template row
and any scalar substituted first is duplicated into every line. Two folders would separate the
ordering rule from the code it constrains.

### Why `DocxImages` stops borrowing `sample.pptx`

Today the image bytes are extracted from `sample.pptx`, which `ConsoleSample` already ships for its
PPTX section — a deliberate dodge so no picture is committed whose only job is to be a picture.

Once `Presentations` is a separate project that stops being clever and starts being confusing: an
image sample that needs a PowerPoint file, for reasons visible nowhere in the folder. `DocxImages`
generates a small PNG in code instead. Only `Presentations` links `sample.pptx`.

### `Spreadsheets` absorbs deferred work

`SheetNames` and `ReadSheet` shipped in v0.6.0. They could not be demonstrated when they were
written, because samples reference the published package and the methods had not been released. This
is where that lands — no separate change.

## Shared configuration

`ConsoleSample.csproj` carries a fifteen-line comment explaining why `Version="*"` is deliberate and
why a version *floor* is actively wrong (NuGet resolves a minimum-version range to the **lowest**
satisfying version, so a floor pins the sample to an old release while newer ones go unexercised).

Copied into five files, that comment rots in four of them.

`samples/Directory.Build.props` therefore carries the four properties every sample shares —
`TargetFramework` (`net8.0`, which both existing samples already use), `ImplicitUsings`, `Nullable`,
`IsPackable` — and the `Version="*"` rationale as a comment. The prose version lives in
`samples/README.md`; each `.csproj` keeps a one-line pointer to it rather than a copy.

**`Directory.Build.props` must NOT declare the `Ank.DocToolkit` package reference.** `MinimalApi`
deliberately references *only* `Ank.DocToolkit.Extensions.DependencyInjection`, which proves the
extensions package drags the core package in transitively the way a consumer's restore would. A props
file adding an explicit core reference to every sample would silently destroy that proof — the
project would still build if the transitive dependency broke. Each sample declares its own
`PackageReference`.

`MinimalApi` also uses `Microsoft.NET.Sdk.Web` rather than `Microsoft.NET.Sdk`. That is the `Sdk`
attribute on the `Project` element, which `Directory.Build.props` does not touch, so sharing
properties across both is safe.

This adds no lockfile: `samples/` deliberately carries none, and `Directory.Build.props` does not
change that.

## READMEs

**All six folders** get a `README.md`, `MinimalApi` included — its documentation currently lives in
`samples/README.md` and moves down with the rest.

Each covers four things: what the sample shows, the one command to run it, what it prints, and
**the non-obvious rule, if it has one**.

That last part is the point. `DocxTemplating`'s README exists to say "`FillRows` before
`ReplaceText`, or scalars are duplicated into every line" — a rule currently buried in
`samples/README.md`, where nobody reading the templating code will find it.

`samples/README.md` shrinks to an index: a table of folder → the question it answers, plus the
paragraph explaining why every sample references the **published** package rather than this source,
and what that buys (they exercise the artifact a consumer restores, so a capability merged but not
released cannot appear here until it is).

## What else changes

Three things outside `samples/` break unless they change in the same commit:

- **`README.md`** runs `dotnet run --project samples/ConsoleSample` in its build-and-test block.
  That path stops existing.
- **`CLAUDE.md`**'s *Samples and docs site* section names `ConsoleSample` and documents the
  `sample.pptx` borrowing trick. The trick is being removed, so the paragraph explaining it goes
  too — not just the name.
- **`DocToolkit.sln`** needs five project entries added and two removed.
- **`MinimalApiSample.csproj`**'s own comment says *"Version=`*` tracks the newest published release —
  see ConsoleSample.csproj"*. That file is being deleted, so the pointer dangles. It becomes a
  pointer to `samples/README.md`.

Both renames break external links, and this repo is public with the READMEs as its only navigation.
The whole change therefore ships as **one pull request**, so no intermediate state has dangling
paths.

## Verification

- `dotnet build DocToolkit.sln -c Release -warnaserror --no-incremental` at **0 warnings**.
- **Every sample run once, not merely built.** CI only builds them, so a sample that compiles and
  throws at runtime would ship green. This is the check that matters most and the one CI cannot do.
- Each README's stated command copy-pasted and confirmed to work.
- No `samples/**` lockfile appears.

## Success criteria

- A reader wanting one capability opens one folder and reads one short file.
- No capability is demonstrated in two places.
- The `Version="*"` reasoning exists once.
- `README.md` and `CLAUDE.md` contain no path that no longer exists.
- Every sample runs.

## Out of scope

`IWorkbookEditor` parity in the DI extensions package. It was blocked by the same cause — the DI
project references the published `Ank.DocToolkit` — and is unblocked by v0.6.0, but it is a change to
a shipping package, not to samples, and belongs in its own change.
