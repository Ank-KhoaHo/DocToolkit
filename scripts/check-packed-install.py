#!/usr/bin/env python3
"""Install the LOCALLY PACKED packages into a throwaway project and run one conversion per format.

WHY THIS EXISTS. Issue #383, filed after 0.36.0 shipped carrying `DocToolkit.dll` alone. Three
checks look like they cover that class of defect and none of them did:

  * the test suite builds against PROJECT REFERENCES, not the artifact that ships;
  * `check-package-contents.py` proves the nupkg CONTAINS every assembly - it executes nothing;
  * `install-smoke.yml` compiles and runs a real consumer app, but on `workflow_run` AFTER
    Release, so a defect it catches is already permanent. A nuget.org version can be unlisted,
    never deleted or replaced.

So the only check that executed a consumer ran after the irreversible step. That gap is the
limitation `check-package-contents.py`'s own docstring records: it proves the count, not that a
consumer can compile.

WHAT THIS CATCHES BEYOND THE ASSEMBLY COUNT, each a real hole rather than a hypothetical:

  * a MISSING TRANSITIVE DEPENDENCY. `CLAUDE.md` records the near-miss: dropping
    `PrivateAssets="all"` produced a nuspec with NO <dependency> entries while still shipping the
    assemblies that need them - "a package that installs cleanly, restores cleanly, and throws
    TypeLoadException on the consumer's machine".
  * AN ASSEMBLY THAT PACKS BUT FAILS TO LOAD. Invisible to a zip listing.
  * A WRONG TARGET FRAMEWORK. `lib/net8.0/` present but not consumable.

TWO RULES INHERITED FROM `install-smoke.yml`, both of which it paid for:

  1. **The consumer project is created OUTSIDE the repository.** A clean-machine test that picks up
     `global.json`, `Directory.Build.props` or a `NuGet.config` is not a clean-machine test.
  2. **It ASSERTS THE VERSION IT INSTALLED.** `install-smoke.yml` verified the PREVIOUS release
     eight times out of eight before anyone noticed, and the tell was that a green run never said
     which version it tested. Here the same trap is live in a different form: the global package
     cache, or nuget.org itself, can satisfy `dotnet add package` and the run then passes against
     something other than the artifact just built. `NUGET_PACKAGES` is pointed at an empty
     directory and the resolved versions are read back and compared.

    python scripts/check-packed-install.py --artifacts artifacts
    python scripts/check-packed-install.py --self-test
"""

import argparse
import glob
import io
import json
import os
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile

REPO = pathlib.Path(__file__).resolve().parent.parent

CORE = "Ank.DocToolkit"
EXTENSIONS = "Ank.DocToolkit.Extensions.DependencyInjection"

# `Ank.DocToolkit.1.2.3.nupkg` and `Ank.DocToolkit.1.2.3-beta.1.nupkg`. The id is anchored so the
# extensions package cannot be read as the core one with a very odd version - the same
# longest-match trap check-package-contents.py hit.
NUPKG = re.compile(r"^(?P<id>Ank\.DocToolkit(?:\.Extensions\.DependencyInjection)?)"
                   r"\.(?P<version>\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?)\.nupkg$")

# One capability per shipped format, plus the DI container. Modelled on install-smoke.yml's
# program: each line asserts an OUTCOME, never merely that a call returned.
CONSUMER_PROGRAM = '''using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;

static void Check(bool ok, string what)
{
    Console.WriteLine($"{(ok ? "ok  " : "FAIL")} {what}");
    if (!ok) Environment.Exit(1);
}

// DOCX, from DocToolkit.Docx - one of the assemblies 0.36.0 shipped without.
byte[] docx = DocxEditor.Create(new[] { DocxBlock.Paragraph("Hello from a packed install.") });
Check(DocxEditor.ExtractText(docx).Contains("packed install"), "DocxEditor round-trip");

// DocToolkit.Primitives, the assembly whose absence produced CS0012 on PageSetup.
byte[] a4 = DocxEditor.Create(new[] { DocxBlock.Paragraph("sized") }, PageSetup.A4);
Check(a4.Length > 0, "PageSetup reaches the API");

// DocToolkit.Xlsx and DocToolkit.Pptx.
byte[] xlsx = WorkbookEditor.Create(new[] { XlsxSheet.Named("S", new[] { new object?[] { "n", 41 } }) });
Check(WorkbookEditor.ReadCell(xlsx, "S", "B1") == "41", "WorkbookEditor round-trip");

byte[] pptx = PresentationEditor.Create(new[] { PptxSlide.Titled("Title", "bullet") });
Check(PresentationEditor.SlideCount(pptx) == 1, "PresentationEditor round-trip");

// The PDF renderer, and DocToolkit.Pdf reading it back.
byte[] pdf = DocxToPdfConverter.Convert(docx);
Check(pdf.Length > 1000 && pdf[0] == (byte)'%', "DocxToPdfConverter produces a PDF");
Check(PdfEditor.PageCount(pdf) >= 1, "PdfEditor reads it back");

// DocToolkit.Html, through the converter that composes it. ConvertAsync, not Convert: the HTML
// converters are async-only, because HtmlToOpenXml's own entry point is. Top-level await is what a
// consumer would write here too.
byte[] fromHtml = await HtmlToDocxConverter.ConvertAsync("<p>markup</p>");
Check(DocxEditor.ExtractText(fromHtml).Contains("markup"), "HtmlToDocxConverter round-trip");

// DocxReview, the newest surface - so a newly added assembly is exercised too.
Check(DocxReview.Inspect(docx).Comments.Count == 0, "DocxReview.Inspect");

// The extensions package, through a real container.
var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();
Check(provider.GetRequiredService<IWorkbookEditor>().ReadCell(xlsx, "S", "B1") == "41",
      "IWorkbookEditor resolved from the container");
Check(provider.GetRequiredService<IDocxReview>().Inspect(docx).Revisions.Count == 0,
      "IDocxReview resolved from the container");

Console.WriteLine("A packed install of both packages works.");
'''


def run(args, cwd, env=None, check=True):
    result = subprocess.run(args, cwd=cwd, env=env, capture_output=True, text=True)
    if check and result.returncode != 0:
        print(result.stdout[-4000:])
        print(result.stderr[-4000:], file=sys.stderr)
        sys.exit(f"::error::`{' '.join(args)}` failed with exit {result.returncode}.")
    return result


def discover(artifacts):
    """{package id: version} from the nupkg filenames in `artifacts`."""
    found = {}
    for path in sorted(glob.glob(str(pathlib.Path(artifacts) / "*.nupkg"))):
        m = NUPKG.match(os.path.basename(path))
        if m:
            found[m.group("id")] = m.group("version")

    missing = [p for p in (CORE, EXTENSIONS) if p not in found]
    if missing:
        sys.exit(f"::error::no .nupkg found in {artifacts} for {', '.join(missing)}. This check "
                 "must not pass while unable to find the artifact it exists to install.")

    if found[CORE] != found[EXTENSIONS]:
        sys.exit(f"::error::the two packages carry different versions - core {found[CORE]}, "
                 f"extensions {found[EXTENSIONS]}. They release together from one tag, so a "
                 "mismatch here is a packaging fault, not something to install around.")
    return found


def installed_versions(project):
    """{package id: resolved version} as `dotnet list package` reports it."""
    listing = run(["dotnet", "list", "package"], cwd=project).stdout
    versions = {}
    for line in listing.splitlines():
        parts = line.split()
        # A top-level entry is `> Id requested resolved`; the resolved version is always last.
        if len(parts) >= 3 and parts[0] == ">" and parts[1] in (CORE, EXTENSIONS):
            versions[parts[1]] = parts[-1]
    return versions


def verify(artifacts):
    packed = discover(artifacts)
    expected = packed[CORE]
    feed = str(pathlib.Path(artifacts).resolve())

    print(f"artifacts: {feed}")
    print(f"packed version: {expected}")

    # OUTSIDE the repository, so global.json, Directory.Build.props and any NuGet.config are out
    # of reach. A clean-machine test that inherits the repo's build configuration is not one.
    workspace = tempfile.mkdtemp(prefix="doctoolkit-packed-")
    try:
        env = dict(os.environ)
        # An empty cache, so neither a warm global cache nor nuget.org can satisfy the restore and
        # let this pass against something other than what was just packed.
        env["NUGET_PACKAGES"] = os.path.join(workspace, "empty-cache")
        env["DOTNET_NOLOGO"] = "1"
        env["DOTNET_CLI_TELEMETRY_OPTOUT"] = "1"

        run(["dotnet", "new", "console", "-o", "app"], cwd=workspace, env=env)
        project = os.path.join(workspace, "app")

        # A project-local NuGet.config rather than `dotnet nuget add source`, which writes to the
        # USER-LEVEL config: that mutates the machine, persists after the run, and made the second
        # invocation fail with "the name specified has already been added". A check must not leave
        # the environment changed.
        #
        # This is not the NuGet.config the docstring warns about inheriting - that one would be the
        # REPOSITORY's, carrying its own settings. This one is the consumer's own, declaring where
        # its packages come from, which is what a consumer writes. `<clear/>` makes the source list
        # deterministic rather than whatever the runner happens to have configured.
        config_lines = [
            '<?xml version="1.0" encoding="utf-8"?>',
            "<configuration>",
            "  <packageSources>",
            "    <clear />",
            f'    <add key="packed" value="{feed}" />',
            '    <add key="nuget.org" value="https://api.nuget.org/v3/index.json" />',
            "  </packageSources>",
            "</configuration>",
            "",
        ]
        io.open(os.path.join(project, "NuGet.config"), "w", encoding="utf-8",
                newline=chr(10)).write(chr(10).join(config_lines))

        # NO `--source` here. Restricting restore to the local feed makes the TRANSITIVE
        # dependencies unresolvable - OfficeIMO, ClosedXML, PdfPig and thirteen others live on
        # nuget.org - and a consumer's restore reads both sources too. Measured: with `--source`
        # the restore fails with eighteen NU1101s and never reaches the compile this check exists
        # for.
        for package in (EXTENSIONS, CORE):
            run(["dotnet", "add", "package", package, "--version", expected], cwd=project, env=env)
        run(["dotnet", "add", "package", "Microsoft.Extensions.DependencyInjection"],
            cwd=project, env=env)

        # THE LOAD-BEARING ASSERTION. Everything above can succeed against the wrong version -
        # which is exactly what install-smoke.yml did eight releases running.
        got = installed_versions(project)
        failures = [f"{p}: expected {expected}, installed {got.get(p) or 'none'}"
                    for p in (CORE, EXTENSIONS) if got.get(p) != expected]
        if failures:
            for f in failures:
                print(f"::error::{f}")
            sys.exit("::error::the consumer did not install the packages this run just packed, so "
                     "whatever it exercised is not the artifact under test.")
        print(f"installed: {CORE} {got[CORE]}, {EXTENSIONS} {got[EXTENSIONS]}")

        # AND THE VERSION MATCHING IS NOT ENOUGH, because nuget.org may already carry that same
        # version - so a passing run could be exercising the PUBLISHED package while claiming to
        # test the packed one. That is install-smoke.yml's job, not this one, and the two must not
        # be confused: this check exists to run BEFORE the publish.
        #
        # So compare bytes. The restored package sits under the empty NUGET_PACKAGES cache; if it
        # is not byte-identical to what was just packed, this run is not testing the artifact.
        cache = pathlib.Path(env["NUGET_PACKAGES"])
        for package in (CORE, EXTENSIONS):
            restored = cache / package.lower() / expected / f"{package.lower()}.{expected}.nupkg"
            local = pathlib.Path(feed) / f"{package}.{expected}.nupkg"
            if not restored.exists():
                sys.exit(f"::error::{package} {expected} is not in the restore cache at "
                         f"{restored}; cannot confirm which artifact was installed.")
            if restored.read_bytes() != local.read_bytes():
                sys.exit(f"::error::the restored {package} {expected} is NOT byte-identical to the "
                         "one just packed, so something else served it - probably nuget.org. This "
                         "check must exercise the artifact about to be published, not the one "
                         "already published.")
        print("both restored packages are byte-identical to the packed artifacts")

        io.open(os.path.join(project, "Program.cs"), "w", encoding="utf-8", newline="\n").write(
            CONSUMER_PROGRAM)

        result = run(["dotnet", "run", "-c", "Release"], cwd=project, env=env, check=False)
        print(result.stdout.strip())
        if result.returncode != 0:
            print(result.stderr[-4000:], file=sys.stderr)
            sys.exit("::error::a consumer of the packed packages could not compile or run. This is "
                     "the 0.36.0 failure mode: the nupkg installs and the code does not work.")

        # The premise a consumer is buying, checked against the graph THEY get.
        native = [str(p) for p in pathlib.Path(project).rglob("*")
                  if "bin" in p.parts and p.suffix in (".so", ".dylib")]
        if native:
            for n in native:
                print(f"  {n}")
            sys.exit("::error::the packed packages put native binaries into a consumer's output.")

        print()
        print(f"A clean consumer installs and runs {expected} from the packed artifacts, with no "
              "native binaries.")
        return 0
    finally:
        shutil.rmtree(workspace, ignore_errors=True)


def self_test():
    """Controls for the parts that can be exercised without a dotnet round-trip."""
    bad = 0
    cases = [
        ("Ank.DocToolkit.1.2.3.nupkg", CORE, "1.2.3"),
        ("Ank.DocToolkit.0.37.0.nupkg", CORE, "0.37.0"),
        ("Ank.DocToolkit.1.2.3-beta.1.nupkg", CORE, "1.2.3-beta.1"),
        # The longest-match trap check-package-contents.py actually hit: this must NOT read as core.
        ("Ank.DocToolkit.Extensions.DependencyInjection.1.2.3.nupkg", EXTENSIONS, "1.2.3"),
        ("Ank.DocToolkit.1.2.3.snupkg", None, None),
        ("Something.Else.1.0.0.nupkg", None, None),
    ]
    for name, want_id, want_version in cases:
        m = NUPKG.match(name)
        got_id = m.group("id") if m else None
        got_version = m.group("version") if m else None
        ok = (got_id, got_version) == (want_id, want_version)
        print(f"  {'ok  ' if ok else 'FAIL'} {name} -> {got_id} {got_version}")
        bad += 0 if ok else 1

    with tempfile.TemporaryDirectory() as empty:
        try:
            discover(empty)
            print("  FAIL an empty artifacts directory passed")
            bad += 1
        except SystemExit:
            print("  ok   an empty artifacts directory is refused")

    print()
    if bad:
        print(f"::error::self-test failed {bad} case(s)")
        return 1
    print("self-test passed")
    return 0


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--artifacts", default="artifacts")
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()

    if args.self_test:
        return self_test()
    return verify(args.artifacts)


if __name__ == "__main__":
    sys.exit(main())
