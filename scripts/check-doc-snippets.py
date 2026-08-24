#!/usr/bin/env python3
"""Fail when a guide's code snippet points at a region that does not exist.

The guides in docfx/guides/ do not contain their own code. Every code block is
a reference into a file under samples/ - which is in DocToolkit.sln, so CI
compiles it with -warnaserror on Linux, Windows, macOS and linux-arm64. That is
the whole point: a snippet in the documentation cannot drift from the API,
because the snippet IS the sample, and the sample stops compiling.

WHY THIS SCRIPT HAS TO EXIST. The obvious assumption is that `dotnet docfx
--warningsAsErrors` already covers this, the way it covers a dead file link.
Measured on 2026-08-09, it does not. Renaming a referenced region to one that
does not exist produced:

    exit code 0, no warning, and this in the rendered page:
        <pre><code class="lang-csharp"></code></pre>

An empty code block under prose that says "three things worth noticing in those
three lines". Worse than a broken link, which at least announces itself - this
renders as a guide that simply has nothing in it, and the pipeline stays green.

So the check is not belt-and-braces. It is the only thing standing between a
renamed region and a silently empty page.

WHAT IT CHECKS
  - the referenced file exists
  - the named region exists in it
  - the region is not empty or whitespace-only (an empty #region would satisfy
    a name check while rendering exactly the blank block above)
  - at least one reference was found at all, so the check cannot pass by
    silently matching nothing
"""

import io
import os
import re
import sys

DOCS_ROOT = "docfx"

# The shipped doc comments, whose <code source> tags render onto the API pages. Walked for the
# reason given in references(): nothing was checking them.
SRC_ROOT = "src"
SKIP_DIRS = {"bin", "obj"}

# DocFX's snippet syntax: [!code-csharp[](path#region)] - the label is usually
# empty, the fragment names a #region or an Lstart-Lend line range.
SNIPPET = re.compile(r"\[!code-(\w+)\[[^\]]*\]\(([^)#]+)#([^)]+)\)\]")

# The same thing in HTML form, which DocFX also accepts.
CODE_TAG = re.compile(r"""<code\s+source=["']([^"']+)["']\s+region=["']([^"']+)["']""")

REGION_START = re.compile(r"^\s*#region\s+(\S.*?)\s*$")
REGION_END = re.compile(r"^\s*#endregion")


def region_body(source, name):
    """Return the lines inside `#region name`, or None when it is not there.

    Tracks nesting depth so an inner region does not end an outer one.
    """
    lines = source.splitlines()
    for index, line in enumerate(lines):
        start = REGION_START.match(line)
        if not start or start.group(1) != name:
            continue

        depth, body = 1, []
        for inner in lines[index + 1:]:
            if REGION_START.match(inner):
                depth += 1
            elif REGION_END.match(inner):
                depth -= 1
                if depth == 0:
                    return body
                continue
            body.append(inner)
        return body  # unterminated region; the emptiness check below still applies
    return None


def references():
    """Every snippet reference in the guides AND in src/'s XML doc comments.

    THE SECOND HALF WAS MISSING UNTIL 2026-08-24, and it is the half CLAUDE.md warns about most
    loudly: a `<code source>` with a wrong path or region FAILS SILENTLY - DocFX emits no warning
    and renders an empty `<pre><code>` on the API page. This walked `docfx/` only, so fifteen tags
    in `src/` doc comments were checked by nothing at all.

    They were all intact when the gap was found, and the reason is luck rather than design: every
    one of them says `../../tests/...`, and the per-concern project split moved the files carrying
    them from `src/DocToolkit/` to `src/DocToolkit.Docx/` - the SAME nesting depth, so `../..`
    still reached the repository root. A split that had nested a project one level deeper would
    have broken all fifteen, on the public docs site, with every check green.
    """
    for folder, _, files in os.walk(DOCS_ROOT):
        if "_site" in folder.replace("\\", "/").split("/"):
            continue
        for name in sorted(files):
            if not name.endswith(".md"):
                continue
            path = os.path.join(folder, name)
            text = io.open(path, encoding="utf-8").read()
            for _, target, fragment in SNIPPET.findall(text):
                yield path, target, fragment
            for target, fragment in CODE_TAG.findall(text):
                yield path, target, fragment

    for folder, dirs, files in os.walk(SRC_ROOT):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
        for name in sorted(files):
            if not name.endswith(".cs"):
                continue
            path = os.path.join(folder, name)
            text = io.open(path, encoding="utf-8").read()
            for target, fragment in CODE_TAG.findall(text):
                yield path, target, fragment


def main():
    failures = []
    checked = 0
    from_docs = 0
    from_src = 0

    for doc, target, fragment in references():
        checked += 1
        if doc.replace("\\", "/").startswith(SRC_ROOT + "/"):
            from_src += 1
        else:
            from_docs += 1
        resolved = os.path.normpath(os.path.join(os.path.dirname(doc), target))

        if not os.path.isfile(resolved):
            failures.append(f"{doc}: references {target}, which does not exist")
            continue

        # Line ranges carry no region name; DocFX resolves them itself and a
        # wrong one is visible rather than blank, so only regions are checked.
        if re.fullmatch(r"L\d+(-L?\d+)?", fragment):
            continue

        body = region_body(io.open(resolved, encoding="utf-8").read(), fragment)
        if body is None:
            failures.append(f"{doc}: {target} has no '#region {fragment}'")
        elif not any(line.strip() for line in body):
            failures.append(f"{doc}: '#region {fragment}' in {target} is empty")

    if from_docs == 0:
        sys.exit("::error::No snippet references found under "
                 f"{DOCS_ROOT}/. Either the guides stopped using them or this "
                 "check stopped matching them; both are worth a red build.")

    if from_src == 0:
        sys.exit(f"::error::No <code source> tags found under {SRC_ROOT}/. The API pages' examples "
                 "are written as region references in doc comments, so finding none means either "
                 "they were removed or this check stopped matching them. Counted separately from "
                 "the guides on purpose: one total would let either source fall to zero unnoticed.")

    print(f"{checked} snippet reference(s) checked: "
          f"{from_docs} in {DOCS_ROOT}/, {from_src} in {SRC_ROOT}/ doc comments.")

    if not failures:
        print("Every referenced region exists and has content.")
        return 0

    print()
    for failure in failures:
        print(f"::error::{failure}")
    print()
    print("DocFX renders a missing region as an EMPTY code block and still exits 0, "
          "so this would have shipped as a guide with nothing in it. Fix the region "
          "name, or restore the region in the sample.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
