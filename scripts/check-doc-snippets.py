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

import glob
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


# ---------------------------------------------------------------------------
# The sibling failure: an @uid cross-reference that silently renders as TEXT.
#
# MEASURED 2026-08-26, on this repository's own shipped docs site. Two guides
# carried an @uid that resolved to nothing, and `dotnet docfx --warningsAsErrors`
# reported **0 warnings, 0 errors** for both - exactly like the empty code block
# above. The reader gets a paragraph with a raw "@DocToolkit.XlsxRuleKind" in it.
#
# Two distinct causes, and only the first is guessable from reading the markdown:
#
#   1. A TRAILING ")". "(@DocToolkit.XlsxRuleKind)" never resolves - the shorthand
#      swallows the bracket into the uid. Three of these were introduced at once.
#   2. A METHOD uid without its overload-group star. "@DocToolkit.PptxSlide.Titled"
#      matches nothing, because the real uids are
#      "DocToolkit.PptxSlide.Titled(System.String,System.String[])" and
#      "DocToolkit.PptxSlide.Titled*". This one had been live on the published
#      site and is invisible to any purely syntactic rule.
#
# So the check has two tiers and SAYS WHICH IT RAN, rather than implying the
# wider cover when the narrower one is all that was possible.
# ---------------------------------------------------------------------------

# The trailing `N is the GENERIC ARITY suffix - DocToolkit.ConversionResult`1 is a real
# uid, not a uid followed by inline code. Missing that produced this guard's first false
# positive, and a guard that cries wolf is a guard somebody switches off.
UID_IN_PROSE = re.compile(
    r"(?<![\w`])@([A-Za-z_][A-Za-z0-9_]*(?:\.[A-Za-z0-9_]+)+(?:`\d+)?\*?)")


def strip_code(markdown):
    """Drop FENCED code, where an @uid is a literal rather than a link.

    Inline code is deliberately left alone. Stripping it glued three separate spans
    into one imaginary uid - ConversionResultValueWarningsHasLoss - and reported it as
    broken. The lookbehind on UID_IN_PROSE already refuses a uid opening inside a
    backtick span, so nothing is lost by leaving inline code in place.
    """
    return re.sub(r"```.*?```", "", markdown, flags=re.S)


def known_uids():
    """The uid set DocFX itself generated, or None when metadata has not been built.

    Derived rather than hand-listed - the principle gen-third-party-notices.py and
    check-readme-coverage.py already follow. docfx/api/ is gitignored and absent on a
    fresh checkout, so this returns None there and the caller degrades loudly.
    """
    files = glob.glob("docfx/api/*.yml")
    if not files:
        return None

    uids = set()
    for path in files:
        for line in io.open(path, encoding="utf-8"):
            match = re.match(r"^-?\s*uid:\s*(\S+)\s*$", line)
            if match:
                uids.add(match.group(1))
    return uids or None


def check_xrefs():
    failures = []
    scanned = 0
    uids = known_uids()

    for path in sorted(glob.glob(f"{DOCS_ROOT}/**/*.md", recursive=True)):
        text = strip_code(io.open(path, encoding="utf-8").read())
        for match in UID_IN_PROSE.finditer(text):
            scanned += 1
            uid, tail = match.group(1), text[match.end():match.end() + 1]

            # Tier 1 - syntax. True regardless of what the API surface holds.
            if tail == ")":
                failures.append(
                    f"{path}: @{uid} is followed by ')', which the shorthand swallows into the "
                    f"uid. Write [{uid.rsplit('.', 1)[-1]}](xref:{uid}) instead.")
                continue

            # Tier 2 - resolution. Only possible once DocFX has emitted metadata.
            #
            # The uid is looked up EXACTLY as written. An earlier version accepted a bare
            # method name whenever the "*" overload group existed - which made it blind to
            # the one bug on the real site that motivated the whole check, because the
            # star's PRESENCE is precisely what proves the bare form is a method and will
            # not resolve. Sabotage caught that; reading the code had not.
            if uids is not None and uid not in uids:
                star = f"{uid}*"
                if star in uids or any(u.startswith(f"{uid}(") for u in uids):
                    hint = f" It is a method - write xref:{star}, the overload group."
                else:
                    hint = " No such uid was generated."
                failures.append(f"{path}: @{uid} matches no uid.{hint}")

    if not failures:
        if uids is None:
            print(f"{scanned} @uid reference(s) checked for SYNTAX only - docfx/api/ has not "
                  "been generated, so whether each one RESOLVES was not verified. Run "
                  "`dotnet docfx docfx/docfx.json` first for the full check.")
        else:
            print(f"{scanned} @uid reference(s) checked against {len(uids)} generated uids; "
                  "every one resolves.")
        return []

    return failures


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

    xref_failures = check_xrefs()
    failures.extend(xref_failures)

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
    print("DocFX renders a missing region as an EMPTY code block, and an unresolved @uid as "
          "RAW TEXT, and still exits 0 for both - so this would have shipped as a guide with "
          "nothing in it, or with an @DocToolkit.Something sitting in a sentence.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
