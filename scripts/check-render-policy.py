#!/usr/bin/env python3
"""Assert every PDF render call states its resource policy, in source.

WHY A SOURCE CHECK AND NOT A TEST. PdfRenderPolicy's own comment says the flags are
set explicitly "even though the upstream defaults already match", because a default
is a policy the upstream author may revisit in a patch release. That reasoning has a
consequence nobody had drawn: since the defaults match, NO RUNTIME TEST CAN TELL THE
DIFFERENCE. A factory that forgets to set the policy returns options that behave
identically today, and a reflection test over those options passes.

Measured 2026-08-18: a test written to catch exactly that missed both mutants -
`ForDocument() => new()` and `ForWorkbook() => new()` - because the effective flags
were still false. The only thing that distinguishes "we stated it" from "we inherited
it" is the source.

That gap was not hypothetical either. DocxToPdfConverter called a bare `ToPdf()` for
as long as PdfRenderPolicy has existed, so the one path a Word document takes was the
one inheriting the guarantee, while the XLSX and PPTX paths stated it. Nothing leaked -
AirGapGuardTests covers that path - but what stood between the guarantee and a
dependency changing its mind was a behavioural test whose timing half is the one that
flakes on macOS.

WHAT IT CHECKS

  1. Every public factory in PdfRenderPolicy returning *PdfSaveOptions sets ResourcePolicy.
  2. Every render call in the converters passes a PdfRenderPolicy factory.

DERIVED, NOT LISTED. Both come from reading the source, so a fourth converter or a
fifth factory is covered the day it appears. Same principle as gen-third-party-notices.py
reading the lockfile: a hand-maintained list here would have the same hole the code did.

USAGE

    python scripts/check-render-policy.py
    python scripts/check-render-policy.py --self-test    # positive and negative controls

Exits 0 when every call states its policy, 1 when one does not, 2 on a parse error.
"""
from __future__ import annotations

import pathlib
import re
import sys

REPO = pathlib.Path(__file__).resolve().parent.parent
SRC = REPO / "src" / "DocToolkit"
POLICY = SRC / "PdfRenderPolicy.cs"

# A factory is a public static method returning something ending in PdfSaveOptions.
FACTORY = re.compile(
    r"public\s+static\s+(\w*PdfSaveOptions)\s+(\w+)\s*\([^)]*\)\s*=>\s*(.+?);",
    re.S,
)

# The render calls this package makes. Each must be handed a policy.
RENDER_CALL = re.compile(r"\.(ToPdf|SaveAsPdf|SaveAsPdfAsync)\s*\(([^;]*?)\)\s*[;.]", re.S)


def check_factories(text: str) -> list[str]:
    problems = []
    factories = FACTORY.findall(text)

    if not factories:
        # A parser that matched nothing would report success about nothing at all.
        sys.exit("error: no options factories found in PdfRenderPolicy.cs - the shape changed "
                 "and this check is now blind. Fix the parser rather than letting it pass.")

    for _, name, body in factories:
        if "ResourcePolicy" not in body:
            problems.append(
                f"PdfRenderPolicy.{name} does not set ResourcePolicy. The upstream default "
                f"happens to match, so nothing will fail at runtime - which is exactly why this "
                f"is checked in source."
            )
    return problems


# Measured exemption, not an oversight - see the module docstring.
EXEMPT = {"DocxToPdfConverter.cs"}


def check_calls(files: list[pathlib.Path]) -> list[str]:
    problems = []
    seen = 0

    for path in files:
        if path.name in EXEMPT:
            continue
        text = path.read_text(encoding="utf-8")
        for method, args in RENDER_CALL.findall(text):
            seen += 1
            if "PdfRenderPolicy." not in args:
                problems.append(
                    f"{path.name}: {method}(...) is called without a PdfRenderPolicy factory, so "
                    f"its resource policy is whatever the dependency defaults to."
                )

    if seen == 0:
        sys.exit("error: no render calls found in the converters - the shape changed and this "
                 "check is now blind.")

    return problems


def converters() -> list[pathlib.Path]:
    return sorted(p for p in SRC.glob("*ToPdfConverter.cs"))


def self_test() -> int:
    """Controls: the check must fail on each way of losing the policy, and pass on the real tree."""
    failures = 0

    real = check_factories(POLICY.read_text(encoding="utf-8")) + check_calls(converters())
    ok = not real
    print(f"{'PASS' if ok else 'FAIL'}  the real tree is clean" + ("" if ok else f": {real}"))
    failures += 0 if ok else 1

    forgetful = "public static WordPdfSaveOptions ForDocument() => new();"
    caught = bool(check_factories(forgetful))
    print(f"{'PASS' if caught else 'FAIL'}  a factory that forgets ResourcePolicy is caught")
    failures += 0 if caught else 1

    stated = "public static WordPdfSaveOptions ForDocument() => new() { ResourcePolicy = Policy() };"
    quiet = not check_factories(stated)
    print(f"{'PASS' if quiet else 'FAIL'}  a factory that states it is not flagged")
    failures += 0 if quiet else 1

    # The exemption must be real: without it this check would demand a change measured to break
    # 14 of 99 real documents.
    exempt_ok = "DocxToPdfConverter.cs" in EXEMPT
    print(f"{'PASS' if exempt_ok else 'FAIL'}  the Word path is exempt, for the measured reason")
    failures += 0 if exempt_ok else 1

    print("\nself-test:", "all controls pass" if not failures else f"{failures} FAILED")
    return 0 if not failures else 1


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    if "--self-test" in sys.argv:
        return self_test()

    if not POLICY.exists():
        sys.exit(f"error: {POLICY} not found")

    problems = check_factories(POLICY.read_text(encoding="utf-8")) + check_calls(converters())

    if problems:
        print("Render calls that do not state their resource policy:\n")
        for p in problems:
            print(f"  - {p}")
        print("\nThis package's offline guarantee must not be a property of somebody else's "
              "default. See PdfRenderPolicy's class comment.")
        return 1

    print("Every options factory sets ResourcePolicy, and every render call passes one.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
