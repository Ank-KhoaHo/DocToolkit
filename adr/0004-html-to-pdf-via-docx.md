# 4. HTML → PDF pivots through DOCX

**Status:** accepted

## Context

There is no permissively-licensed, NuGet-only, Linux-safe library that renders HTML to PDF directly.
The free renderers that exist **are browsers**, and a browser is a native binary — several hundred
megabytes of it, plus its own CVE feed.

## Decision

`HtmlToPdfConverter` composes the two converters that do exist: HTML → DOCX, then DOCX → PDF. It is
a composition, not a third conversion, and contains no rendering logic of its own.

## Consequences

**Fidelity is bounded by what HtmlToOpenXml maps into WordprocessingML, not by what a browser would
render.** Complex CSS layout — flexbox, grid, floats, absolute positioning — does not survive. Text,
headings, tables, lists, inline styling and images do.

That is a real limitation and is documented as one in the package README rather than left for a
consumer to discover. Someone expecting browser-grade output will be disappointed, and should be
told before they take the dependency rather than after.

## What would change this

A permissively-licensed, pure-managed HTML renderer existing. If one appears, the constraint to
check first is the resolved dependency graph, not the rendering quality — the reason this decision
exists is that every candidate so far failed constraint 2 before fidelity was ever in question.

**Do not reimplement conversion inside `HtmlToPdfConverter`.** Keeping it a composition is what
guarantees the DOCX and PDF outputs cannot disagree.
