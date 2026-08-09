# 1. The four constraints, enforced by CI

**Status:** accepted

## Context

.NET has no shortage of document libraries. Choosing one usually means accepting at least one of:
a licence that is not free for commercial use, a native binary that breaks on Linux or arm64, an
external renderer (a browser or LibreOffice) that has to be installed into the image, or a runtime
network fetch that fails on an air-gapped machine.

This package exists only because it satisfies all four at once. Any change that breaks one makes it
pointless — there would be no reason to choose it over the incumbents.

## Decision

Four constraints, and **each is enforced by CI rather than stated as an intention**:

1. **Permissive licences only** — MIT / Apache-2.0 / BSD.
2. **NuGet only** — no browser download, no LibreOffice, **no native binaries**.
3. **Runs everywhere .NET does** — the full suite runs on Linux x64, Linux arm64, Windows and macOS.
4. **No runtime network I/O** by default.

## Consequences

The important property is that **all four are facts about the resolved dependency graph, not about
this code**. A single `dotnet add package` can break every one of them without a line of our source
changing, and it will restore, build and pass a naive test run while doing so.

That is why they are guards and not guidance: a dependency guard test, a native-binary scan, a
four-platform matrix, and 37 tests that count sockets. It has already caught a real breach (see
[ADR 2](0002-no-native-binaries.md)).

## What would change this

Nothing short of the package's purpose changing. A capability that cannot be delivered under these
constraints is out of scope, not a reason to relax them — but "cannot" should be **measured**, not
assumed. Two features were shelved as impossible under these constraints and turned out to be
straightforward once someone actually checked the dependency graph.
