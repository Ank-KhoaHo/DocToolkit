# 5. Stay below 1.0.0 permanently

**Status:** accepted

## Context

Under semantic versioning, `0.x` means anything may change. Reaching 1.0.0 is a promise that
breaking changes will be rare and signposted by a major bump — a promise this package is not in a
position to keep, and would rather not make falsely.

## Decision

`0.x.y` forever, and **enforced in configuration rather than intended**:
`release-please-config.json` sets `bump-minor-pre-major: true`, so a breaking change bumps the
*minor* version (0.5.0 → 0.6.0) instead of jumping to 1.0.0.

Left at its default, release-please treats a breaking change before 1.0 as the signal to *leave*
pre-1.0 — a single `feat!:` commit would have shipped 1.0.0 with nobody choosing it.

## Consequences

Consumers get no semver protection, so the burden moves to being explicit: the public API is pinned
to an approved file, and any change to it appears as a reviewable diff in the pull request that
makes it. That review is the only protection a `0.x` package can offer, which is why it is
mechanical rather than a matter of care.

It also means **a member can never really be removed**. Anything added is permanent in practice, and
"we can take it out later" is not available as an argument for adding something speculative.

## What would change this

A deliberate decision that the API is stable and the maintenance commitment is sustainable — not the
version number simply feeling overdue.
