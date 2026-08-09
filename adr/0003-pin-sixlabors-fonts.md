# 3. Pin `SixLabors.Fonts` to `[1.0.0]`

**Status:** accepted

## Context

`OfficeIMO.Word` requests `SixLabors.Fonts` in the range `[1.0.0, 3.0.0)`, so an ordinary restore is
free to resolve 2.x.

**Version 2.x moves to the Six Labors Split License**: Apache-2.0 below $1M annual revenue, and
commercial above it. Floating to 2.x would therefore move this package off permissive licensing —
silently, on somebody else's release schedule, with no change to this repository at all.

## Decision

Pin it exactly: `[1.0.0]`. CI asserts the pin still holds, and Dependabot is told to ignore majors
for this package specifically so a doomed pull request is not reopened every week.

## Consequences

The pin is a licensing wall, not a compatibility one, and that distinction matters when someone
finds it and assumes it is stale. Bug fixes in 2.x are unavailable.

It also has to be repeated in more than one place. Dependabot `ignore` rules are scoped to the block
that declares them, and the test projects reach `src/` through `ProjectReference` — a grouped update
from the tests block once proposed `SixLabors.Fonts [1.0.0] -> [3.0.0]`, straight past the wall. CI
caught it.

## What would change this

Six Labors changing the licence terms of a later major, or `OfficeIMO.Word` dropping the dependency.
Either needs the new terms read and the resolved graph re-measured — not an assumption that a newer
version is fine.
