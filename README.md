# nuget-stats

Data branch. **Orphan — it shares no history with `main` and is never merged into it.**
Written once a day at 13:17 UTC by `.github/workflows/nuget-stats.yml`, which lives on `main`.

| file | what it is |
|---|---|
| `downloads.csv` | one row per `(date, package, version)`; `cumulative` is the all-time total |
| `runs.csv` | one row per `(date, workflow)`; our own CI activity that day |
| `report.md` | rendered summary — GitHub displays this when you click it |
| `report.html` | the same, self-contained, for opening locally |

## Why this exists

Measured 2026-08-08: 677 restoring workflow runs in six weeks, each re-downloading the full
dependency closure. Roughly 880 of the four newest versions' downloads were our own CI. The
nuget.org counter could not answer "does anyone actually use this?" NuGet caching removed that
traffic; this branch records what is left.

## This records. It does not yet interpret.

**Nothing here tells you whether anyone outside our own CI uses the package.** An adoption
signal was designed, built, reviewed and removed before this shipped, for two measured reasons:

- Its gate required a date present in `runs.csv` whose run count was zero — but the only thing
  that writes that file is a counter, which never yields zero, so a zero-run day produced no row
  at all. The two halves of the condition were mutually exclusive and it could never fire.
- Repairing that alone would have made it wrong in the dangerous direction. Downloads
  accumulated across a multi-day gap were charged entirely to the CI status of the gap's *last*
  day: 60 downloads spread over three days each carrying 40 CI runs rendered as external usage.

Collection ships regardless, because **the history cannot be rebuilt**. The nuget.org search API
serves only current cumulative totals — a day not collected is lost permanently, while any
analysis can be re-run over the whole history whenever a sound method exists. That asymmetry is
the entire reason this branch exists now rather than later.

## Reading the data honestly

A `*` after a version marks a clamped reading: a region reported a total below one already
recorded, and cumulative counts cannot decrease, so the previous high was kept.

What no later analysis of this data will ever be able to tell you:

- **No client split.** A crawler and a developer are indistinguishable here. The per-client
  breakdown on nuget.org is rendered client-side — the served HTML holds one `<td>`, no version
  strings and no client names — so it cannot be fetched without a headless browser.
- **One person downloading ten times looks like ten people.**
- **Proxy consumers are invisible.** One JFrog Artifactory fetch may serve an entire company, so
  this under-counts in a direction it cannot measure.
- **A crawler floor exists** of roughly two downloads per version per day, on every version
  regardless of age. Subtract it before concluding anything.
- **The index lags one to two days**, and the two search regions disagree — measured 2026-08-08,
  one reported 3642 where the other reported 5410 for the same package at the same instant. Both
  are queried and the higher wins.

Raw facts only. Every delta and sparkline is computed at render time, so a bug in the maths costs
a re-render rather than corrupting the history.
