#!/usr/bin/env python3
"""Record this package's nuget.org download counts daily.

This RECORDS; it does not yet INTERPRET. An adoption signal was designed,
built, and removed before shipping - see summarise()'s docstring for exactly
how it failed. Collection ships anyway because the history cannot be rebuilt:
the search API serves only current cumulative totals, so every uncollected day
is lost permanently, while any analysis can be re-run over the whole history
whenever a sound method exists.

WHY THIS EXISTS. Measured 2026-08-08: 677 restoring workflow runs in six weeks,
each re-downloading the full dependency closure, so roughly 880 of the four
newest versions' downloads were self-inflicted. The nuget.org counter could not
answer "does anyone actually use this?" because CI drowned it. Caching (added to
ci.yml on 2026-08-08) removed that traffic; this script measures what is left.

TWO REGIONS, AND THAT IS NOT BELT-AND-BRACES. Measured 2026-08-08, the same
package at the same instant:

    azuresearch-usnc   totalDownloads 3642, with 0.10.0 and 0.11.0 both 0
    azuresearch-ussc   totalDownloads 5410, with 0.10.0 at 279 and 0.11.0 at 247

A 33% divergence, and ussc matched the nuget.org stats page exactly while usnc
was days stale. Querying one region does not yield a less reliable number, it
yields a WRONG one, and nothing in the response says so. Every version therefore
takes the maximum across regions.

WHAT THIS CANNOT SEE. The per-client breakdown - NuGet MSBuild task vs browser
vs crawler - is rendered client-side by Knockout on nuget.org/stats: the served
HTML contains one <td>, zero version strings and no client names. Scraping it
would need headless Chromium, which is exactly the dependency docfx.json already
sets _enableSearch:false to avoid. So here a bot and a developer look identical.

CUMULATIVE NEVER DECREASES. Region divergence would otherwise manufacture
negative deltas: a day sampled from ussc followed by a day when only usnc
answered would look like downloads being handed back. Values are clamped to the
highest previously recorded figure and the row is flagged.

IDEMPOTENT BY KEY. Re-running for a date that already has rows updates them in
place. A retried run, a manual dispatch, or a schedule slipping past midnight
would otherwise append duplicates - and a duplicate row silently doubles that
day's delta, which reads exactly like the external usage this exists to detect.

Usage:
    python scripts/nuget-stats.py collect --data-dir DIR --runs-json FILE
    python scripts/nuget-stats.py render  --data-dir DIR

    # Offline, against the committed fixture - no network, deterministic:
    python scripts/nuget-stats.py collect --data-dir /tmp/x \\
        --api-fixture scripts/testdata/nuget-stats-fixture.json --date 2026-08-09
    python scripts/nuget-stats.py render --data-dir /tmp/x
"""

import argparse
import collections
import csv
import datetime
import html
import http.client
import json
import os
import sys
import urllib.error
import urllib.request

PACKAGES = (
    "Ank.DocToolkit",
    "Ank.DocToolkit.Extensions.DependencyInjection",
)

# Both, always. See the module docstring: one region is wrong, not merely worse.
REGIONS = ("usnc", "ussc")

QUERY = (
    "https://azuresearch-{region}.nuget.org/query"
    "?q=packageid:{package}&prerelease=true&semVerLevel=2.0.0"
)

TIMEOUT_SECONDS = 30

DOWNLOADS_HEADER = ("date", "package", "version", "cumulative", "clamped")
RUNS_HEADER = ("date", "workflow", "runs")


def parse_payload(payload):
    """{version: downloads} from one search-API response.

    Returns {} for a package the region has not indexed - the API answers with
    an empty `data` list rather than an error, and that is a normal state for a
    version published minutes ago, not a failure.
    """
    data = payload.get("data") or []
    if not data:
        return {}
    versions = data[0].get("versions") or []
    return {entry["version"]: int(entry["downloads"]) for entry in versions}


def reconcile(per_region):
    """Merge per-region {version: downloads} maps, taking the max per version."""
    merged = {}
    for counts in per_region:
        for version, downloads in counts.items():
            if downloads > merged.get(version, -1):
                merged[version] = downloads
    return merged


def fetch_region(package, region):
    url = QUERY.format(region=region, package=package.lower())
    with urllib.request.urlopen(url, timeout=TIMEOUT_SECONDS) as response:
        return parse_payload(json.load(response))


def read_rows(path):
    if not os.path.exists(path):
        return []
    with open(path, newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle))


def write_rows(path, header, rows):
    """Write via a temp file and an atomic replace.

    `open(path, "w")` truncates before a single row is written, so a
    DictWriter that raises partway - a row carrying a field the header does
    not declare, say - leaves the file empty. Measured: a runs.csv holding a
    legitimate prior row was reduced to its header alone, and because the
    caller catches that exception it was reported as "continuing" while the
    history was already gone.

    This matters most for downloads.csv. Run counts can be rebuilt any time
    with `gh run list --created <date>`; download counts cannot, because the
    search API serves only the current cumulative total. A truncated
    downloads.csv is history nothing can reconstruct.

    os.replace is atomic on both POSIX and Windows, so the original file
    survives untouched unless the whole new file was written successfully.
    """
    tmp = f"{path}.tmp"
    try:
        with open(tmp, "w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=list(header))
            writer.writeheader()
            writer.writerows(rows)
        os.replace(tmp, path)
    except BaseException:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise


def upsert(rows, new_rows, key_fields):
    """Replace rows matching a key, append the rest. Order is preserved."""
    index = {tuple(row[f] for f in key_fields): i for i, row in enumerate(rows)}
    for row in new_rows:
        key = tuple(row[f] for f in key_fields)
        if key in index:
            rows[index[key]] = row
        else:
            index[key] = len(rows)
            rows.append(row)
    return rows


def highest_recorded(rows, package, version):
    """The largest cumulative ever recorded for this version, or None.

    Cumulative counts only grow, so the maximum over history is the floor any
    new reading must not fall below.
    """
    best = None
    for row in rows:
        if row.get("package") != package or row.get("version") != version:
            continue
        try:
            value = int(row["cumulative"])
        except (TypeError, ValueError, KeyError):
            print(f"warning: unreadable cumulative for {package} {version} "
                  f"on {row.get('date', '?')}; ignoring that row", file=sys.stderr)
            continue
        if best is None or value > best:
            best = value
    return best


def gather_downloads(fixture):
    """{package: (counts, complete)} where counts is {version: downloads}.

    `complete` is False when a region we expected to hear from ERRORED, which
    is a different thing from a region answering "I have not indexed this yet".
    The first means our number may be missing data; the second is a normal
    state for a version published minutes ago. Fix 4 depends on the difference.

    A malformed payload degrades that region only. Anything else takes down the
    healthy package alongside the broken one, and this job runs unattended.
    """
    readable = (AttributeError, TypeError, ValueError, KeyError)
    gathered = {}
    for package in PACKAGES:
        per_region = []
        failures = 0
        if fixture is not None:
            for payload in fixture.get(package, []):
                try:
                    per_region.append(parse_payload(payload))
                except readable as error:
                    failures += 1
                    print(f"warning: {package} fixture payload unreadable: {error}",
                          file=sys.stderr)
        else:
            for region in REGIONS:
                try:
                    per_region.append(fetch_region(package, region))
                except (urllib.error.URLError, OSError,
                        http.client.HTTPException, *readable) as error:
                    failures += 1
                    print(f"warning: {package} via {region}: {error}", file=sys.stderr)
        if not per_region:
            # Every region failed. Skip the package rather than writing a zero,
            # which would look like every user vanishing overnight.
            print(f"warning: no region answered for {package}; skipping", file=sys.stderr)
            continue
        gathered[package] = (reconcile(per_region), failures == 0)
    return gathered


def collect(args):
    today = args.date or datetime.datetime.now(datetime.timezone.utc).date().isoformat()
    try:
        datetime.date.fromisoformat(today)
    except ValueError:
        print(f"error: --date must be YYYY-MM-DD, got {today!r}", file=sys.stderr)
        return 2

    os.makedirs(args.data_dir, exist_ok=True)
    downloads_path = os.path.join(args.data_dir, "downloads.csv")
    runs_path = os.path.join(args.data_dir, "runs.csv")

    # The CI half first, and deliberately independent of the download fetch.
    # These are unrelated signals and a nuget.org outage must not also cost us
    # the day's CI record. The asymmetry is not arbitrary: run counts can be
    # rebuilt at any time with `gh run list --created <date>`, while download
    # counts cannot - the search API serves only the current cumulative total,
    # so a lost day is lost permanently.
    if args.runs_json:
        try:
            with open(args.runs_json, encoding="utf-8") as handle:
                runs = json.load(handle)
            counts = collections.Counter(run["workflowName"] for run in runs)
            run_rows = [
                {"date": today, "workflow": name, "runs": str(count)}
                for name, count in sorted(counts.items())
            ]
            write_rows(
                runs_path,
                RUNS_HEADER,
                upsert(read_rows(runs_path), run_rows, ("date", "workflow")),
            )
            print(f"runs: {len(run_rows)} workflows for {today}")
        except (OSError, ValueError, TypeError, KeyError) as error:
            # The rebuildable signal must never block the irrecoverable one.
            # Run counts can be recovered any time with `gh run list --created`;
            # download counts cannot, because the search API serves only the
            # current cumulative total. So a bad runs.json costs a warning and
            # nothing more - it must not stop downloads.csv being written.
            print(f"warning: runs not recorded ({error}); continuing",
                  file=sys.stderr)

    fixture = None
    if args.api_fixture:
        with open(args.api_fixture, encoding="utf-8") as handle:
            fixture = json.load(handle)

    existing = read_rows(downloads_path)
    gathered = gather_downloads(fixture)
    if not gathered:
        print("error: no download data from any region for any package", file=sys.stderr)
        return 1

    fresh = []
    for package, (counts, complete) in gathered.items():
        for version, downloads in sorted(counts.items()):
            floor = highest_recorded(existing, package, version)

            if floor is None and not complete:
                # First ever sighting of this version, from an incomplete read.
                # With no floor there is nothing to clamp against, so a stale
                # region alone would write a too-low number as fact - and the
                # correction would arrive later as a false SPIKE, on a day that
                # may have had no CI activity, which the report would then flag
                # as external usage. Skipping costs one day and cannot mislead.
                print(f"warning: {package} {version} seen only in an incomplete "
                      f"read with no prior value; not recording", file=sys.stderr)
                continue

            clamped = floor is not None and downloads < floor
            if clamped:
                print(
                    f"warning: {package} {version} reported {downloads} "
                    f"below recorded {floor}; clamping",
                    file=sys.stderr,
                )
            fresh.append({
                "date": today,
                "package": package,
                "version": version,
                "cumulative": str(floor if clamped else downloads),
                "clamped": "yes" if clamped else "",
            })

    write_rows(
        downloads_path,
        DOWNLOADS_HEADER,
        upsert(existing, fresh, ("date", "package", "version")),
    )
    print(f"downloads: {len(fresh)} rows for {today}")
    return 0


SPARK = "▁▂▃▄▅▆▇█"


def sparkline(values):
    """Unicode bars for a series of daily deltas.

    Negative values are floored to zero rather than passed through: Python's
    negative indexing would select SPARK[-1], the TALLEST bar, so a decrease
    would render as a spike. `collect` clamps cumulative counts so this should
    be unreachable, but a chart that lies when its invariant breaks is worse
    than one that flatlines.
    """
    if not values:
        return ""
    top = max(values)
    if top <= 0:
        return SPARK[0] * len(values)
    return "".join(
        SPARK[min(len(SPARK) - 1, max(0, v) * (len(SPARK) - 1) // top)]
        for v in values
    )


def series(rows):
    """{(package, version): {date: cumulative}} from downloads.csv rows.

    A row with an unreadable cumulative is skipped with a warning rather than
    raising, matching highest_recorded. One hand-edited or legacy row must not
    cost the whole report.
    """
    out = collections.defaultdict(dict)
    for row in rows:
        try:
            out[(row["package"], row["version"])][row["date"]] = int(row["cumulative"])
        except (TypeError, ValueError, KeyError):
            print(f"warning: skipping unreadable downloads row {row!r}", file=sys.stderr)
    return out


def runs_by_date(rows):
    """{date: total runs} from runs.csv rows, skipping unreadable ones."""
    out = collections.Counter()
    for row in rows:
        try:
            out[row["date"]] += int(row["runs"])
        except (TypeError, ValueError, KeyError):
            print(f"warning: skipping unreadable runs row {row!r}", file=sys.stderr)
    return out


def cell(value):
    """Blank, not zero, for a figure that is not yet measurable.

    "We measured no change" and "we do not have enough samples" are different
    answers, and printing 0 for both hides which one you are looking at.
    """
    return "—" if value is None else str(value)


def span_days(dates):
    """Calendar days between the first and last of a sorted date list.

    Returns None rather than raising when a date will not parse. One
    hand-edited row, merge artifact or legacy row must not take down the whole
    report - the same tolerance highest_recorded applies to a malformed
    cumulative. Callers that need a number fall back to 1.
    """
    if len(dates) < 2:
        return None
    try:
        first = datetime.date.fromisoformat(dates[0])
        last = datetime.date.fromisoformat(dates[-1])
    except (TypeError, ValueError):
        print(f"warning: unparseable date in {dates[0]!r}..{dates[-1]!r}; "
              f"span not computed", file=sys.stderr)
        return None
    return (last - first).days


def summarise(downloads_rows, runs_rows):
    """One record per (package, version), highest total first within a package.

    Deliberately computes no adoption signal. An earlier version carried a
    `quiet` column - downloads accumulated on days our own CI ran nothing -
    and it was removed rather than repaired, for two measured reasons.

    It could never fire: the gate required a date present in runs.csv whose
    run count was zero, but the only thing that writes runs.csv is a Counter,
    which never yields zero, so a zero-run day produced no row at all and the
    two halves of the condition were mutually exclusive.

    And repairing that alone would have made it lie in the dangerous
    direction: downloads accumulated across a multi-day gap were charged
    entirely to the CI status of the gap's LAST day, so 60 downloads spread
    over three days carrying 40 CI runs each rendered as external usage.

    What remains is only what the data supports: totals, deltas, and our own
    CI volume beside them. Answering "does anyone actually use this?" needs a
    method this file does not yet have - but the history being recorded here
    cannot be rebuilt later, so it is worth collecting while that is designed.
    """
    by_key = series(downloads_rows)
    per_day = runs_by_date(runs_rows)

    all_dates = sorted({row["date"] for row in downloads_rows})
    recent = all_dates[-7:]
    clamped_keys = {
        (row["package"], row["version"])
        for row in downloads_rows
        if row.get("clamped")
    }

    summary = []
    for (package, version), points in by_key.items():
        dates = sorted(points)
        latest = points[dates[-1]]

        window_dates = [d for d in dates if d in recent]
        window = [points[d] for d in window_dates]
        delta = window[-1] - window[0] if len(window) > 1 else None
        # The window is whichever of THIS version's dates fall in the last seven
        # seen anywhere, so it is not always seven days. Report the span rather
        # than labelling a three-day figure "7d".
        span = span_days(window_dates)
        daily = [window[i + 1] - window[i] for i in range(len(window) - 1)]
        runs_recent = sum(per_day.get(d, 0) for d in recent)

        summary.append({
            "package": package,
            "version": version,
            "latest": latest,
            "delta": delta,
            "span": span,
            "runs_recent": runs_recent,
            "clamped": (package, version) in clamped_keys,
            "spark": sparkline(daily),
        })
    summary.sort(key=lambda r: (r["package"], -r["latest"]))
    return summary, all_dates


LIMITS = """\
**This does not yet tell you whether anyone outside our own CI uses the package.**
It records the history needed to answer that once a sound method exists, because
the nuget.org search API serves only current totals - a day not collected is lost
permanently, while the analysis can be re-run over the whole history at any time.

A `*` marks a version whose reading was clamped: a region reported a total below
one already recorded, and cumulative counts cannot decrease, so the previous high
was kept.

What this can never tell you, however it is analysed later:

- **No client split.** A crawler and a developer are indistinguishable here; the
  per-client breakdown on nuget.org is rendered client-side and cannot be fetched
  without a headless browser.
- **One person downloading ten times looks like ten people.**
- **Proxy consumers are invisible.** One JFrog Artifactory fetch may serve a whole
  company, so this under-counts in a direction it cannot measure.
- **The index lags one to two days**, and the two search regions disagree - measured
  2026-08-08, one reported 3642 where the other reported 5410. Both are queried and
  the higher wins.
"""


def render_markdown(summary, dates):
    recent = dates[-7:]
    recent_span = span_days(recent)
    window_label = f" /{recent_span}d" if recent_span else ""
    lines = [
        "# nuget.org downloads",
        "",
        f"Generated from {len(dates)} day(s) of samples"
        + (f", {dates[0]} to {dates[-1]}." if dates else "."),
        "",
        f"| package | version | total | recent Δ | CI runs{window_label} | trend |",
        "|---|---|---:|---:|---:|---|",
    ]
    for row in summary:
        short = row["package"].replace("Ank.DocToolkit", "core").replace(
            "core.Extensions.DependencyInjection", "extensions")
        delta = cell(row["delta"])
        if row["span"]:
            delta += f" /{row['span']}d"
        version = row["version"] + ("*" if row["clamped"] else "")
        lines.append(
            f"| {short} | {version} | {row['latest']} | {delta} | "
            f"{row['runs_recent']} | `{row['spark']}` |"
        )
    lines += ["", LIMITS]
    return "\n".join(lines) + "\n"


def render_html(summary, dates):
    span = f"{dates[0]} to {dates[-1]}" if dates else "no samples yet"
    recent = dates[-7:]
    recent_span = span_days(recent)
    window_label = f" /{recent_span}d" if recent_span else ""
    body = []
    for row in summary:
        delta = cell(row["delta"])
        if row["span"]:
            delta += f" /{row['span']}d"
        version = html.escape(row["version"]) + ("*" if row["clamped"] else "")
        body.append(
            f"<tr><td>{html.escape(row['package'])}</td><td>{version}</td>"
            f"<td class=n>{row['latest']}</td><td class=n>{delta}</td>"
            f"<td class=n>{row['runs_recent']}</td>"
            f"<td class=s>{row['spark']}</td></tr>"
        )
    return f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>nuget.org usage</title>
<style>
  :root {{ color-scheme: light dark; --fg:#111; --bg:#fff; --line:#d0d7de; }}
  @media (prefers-color-scheme: dark) {{
    :root {{ --fg:#e6edf3; --bg:#0d1117; --line:#30363d; }}
  }}
  body {{ font:15px/1.5 system-ui,sans-serif; color:var(--fg); background:var(--bg);
         margin:0; padding:2rem 1rem; }}
  main {{ max-width:60rem; margin:0 auto; }}
  table {{ border-collapse:collapse; width:100%; }}
  th,td {{ padding:.4rem .6rem; border-bottom:1px solid var(--line); text-align:left; }}
  .n {{ text-align:right; font-variant-numeric:tabular-nums; }}
  .s {{ font-family:ui-monospace,monospace; letter-spacing:1px; }}
  .wrap {{ overflow-x:auto; }}
</style></head><body><main>
<h1>nuget.org usage</h1>
<p>{len(dates)} day(s) of samples, {span}.</p>
<div class="wrap"><table>
<thead><tr><th>package</th><th>version</th><th class=n>total</th><th class=n>recent Δ</th>
<th class=n>CI runs{window_label}</th><th>trend</th></tr></thead>
<tbody>
{chr(10).join(body)}
</tbody></table></div>
<p><strong>This does not yet tell you whether anyone outside our own CI uses the
package.</strong> It records the history needed to answer that once a sound method
exists - nuget.org serves only current totals, so a day not collected is lost
permanently, while the analysis can be re-run over the whole history at any time.
<code>/Nd</code> is the number of calendar days a figure covers. A dash means not
enough samples yet, which is different from zero. A <code>*</code> after a version
marks a clamped reading: a region reported a total below one already recorded.</p>
</main></body></html>
"""


def render(args):
    downloads_rows = read_rows(os.path.join(args.data_dir, "downloads.csv"))
    runs_rows = read_rows(os.path.join(args.data_dir, "runs.csv"))
    if not downloads_rows:
        print("error: no downloads.csv rows to render", file=sys.stderr)
        return 1
    summary, dates = summarise(downloads_rows, runs_rows)
    with open(os.path.join(args.data_dir, "report.md"), "w", encoding="utf-8") as handle:
        handle.write(render_markdown(summary, dates))
    with open(os.path.join(args.data_dir, "report.html"), "w", encoding="utf-8") as handle:
        handle.write(render_html(summary, dates))
    print(f"rendered {len(summary)} rows over {len(dates)} day(s)")
    return 0


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    sub = parser.add_subparsers(dest="command", required=True)

    collect_parser = sub.add_parser(
        "collect", help="fetch counts and upsert them into the CSVs")
    collect_parser.add_argument("--data-dir", required=True)
    collect_parser.add_argument("--runs-json", help="gh run list --json output")
    collect_parser.add_argument("--api-fixture", help="offline payloads instead of the API")
    collect_parser.add_argument("--date", help="override the date, for testing")
    collect_parser.set_defaults(func=collect)

    render_parser = sub.add_parser("render", help="write report.md and report.html")
    render_parser.add_argument("--data-dir", required=True)
    render_parser.set_defaults(func=render)

    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
