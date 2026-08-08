#!/usr/bin/env python3
"""Record this package's nuget.org download counts daily, so that real adoption
becomes visible against the noise our own CI generates.

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
"""

import argparse
import collections
import csv
import datetime
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
    with open(path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(header))
        writer.writeheader()
        writer.writerows(rows)


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
                except (urllib.error.URLError, OSError, *readable) as error:
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

    args = parser.parse_args()
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
