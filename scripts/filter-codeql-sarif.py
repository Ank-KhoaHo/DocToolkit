"""Drop CodeQL results that sit in generated build output (bin/, obj/) before the SARIF is uploaded (C47).

WHY THIS EXISTS. .github/codeql-config.yml lists `paths-ignore: **/obj/**`, and for C# with
`build-mode: manual` that line does nothing: GitHub applies paths/paths-ignore only to languages it
analyses without building; for a built compiled language "you must specify appropriate build steps"
(GitHub docs, customizing advanced setup, read 2026-11-14). The build cannot leave obj/ out - the
[GeneratedRegex] source generator writes its code there and the library needs it - so two alerts
landed in RegexGenerator.g.cs on 2026-11-08, in code nobody writes and every build regenerates.
codeql.yml now analyses with `upload: failure-only`, runs this on the SARIF, and uploads the result.

A result is generated when its first location's path has a whole `bin` or `obj` segment.
Whole segments only: `src/objects/` and `obj.cs` are real code.

    python scripts/filter-codeql-sarif.py <file.sarif> [...]   # rewrites each file in place
    python scripts/filter-codeql-sarif.py --self-test
"""
import json
import os
import sys
import tempfile

GENERATED_SEGMENTS = {"bin", "obj"}


def is_generated(uri):
    """True when `uri` (relative, absolute file:// or Windows-style) has a whole bin/ or obj/ segment."""
    if not uri or not isinstance(uri, str):
        return False
    normalized = uri.replace("\\", "/")
    return any(seg in GENERATED_SEGMENTS for seg in normalized.split("/"))


def drop_generated(sarif):
    """Remove generated results from every run of a parsed SARIF dict, in place, keeping order.

    A result with no locations is kept. Returns the number removed."""
    removed = 0
    for run in sarif.get("runs", []):
        if not isinstance(run, dict):
            continue
        results = run.get("results")
        if not isinstance(results, list):
            continue
        kept = []
        for result in results:
            if not isinstance(result, dict):
                kept.append(result)
                continue
            locations = result.get("locations")
            if not locations or not isinstance(locations, list):
                kept.append(result)
                continue
            first = locations[0]
            uri = None
            if isinstance(first, dict):
                phys = first.get("physicalLocation")
                if isinstance(phys, dict):
                    art = phys.get("artifactLocation")
                    if isinstance(art, dict):
                        uri = art.get("uri")
            if uri and is_generated(uri):
                removed += 1
            else:
                kept.append(result)
        run["results"] = kept
    return removed


def filter_file(path):
    """Read the SARIF at `path`, drop generated results, write it back as UTF-8 JSON. Returns the count."""
    with open(path, "r", encoding="utf-8") as f:
        sarif = json.load(f)
    count = drop_generated(sarif)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(sarif, f)
    return count


def _result(uri):
    return {"ruleId": "cs/x", "locations": [{"physicalLocation": {"artifactLocation": {"uri": uri}}}]}


def self_test():
    bad = []

    def expect(name, ok):
        if not ok:
            bad.append(name)

    expect("obj/ under a project is generated",
           is_generated("src/DocToolkit.Primitives/obj/Release/net10.0/generated/RegexGenerator.g.cs") is True)
    expect("bin/ is generated", is_generated("src/DocToolkit/bin/Release/net8.0/x.cs") is True)
    expect("a top-level obj/ is generated", is_generated("obj/x.cs") is True)
    expect("a folder that starts with obj is real code", is_generated("src/objects/x.cs") is False)
    expect("a folder that ends with obj is real code", is_generated("src/Robj/x.cs") is False)
    expect("a file named obj.cs is real code", is_generated("src/DocToolkit/obj.cs") is False)
    expect("a folder named binary is real code", is_generated("src/binary/x.cs") is False)
    expect("plain source is real code", is_generated("src/DocToolkit/DocxToPdfConverter.cs") is False)
    expect("backslashes are path separators", is_generated("src\\DocToolkit\\obj\\x.cs") is True)
    expect("an absolute file:// uri works",
           is_generated("file:///home/runner/work/DocToolkit/DocToolkit/src/A/obj/x.cs") is True)

    sarif = {"runs": [
        {"results": [_result("src/A/obj/Release/x.g.cs"), _result("src/A/B.cs"), {"ruleId": "cs/nolocation"}]},
        {"results": [_result("tests/T/Test.cs"), _result("src/A/bin/y.cs")]},
        {"tool": {}},
    ]}
    removed = drop_generated(sarif)
    expect("two generated results removed", removed == 2)
    expect("first run keeps real code and the location-less result, in order",
           sarif["runs"][0]["results"] == [_result("src/A/B.cs"), {"ruleId": "cs/nolocation"}])
    expect("second run keeps its real result", sarif["runs"][1]["results"] == [_result("tests/T/Test.cs")])
    expect("a run without results is left alone", sarif["runs"][2] == {"tool": {}})
    expect("a clean SARIF loses nothing", drop_generated({"runs": [{"results": [_result("src/A/B.cs")]}]}) == 0)

    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "csharp.sarif")
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"version": "2.1.0", "runs": [{"results": [_result("src/A/obj/x.cs"), _result("src/A/B.cs")]}]}, f)
        count = filter_file(path)
        with open(path, encoding="utf-8") as f:
            back = json.load(f)
        expect("filter_file reports the count", count == 1)
        expect("filter_file writes the filtered SARIF back",
               back == {"version": "2.1.0", "runs": [{"results": [_result("src/A/B.cs")]}]})

    for name in bad:
        print(f"::error::self-test case '{name}' failed")
    if bad:
        print(f"::error::self-test failed {len(bad)} case(s)")
        return 1
    print("self-test passed")
    return 0


def main(argv):
    if "--self-test" in argv:
        return self_test()
    if not argv:
        print("usage: filter-codeql-sarif.py <file.sarif> [...] | --self-test")
        return 2
    for path in argv:
        print(f"{path}: dropped {filter_file(path)} result(s) in bin/ or obj/")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
