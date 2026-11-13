#!/usr/bin/env python3
"""Fail when a card on the docs landing page shows a call the library does not have.

WHY THIS SCRIPT HAS TO EXIST. The landing page (docfx/index.md) shows six task
cards, and each card shows its one call as text, for example
`DocxToPdfConverter.Convert(docx)`. Every other piece of code in the docs is a
reference into samples/, which CI compiles, so it cannot drift from the API
(see check-doc-snippets.py). Card text is not compiled. Rename a method and
the landing page - the first page a visitor reads - goes on advertising a call
that no longer exists, while every build and every other check stays green.

WHAT IT CHECKS, against tests/DocToolkit.Tests/PublicApi/*.approved.txt (the
API file the approval tests hold to what actually ships):
  - the class exists
  - the method exists on that class
  - some overload accepts that many arguments (required <= n <= total)
  - at least MIN_CALLS calls were found, so the check cannot pass by matching
    nothing

It reads only the cards, between LANDING_START and LANDING_END: other inline
code on the page (`services.AddDocToolkit()` under "Two packages") is prose,
not a card call.

    python scripts/check-landing-calls.py --self-test
    python scripts/check-landing-calls.py
"""

import glob
import re
import sys

INDEX = "docfx/index.md"
APPROVED_GLOB = "tests/DocToolkit.Tests/PublicApi/*.approved.txt"
LANDING_START = "<!-- landing-cards:start -->"
LANDING_END = "<!-- landing-cards:end -->"
MIN_CALLS = 6


def split_params(param_str):
    param_str = param_str.strip()
    if not param_str:
        return []
    params = []
    current = []
    angle = 0
    paren = 0
    bracket = 0
    for ch in param_str:
        if ch == '<':
            angle += 1
        elif ch == '>':
            if angle > 0:
                angle -= 1
        elif ch == '(':
            paren += 1
        elif ch == ')':
            if paren > 0:
                paren -= 1
        elif ch == '[':
            bracket += 1
        elif ch == ']':
            if bracket > 0:
                bracket -= 1
        elif ch == ',' and angle == 0 and paren == 0 and bracket == 0:
            params.append("".join(current).strip())
            current = []
            continue
        current.append(ch)
    if current:
        params.append("".join(current).strip())
    return params


def split_args(arg_str):
    arg_str = arg_str.strip()
    if not arg_str:
        return []
    args = []
    current = []
    in_quote = False
    quote_char = None
    escape = False
    depth = 0

    for ch in arg_str:
        if escape:
            current.append(ch)
            escape = False
            continue
        if ch == '\\' and in_quote:
            current.append(ch)
            escape = True
            continue
        if in_quote:
            current.append(ch)
            if ch == quote_char:
                in_quote = False
                quote_char = None
            continue
        if ch in ('"', "'"):
            in_quote = True
            quote_char = ch
            current.append(ch)
            continue
        if ch in ('(', '[', '{', '<'):
            depth += 1
            current.append(ch)
            continue
        if ch in (')', ']', '}', '>'):
            if depth > 0:
                depth -= 1
            current.append(ch)
            continue
        if ch == ',' and depth == 0:
            args.append("".join(current).strip())
            current = []
            continue
        current.append(ch)

    if current:
        args.append("".join(current).strip())
    return args


def load_api(texts):
    """Map class name -> method name -> list of (required, total) per overload.

    `texts` are the contents of approved files. A parameter with a default
    (`= ...`) is optional. Parameters split on commas outside <>, () and [].
    """
    api = {}
    for text in texts:
        current_class = None
        in_class = False
        class_depth = None
        depth = 0
        for line in text.splitlines():
            stripped = line.strip()

            m_cls = re.match(r"^public\s+static\s+class\s+([A-Za-z0-9_]+)", stripped)
            if m_cls:
                current_class = m_cls.group(1)
                api.setdefault(current_class, {})
                in_class = False
                class_depth = None
            elif re.match(r"^public\s+(?:sealed\s+|abstract\s+)?class\s+", stripped):
                current_class = None
                in_class = False

            if current_class is not None and not in_class:
                if '{' in stripped:
                    in_class = True
                    class_depth = depth + stripped.count('{')

            if current_class is not None and in_class:
                m_mth = re.match(r"^public\s+static\s+(.+)\s+([A-Za-z0-9_]+)(?:<[^>]+>)?\s*\((.*)\)\s*\{", stripped)
                if m_mth:
                    method_name = m_mth.group(2)
                    raw_params = m_mth.group(3)
                    params = split_params(raw_params)
                    total = len(params)
                    required = sum(1 for p in params if '=' not in p)
                    api[current_class].setdefault(method_name, []).append((required, total))

            depth += stripped.count('{') - stripped.count('}')
            if in_class and class_depth is not None and depth < class_depth:
                current_class = None
                in_class = False
                class_depth = None

    return api


def find_calls(index_text):
    """Return [(class, method, arg_count)] for each card call between the markers.

    A card call is the whole text of a <code> element matching Class.Method(args).
    Arguments split on top-level commas only: a string literal or a nested
    bracket is one argument. `Class.Method()` has 0 arguments.
    Raises ValueError when either marker is missing.
    """
    if LANDING_START not in index_text or LANDING_END not in index_text:
        raise ValueError(f"Missing markers: {LANDING_START} and/or {LANDING_END}")
    start_idx = index_text.index(LANDING_START) + len(LANDING_START)
    if LANDING_END not in index_text[start_idx:]:
        raise ValueError("Landing end marker occurs before start marker")
    end_idx = index_text.index(LANDING_END, start_idx)
    section = index_text[start_idx:end_idx]

    calls = []
    code_matches = re.findall(r"<code\b[^>]*>(.*?)</code>", section, re.DOTALL)
    for raw_code in code_matches:
        code = raw_code.strip()
        m = re.match(r"^([A-Za-z0-9_]+)\.([A-Za-z0-9_]+)\((.*)\)$", code, re.DOTALL)
        if m:
            cls = m.group(1)
            method = m.group(2)
            raw_args = m.group(3).strip()
            args = split_args(raw_args)
            calls.append((cls, method, len(args)))
    return calls


def check(api, calls):
    """Return error lines; empty when every call resolves and len(calls) >= MIN_CALLS."""
    errors = []
    if len(calls) < MIN_CALLS:
        errors.append(f"Found {len(calls)} card call(s), expected at least {MIN_CALLS}")
    for cls, method, n in calls:
        if cls not in api:
            errors.append(f"Unknown class '{cls}'")
            continue
        if method not in api[cls]:
            errors.append(f"Class '{cls}' has no method '{method}'")
            continue
        overloads = api[cls][method]
        if not any(req <= n <= tot for req, tot in overloads):
            errors.append(f"No overload of '{cls}.{method}' accepts {n} argument(s)")
    return errors


# ---------------------------------------------------------------------------
# Self-test. Written before the implementation (D45, 2026-11-12); do not edit
# the cases to make them pass.

_API = """
    public static class DocxEditor
    {
        public static byte[] ReplaceText(byte[] docx, System.Collections.Generic.IReadOnlyDictionary<string, string> replacements) { }
        public static byte[] ReplaceText(byte[] docx, System.Text.RegularExpressions.Regex pattern, string replacement) { }
    }
    public static class Conv
    {
        public static byte[] Convert(byte[] a, DocToolkit.PdfFontOptions? fonts = null) { }
        public static System.Threading.Tasks.Task<byte[]> ConvertAsync(string html, System.Threading.CancellationToken ct = default) { }
    }
"""


def _page(*calls, outside=""):
    cards = "\n".join(f'<div class="col"><a href="x.md"><code class="small">{c}</code></a></div>' for c in calls)
    return f"# Page\n\n{outside}\n\n{LANDING_START}\n{cards}\n{LANDING_END}\n"


_GOOD = [
    "DocxEditor.ReplaceText(docx, values)",
    'DocxEditor.ReplaceText(docx, new Regex("a, b"), "c")',
    "DocxEditor.ReplaceText(docx, Map(a, b))",
    "Conv.Convert(a)",
    "Conv.Convert(a, fonts)",
    "Conv.ConvertAsync(html)",
]

_CASES = [
    ("six good calls", _page(*_GOOD), True),
    ("unknown class", _page(*_GOOD[:5], "Nope.Convert(a)"), False),
    ("unknown method", _page(*_GOOD[:5], "DocxEditor.Replace(docx, values)"), False),
    ("too few arguments", _page(*_GOOD[:5], "DocxEditor.ReplaceText(docx)"), False),
    ("too many arguments", _page(*_GOOD[:5], "Conv.Convert(a, b, c)"), False),
    ("fewer than six calls", _page(*_GOOD[:5]), False),
    ("a comma inside a string is one argument", _page(*_GOOD[:5], 'Conv.Convert("a, b", fonts)'), True),
    ("code outside the markers is ignored",
     _page(*_GOOD, outside="<code>services.AddDocToolkit()</code>"), True),
]


def self_test():
    api = load_api([_API])
    bad = 0
    for name, page, expect_ok in _CASES:
        errors = check(api, find_calls(page))
        if (not errors) != expect_ok:
            bad += 1
            print(f"::error::self-test case '{name}': expected {'pass' if expect_ok else 'fail'}, got {errors or 'pass'}")
    try:
        find_calls("# no markers\n<code>Conv.Convert(a)</code>\n")
        bad += 1
        print("::error::self-test case 'markers missing': expected ValueError")
    except ValueError:
        pass
    if bad:
        print(f"::error::self-test failed {bad} case(s)")
        return 1
    print("self-test passed")
    return 0


def main(argv):
    if "--self-test" in argv:
        return self_test()
    paths = sorted(glob.glob(APPROVED_GLOB))
    if not paths:
        sys.exit(f"::error::No approved API files at {APPROVED_GLOB}; this check cannot run.")
    api = load_api([open(p, encoding="utf-8").read() for p in paths])
    try:
        calls = find_calls(open(INDEX, encoding="utf-8").read())
    except ValueError as e:
        sys.exit(f"::error::{INDEX}: {e}")
    for cls, method, n in calls:
        print(f"  {cls}.{method} with {n} argument(s)")
    errors = check(api, calls)
    for line in errors:
        print(f"::error::{line}")
    print(f"{len(calls)} card call(s) checked against {len(paths)} approved API file(s)")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
