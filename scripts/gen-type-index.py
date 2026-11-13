#!/usr/bin/env python3
"""Generate the NuGet README's "Every type in the package" index from the approved API (D46).

WHY GENERATED. D46 cut the package README - the page nuget.org renders - down to a front
page and moved its reference to docfx/guides/reference.md. A reader on nuget.org should
still see every public type the package ships, one click from its API page. A list of
102 type names written by hand goes stale on the first release that adds one; this one
is derived from tests/DocToolkit.Tests/PublicApi/*.approved.txt, the file the approval
tests hold to what actually ships, the same source gen-capability-matrix.py and
check-readme-coverage.py read.

WHAT IT WRITES, between BEGIN and END in src/DocToolkit/README.md:
  - one "### <group>" heading per approved file, in GROUPS order;
  - under it, one line: every public type in that file, sorted ordinally by name, each as
    [`Name`](API_URL/DocToolkit.Name.html), joined by ", ";
  - a generic type keeps its parameters in the text and uses docfx's arity file name:
    ConversionResult<T> -> [`ConversionResult<T>`](API_URL/DocToolkit.ConversionResult-1.html).

    python scripts/gen-type-index.py              # rewrite the block
    python scripts/gen-type-index.py --check      # exit 1 if the block is stale
    python scripts/gen-type-index.py --self-test
"""

import glob
import pathlib
import re
import sys

README = "src/DocToolkit/README.md"
APPROVED_GLOB = "tests/DocToolkit.Tests/PublicApi/DocToolkit*.approved.txt"
API_URL = "https://ank-khoaho.github.io/DocToolkit/api"
BEGIN = "<!-- BEGIN GENERATED (scripts/gen-type-index.py) - do not edit by hand -->"
END = "<!-- END GENERATED -->"

# approved file stem (without ".approved.txt") -> heading, in output order
GROUPS = [
    ("DocToolkit", "Core"),
    ("DocToolkit.Docx", "Word (DOCX)"),
    ("DocToolkit.Pdf", "PDF"),
    ("DocToolkit.Pptx", "PowerPoint (PPTX)"),
    ("DocToolkit.Xlsx", "Excel (XLSX)"),
    ("DocToolkit.Primitives", "Shared types"),
]

TYPE = re.compile(r"public (?:static |sealed |abstract |readonly )*(?:class|interface|enum|struct) (\w+(?:<[^>]+>)?)")


def types_in(approved_text):
    """Every public type declared in one approved file, in file order, generics with their parameters."""
    # Preserves declaration order while removing duplicates.
    return list(dict.fromkeys(TYPE.findall(approved_text)))


def link(type_name):
    """The Markdown link for one type name, e.g. ConversionResult<T> -> the -1.html page."""
    # Docfx uses arity file names for generic types.
    if "<" in type_name:
        base, params = type_name.split("<", 1)
        arity = params.rstrip(">").count(",") + 1
        page = f"DocToolkit.{base}-{arity}.html"
    else:
        page = f"DocToolkit.{type_name}.html"
    return f"[`{type_name}`]({API_URL}/{page})"


def render(approved_by_stem):
    """The whole block between BEGIN and END (exclusive), from {stem: approved text}. Raises KeyError
    when a GROUPS stem is missing, and ValueError when a stem is present that GROUPS does not name."""
    # Enforces exact stem correspondence with GROUPS.
    known = {stem for stem, _ in GROUPS}
    for stem in approved_by_stem:
        if stem not in known:
            raise ValueError(f"unknown approved stem: {stem}")
    for stem, _ in GROUPS:
        if stem not in approved_by_stem:
            raise KeyError(f"missing approved stem: {stem}")

    # Blank lines separate groups without leaving leading or trailing blanks in the block.
    sections = []
    for stem, heading in GROUPS:
        types = sorted(types_in(approved_by_stem[stem]))
        links = ", ".join(link(t) for t in types)
        sections.append(f"### {heading}\n\n{links}")
    return "\n\n".join(sections)


def replace_block(readme_text, block):
    """readme_text with the content between BEGIN and END replaced by block. Raises ValueError when
    either marker is missing."""
    # Validates marker presence and sequence before splicing the generated content.
    begin_idx = readme_text.find(BEGIN)
    if begin_idx == -1:
        raise ValueError(f"missing begin marker: {BEGIN}")
    end_idx = readme_text.find(END, begin_idx + len(BEGIN))
    if end_idx == -1:
        raise ValueError(f"missing end marker: {END}")

    before = readme_text[:begin_idx]
    after = readme_text[end_idx + len(END):]
    return f"{before}{BEGIN}\n{block}\n{END}{after}"


def is_stale(readme_text, approved_by_stem):
    """True when the block in readme_text is not what the approved files generate now."""
    return replace_block(readme_text, render(approved_by_stem)) != readme_text


# ---------------------------------------------------------------------------
# Self-test. Written before the implementation (D46, 2026-11-13); do not edit the
# cases to make them pass.

_APPROVED = {
    "DocToolkit": """namespace DocToolkit
{
    public sealed class ConversionResult<T>
    {
        public T Value { get; }
    }
    public static class HtmlToPdfConverter
    {
        public static byte[] Convert(string html) { }
    }
}""",
    "DocToolkit.Docx": """namespace DocToolkit
{
    public enum DocxFormKey
    {
        TagThenAlias = 0,
    }
    public static class DocxEditor
    {
        public static byte[] ReplaceText(byte[] docx) { }
    }
    public interface IBlock { }
}""",
    "DocToolkit.Pdf": "namespace DocToolkit\n{\n    public readonly struct PdfPoint { }\n}",
    "DocToolkit.Pptx": "namespace DocToolkit\n{\n    public static class PresentationEditor { }\n}",
    "DocToolkit.Xlsx": "namespace DocToolkit\n{\n    public abstract class Sheet { }\n}",
    "DocToolkit.Primitives": "namespace DocToolkit\n{\n    public sealed class PageSetup { }\n}",
}

_EXPECTED_DOCX_LINE = (
    f"[`DocxEditor`]({API_URL}/DocToolkit.DocxEditor.html), "
    f"[`DocxFormKey`]({API_URL}/DocToolkit.DocxFormKey.html), "
    f"[`IBlock`]({API_URL}/DocToolkit.IBlock.html)"
)


def self_test():
    bad = []

    def expect(name, ok):
        if not ok:
            bad.append(name)

    expect("generic keeps <T>", types_in(_APPROVED["DocToolkit"]) == ["ConversionResult<T>", "HtmlToPdfConverter"])
    expect("enum, class and interface all found", sorted(types_in(_APPROVED["DocToolkit.Docx"])) == ["DocxEditor", "DocxFormKey", "IBlock"])
    expect("readonly struct found", types_in(_APPROVED["DocToolkit.Pdf"]) == ["PdfPoint"])
    expect("generic links to the arity page",
           link("ConversionResult<T>") == f"[`ConversionResult<T>`]({API_URL}/DocToolkit.ConversionResult-1.html)")

    block = render(_APPROVED)
    headings = [l for l in block.splitlines() if l.startswith("### ")]
    expect("one heading per group, in GROUPS order", headings == [f"### {h}" for _, h in GROUPS])
    expect("a group's types sorted ordinally on one line", _EXPECTED_DOCX_LINE in block.splitlines())

    try:
        render({k: v for k, v in _APPROVED.items() if k != "DocToolkit.Pptx"})
        expect("a missing group raises KeyError", False)
    except KeyError:
        pass
    try:
        render({**_APPROVED, "DocToolkit.Html": "namespace DocToolkit\n{\n    public static class X { }\n}"})
        expect("an unknown approved file raises ValueError", False)
    except ValueError:
        pass

    page = f"# Title\n\nBefore.\n\n{BEGIN}\nstale\n{END}\n\nAfter.\n"
    out = replace_block(page, block)
    expect("only the block is replaced", out.startswith("# Title\n\nBefore.\n\n" + BEGIN) and out.endswith(END + "\n\nAfter.\n")
           and "stale" not in out and block in out)
    expect("replacing twice is stable", replace_block(out, block) == out)
    # --check is what CI runs, so its comparison is tested here rather than trusted (a planted
    # "never stale" bug survived the cases above, 2026-11-13).
    fresh = replace_block(page, render(_APPROVED))
    expect("a fresh page is not stale", is_stale(fresh, _APPROVED) is False)
    expect("a page with an old index is stale", is_stale(page, _APPROVED) is True)
    expect("a page missing one type is stale",
           is_stale(fresh.replace("[`IBlock`]", "[`IBlockX`]"), _APPROVED) is True)
    # The real README has the capability grid's own "<!-- END GENERATED -->" ABOVE this block.
    grid = "<!-- BEGIN GENERATED (scripts/gen-capability-matrix.py) - do not edit by hand -->\ngrid\n" + END
    two = f"{grid}\n\n{BEGIN}\nstale\n{END}\n"
    expect("an earlier END from another generator is left alone",
           replace_block(two, block) == f"{grid}\n\n{BEGIN}\n{block}\n{END}\n")
    try:
        replace_block("# no markers\n", block)
        expect("missing markers raise ValueError", False)
    except ValueError:
        pass

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
    approved = {pathlib.Path(p).name.removesuffix(".approved.txt"): open(p, encoding="utf-8").read()
                for p in sorted(glob.glob(APPROVED_GLOB))}
    path = pathlib.Path(README)
    raw = path.read_bytes().decode("utf-8")
    nl = "\r\n" if "\r\n" in raw else "\n"
    text = raw.replace("\r\n", "\n")
    if "--check" in argv:
        if is_stale(text, approved):
            print(f"::error file={README}::the type index is stale; run python scripts/gen-type-index.py")
            return 1
        print("type index is up to date")
        return 0
    updated = replace_block(text, render(approved))
    path.write_bytes(updated.replace("\n", nl).encode("utf-8"))
    print(f"wrote the type index to {README}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
