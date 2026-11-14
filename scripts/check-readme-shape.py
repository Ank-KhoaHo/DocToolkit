#!/usr/bin/env python3
"""Fail when a README front page grows long again (D46).

WHY THIS EXISTS. On 2026-11-13 the package README - the page nuget.org renders - was
1,730 lines, averaged 24.8 words a sentence, and held 137 sentences over 25 words. The
reference text in it was correct; it was simply the first thing a visitor met. D46 moved
that reference to docfx/guides/reference.md and gave both READMEs a short front page.
Nothing stops a front page from growing back one helpful paragraph at a time, so this
holds the shape that was measured and chosen:

  - README.md (GitHub) at most ROOT_MAX_LINES lines;
  - src/DocToolkit/README.md (nuget.org) at most NUGET_MAX_LINES lines;
  - in each front page, an average of at most AVG_MAX words a sentence and no sentence
    over SENTENCE_MAX words;
  - in the nuget.org front page, no HTML tag and no relative image path, because
    nuget.org renders CommonMark only and allow-listed absolute image URLs only - both
    would show on GitHub and render as nothing on nuget.org;
  - in each front page, the four sections every public README here keeps (D48): Status,
    Limits, Alternatives and FAQ, as level-2 headings outside code blocks;
  - in each, a first paragraph of prose (6 words or more) before the first level-2 heading,
    saying what this is before any instruction.

ROOT_MAX_LINES was 160 for D46's front page; D48 raised it to 200, the nuget.org cap, for
the three sections above. The public template allows 300. Raise it again only on purpose.

The nuget.org front page is everything above TYPE_INDEX_MARKER; the generated type index
below it is a list, not prose.

HOW PROSE IS READ (the same rules as the D46 measurement):
  - removed: fenced code blocks, generated and snippet blocks (BEGIN ... END markers),
    table rows (lines starting with '|'), headings, image lines, HTML-tag lines, comments;
  - a blank line and each list item start a new block, so a bullet with no full stop never
    runs into the next one;
  - inside a block, a sentence ends at '.', '!' or '?' followed by whitespace and then an
    uppercase letter, or at the end of the block - so "0.60.1", "net8.0", ".NET 10" and
    "e.g. a" do not end one;
  - a sentence counts when it has 3 or more words; words are whitespace-separated.

    python scripts/check-readme-shape.py --self-test
    python scripts/check-readme-shape.py
"""

import sys

ROOT = "README.md"
NUGET = "src/DocToolkit/README.md"
TYPE_INDEX_MARKER = "<!-- BEGIN GENERATED (scripts/gen-type-index.py) - do not edit by hand -->"
ROOT_MAX_LINES = 200
NUGET_MAX_LINES = 200
AVG_MAX = 15
SENTENCE_MAX = 25
REQUIRED_SECTIONS = ("Status", "Limits", "Alternatives", "FAQ")
FIRST_PARAGRAPH_MIN_WORDS = 6


def headings(text):
    """Level-2 heading texts, in order, skipping fenced code blocks."""
    found, in_fence = [], False
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith("```"):
            in_fence = not in_fence
        elif not in_fence and line.startswith("## "):
            found.append(line[3:].strip())
    return found


def opening_words(text):
    """Words of prose before the first level-2 heading: badges, images and HTML count as none."""
    opening = []
    for line in text.splitlines():
        if line.startswith("## "):
            break
        opening.append(line)
    return sum(len(sentence.split()) for _, sentence in prose("\n".join(opening)))


def prose(text):
    """Return [(line_number, sentence)] for every counted sentence, line numbers 1-based."""
    import re

    counted = []
    current_block = []

    def flush_block():
        nonlocal current_block
        if not current_block:
            return

        # Line numbers are tracked per character to locate the line where each sentence begins.
        block_text = ""
        char_lineno = []
        for lno, cline in current_block:
            if block_text:
                block_text += "\n"
                char_lineno.append(lno)
            block_text += cline
            char_lineno.extend([lno] * len(cline))

        current_block = []

        if not block_text.strip():
            return

        # Sentences terminate at punctuation followed by capital letters or at the block boundary.
        split_pattern = re.compile(r"(?<=[.!?])\s+(?=[A-Z])")
        prev_end = 0
        for m in split_pattern.finditer(block_text):
            seg = block_text[prev_end : m.start()]
            words = seg.split()
            if len(words) >= 3:
                rel_start = len(seg) - len(seg.lstrip())
                lno = char_lineno[prev_end + rel_start]
                counted.append((lno, " ".join(words)))
            prev_end = m.end()

        seg = block_text[prev_end:]
        words = seg.split()
        if len(words) >= 3:
            rel_start = len(seg) - len(seg.lstrip())
            lno = char_lineno[prev_end + rel_start]
            counted.append((lno, " ".join(words)))

    in_fence = False
    in_marker = False

    for lineno, raw_line in enumerate(text.splitlines(), 1):
        stripped = raw_line.strip()

        if in_fence:
            if stripped.startswith("```"):
                in_fence = False
            continue

        if stripped.startswith("```"):
            in_fence = True
            flush_block()
            continue

        if in_marker:
            if stripped.startswith("<!-- END"):
                in_marker = False
            continue

        if stripped.startswith("<!-- BEGIN "):
            in_marker = True
            flush_block()
            continue

        # Non-prose lines (tables, headings, images, tags, badge lines) end blocks.
        if stripped.startswith(("|", "#", "![", "<", "[![")):
            flush_block()
            continue

        if not stripped:
            flush_block()
            continue

        # Lines starting with list markers start a new block so bullets never merge.
        if stripped.startswith(("- ", "* ")) or re.match(r"^\d+\.\s+", stripped):
            flush_block()
            raw_line = re.sub(r"^\s*(?:[-*]|\d+\.)\s+", "", raw_line)

        # Markdown syntax stripping normalises prose for sentence splitting and word counting.
        cleaned = re.sub(r"<[^>]+>", " ", raw_line)
        cleaned = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", cleaned)
        cleaned = cleaned.replace("**", "").replace("*", "").replace("`", "")

        current_block.append((lineno, cleaned))

    flush_block()
    return counted


def check_sections(text):
    """Return error lines for the sections and the opening (D48). Empty means it passes."""
    errors = []
    found = headings(text)
    missing = [name for name in REQUIRED_SECTIONS if name not in found]
    if missing:
        errors.append(f"no '## {missing[0]}' section; every public README keeps "
                      f"{', '.join(REQUIRED_SECTIONS)} (missing: {', '.join(missing)})")
    words = opening_words(text)
    if words < FIRST_PARAGRAPH_MIN_WORDS:
        errors.append(f"the prose before the first '## ' heading has {words} words "
                      f"(at least {FIRST_PARAGRAPH_MIN_WORDS}): say what this is before any instruction")
    return errors


def check(text, kind):
    """Return error lines for one README; kind is "root" or "nuget". Empty means it passes."""
    import re

    errors = []
    lines = text.splitlines()

    max_lines = ROOT_MAX_LINES if kind == "root" else NUGET_MAX_LINES
    if len(lines) > max_lines:
        errors.append(f"file has {len(lines)} lines (max {max_lines})")

    # The NuGet front page excludes the generated type index below the marker.
    if kind == "nuget":
        marker_idx = None
        for i, line in enumerate(lines):
            if line.strip() == TYPE_INDEX_MARKER.strip():
                marker_idx = i
                break
        if marker_idx is None:
            errors.append("type index marker missing")
            front_lines = []
        else:
            front_lines = lines[:marker_idx]
    else:
        front_lines = lines

    # NuGet front pages disallow raw HTML tags and relative images to guarantee valid CommonMark rendering.
    if kind == "nuget":
        in_fence = False
        for lineno, line in enumerate(front_lines, 1):
            # Code is shown, not rendered: "<h1>" in a C# string is a string, so fenced blocks and
            # inline code spans are skipped (found when the real quick-start snippet failed).
            if line.strip().startswith("```"):
                in_fence = not in_fence
                continue
            if in_fence:
                continue
            tag_match = re.search(r"<[A-Za-z/][^>]*>", re.sub(r"`[^`]*`", "", line))
            if tag_match:
                errors.append(f"line {lineno}: HTML tag not allowed on nuget.org: {tag_match.group(0)}")
            for img_match in re.finditer(r"!\[[^\]]*\]\(([^)]+)\)", line):
                url = img_match.group(1).strip()
                if not (url.startswith("http://") or url.startswith("https://")):
                    errors.append(f"line {lineno}: relative image URL not allowed on nuget.org: {url}")

    counted = prose("\n".join(front_lines))
    for lineno, s in counted:
        words = s.split()
        if len(words) > SENTENCE_MAX:
            first_8 = " ".join(words[:8])
            errors.append(f"line {lineno}: sentence has {len(words)} words (max {SENTENCE_MAX}): '{first_8}...'")

    if counted:
        avg = sum(len(s.split()) for _, s in counted) / len(counted)
        if avg > AVG_MAX:
            errors.append(f"average sentence length is {avg:.1f} words (max {AVG_MAX})")

    return errors


# ---------------------------------------------------------------------------
# Self-test. Written before the implementation (D46, 2026-11-13); do not edit the
# cases to make them pass.

def _words(n, word="word"):
    return " ".join([word] * n)


_SHORT = "This is short. It stays short. Nothing here is long."
_VERSIONS = "Version 0.60.1 targets net8.0 on .NET 10, e.g. in a container."  # 11 words, one sentence
_LONG26 = "Start " + _words(24) + " end."      # 26 words
_AVG16 = " ".join(["One " + _words(15) + "."] * 4)  # four sentences of 16 words

_CASES = [
    ("short sentences pass", "# T\n\n" + _SHORT + "\n", "root", True),
    ("versions and .NET do not end a sentence", _VERSIONS + "\n", "root", True),
    ("a 26-word sentence fails", _SHORT + "\n\n" + _LONG26 + "\n", "root", False),
    ("an average of 16 fails", _AVG16 + "\n", "root", False),
    ("code, tables, HTML and headings are not prose",
     _SHORT + "\n\n```csharp\n" + _words(40) + "\n```\n\n| " + _words(40) + " |\n\n"
     "<p align=\"center\">" + _words(40) + "</p>\n\n## " + _words(40) + "\n", "root", True),
    ("a bullet without a full stop does not run into the next",
     # 14 words each; run together they would be one 28-word sentence and fail.
     "- **No browser** " + _words(12) + "\n- **No binaries** " + _words(12) + "\n", "root", True),
    # D48 raised the root cap from 160 to 200 on purpose; these two cases moved with it.
    ("201 lines fail for the root README", "\n".join(["Short line here."] * 201) + "\n", "root", False),
    ("200 lines pass for the root README", "\n".join(["Short line here."] * 200) + "\n", "root", True),
    ("HTML above the marker fails on nuget",
     "<p align=\"center\">x</p>\n\n" + _SHORT + "\n\n" + TYPE_INDEX_MARKER + "\n", "nuget", False),
    ("a relative image fails on nuget",
     "![hero](assets/art/hero.png)\n\n" + _SHORT + "\n\n" + TYPE_INDEX_MARKER + "\n", "nuget", False),
    ("an absolute raw.githubusercontent image passes on nuget",
     "![hero](https://raw.githubusercontent.com/Ank-KhoaHo/DocToolkit/main/assets/art/hero.png)\n\n"
     + _SHORT + "\n\n" + TYPE_INDEX_MARKER + "\n", "nuget", True),
    ("HTML inside a code block is code, not markup, on nuget",
     _SHORT + "\n\n```csharp\nvar pdf = Convert(\"<h1>Invoice</h1>\");\n```\n\n" + TYPE_INDEX_MARKER + "\n",
     "nuget", True),
    ("a long sentence below the marker is ignored on nuget",
     _SHORT + "\n\n" + TYPE_INDEX_MARKER + "\n" + _LONG26 + "\n", "nuget", True),
    ("the root README may use HTML",
     "<p align=\"center\"><img src=\"assets/icon.png\"></p>\n\n" + _SHORT + "\n", "root", True),
    ("201 lines fail for nuget",
     "\n".join(["Short line here."] * 200) + "\n" + TYPE_INDEX_MARKER + "\n", "nuget", False),
]

# D48: the sections every public README keeps, and an opening that says what this is.
_OPENING = "# T\n\nDocToolkit turns HTML into Word and PDF files.\n\n"
_ALL_FOUR = "## Status\n\nx\n\n## Limits\n\nx\n\n## Alternatives\n\nx\n\n## FAQ\n\nx\n"

_SECTION_CASES = [
    ("an opening and all four sections pass", _OPENING + _ALL_FOUR, True),
    ("a missing Limits section fails", _OPENING + _ALL_FOUR.replace("## Limits", "## Caveats"), False),
    ("a level-3 heading is not a section", _OPENING + _ALL_FOUR.replace("## Status", "### Status"), False),
    ("a heading inside a code block is not a section",
     _OPENING + "```\n## Status\n```\n\n" + _ALL_FOUR.replace("## Status\n\nx\n\n", ""), False),
    ("a heading that only starts with the name is not the section",
     _OPENING + _ALL_FOUR.replace("## FAQ", "## FAQs and more"), False),
    ("badges and an image alone are no opening",
     "# T\n\n[![b](https://x/b.svg)](https://x)\n\n![hero](https://x/h.png)\n\n" + _ALL_FOUR, False),
    ("HTML alone is no opening", "<h1 align=\"center\">T</h1>\n\n<p>" + _words(9) + "</p>\n\n" + _ALL_FOUR, False),
    ("a five-word opening is too short", "# T\n\nIt makes Word files fast.\n\n" + _ALL_FOUR, False),
]


def self_test():
    bad = 0
    for name, text, kind, expect_ok in _CASES:
        errors = check(text, kind)
        if (not errors) != expect_ok:
            bad += 1
            print(f"::error::self-test case '{name}': expected {'pass' if expect_ok else 'fail'}, "
                  f"got {errors or 'pass'}")
    for name, text, expect_ok in _SECTION_CASES:
        errors = check_sections(text)
        if (not errors) != expect_ok:
            bad += 1
            print(f"::error::self-test case '{name}': expected {'pass' if expect_ok else 'fail'}, "
                  f"got {errors or 'pass'}")
    long_line = [n for n, s in prose("One.\n\n" + _LONG26 + "\n") if len(s.split()) > SENTENCE_MAX]
    if long_line != [3]:
        bad += 1
        print(f"::error::self-test: the 26-word sentence should be reported on line 3, got {long_line}")
    if bad:
        print(f"::error::self-test failed {bad} case(s)")
        return 1
    print("self-test passed")
    return 0


def main(argv):
    if "--self-test" in argv:
        return self_test()
    failed = False
    for path, kind in ((ROOT, "root"), (NUGET, "nuget")):
        text = open(path, encoding="utf-8").read()
        errors = check(text, kind) + check_sections(text)
        for line in errors:
            print(f"::error file={path}::{line}")
        failed |= bool(errors)
        if not errors:
            print(f"{path}: front page within shape")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
