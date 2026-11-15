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
    saying what this is before any instruction;
  - in each, the one shape (D50, below).

THE ONE SHAPE (D50). D48 said which sections exist. It fixed no names and no order, so a
section moved, renamed or added at the end still passed. A reader who has seen one README
here should know where to look in the next, so the names and their order are held too:

  - the banner (BANNER) is the first thing in the file; then one '# Name' line, and under
    it a badge line and one line in italics;
  - the first section is 'What it is', then 'Get started', then the optional 'How to use it';
  - the project's own sections come next, in one slot ('Numbers' and 'Proof' today);
  - then 'Status', 'Limits', 'Alternatives', 'FAQ', 'Documentation', 'Licence', in that order;
  - a section of the project's own after 'Status' fails: its slot is before it;
  - no section is empty.

For the order the nuget.org page is read whole, not only its front page: 'Licence' sits
below the generated type list. That list must add no level-2 heading of its own.

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
BANNER = "assets/art/banner.png"
OPENING_SECTIONS = ("What it is", "Get started", "How to use it")
CLOSING_SECTIONS = ("Status", "Limits", "Alternatives", "FAQ", "Documentation", "Licence")
ORDER_REQUIRED = ("What it is", "Get started", "Documentation", "Licence")


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


def outline(text):
    """Return (top, sections) for the whole file.

    top is [(line, in_code)]: every line above the first level-2 heading. sections is
    [(name, body_lines)]. A heading inside a ``` or ~~~ block is code, not a section.
    """
    top, sections, fence = [], [], ""
    for line in text.splitlines():
        mark = line.strip()[:3]
        if fence:
            in_code = True
            if mark == fence:
                fence = ""
        elif mark in ("```", "~~~"):
            in_code = True
            fence = mark
        else:
            in_code = False
        if not in_code and line.startswith("## "):
            sections.append((line[3:].strip(), []))
        elif sections:
            sections[-1][1].append(line)
        else:
            top.append((line, in_code))
    return top, sections


def _is_italic_line(line):
    """True for '*words*', false for '**words**'."""
    return len(line) > 3 and line[0] == "*" and line[-1] == "*" and line[1] != "*" and line[-2] != "*"


def check_order(text):
    """Return error lines for the one shape (D50). Empty means it passes.

    The whole file is read, not only the front page: on the nuget.org page 'Licence' sits
    below the generated type list.
    """
    errors = []
    top, sections = outline(text)
    names = [name for name, _ in sections]

    # The top of the page: the banner, then the name, then badges and one line in italics.
    shown = [line.strip() for line, _ in top if line.strip()]
    first = shown[0] if shown else ""
    if not ((first.startswith("![") or "<img" in first) and BANNER in first):
        errors.append(f"the banner ({BANNER}) is not the first thing in the file")

    written = [line.strip() for line, in_code in top if line.strip() and not in_code]
    titles = [i for i, line in enumerate(written) if line.startswith("# ")]
    if len(titles) != 1:
        errors.append(f"expected one '# Name' line above the first section, found {len(titles)}")
    else:
        under = written[titles[0] + 1:]
        if not any(_is_italic_line(line) for line in under):
            errors.append("no one-line description in italics under the name")
        if not any("<img" in line or "[![" in line for line in under):
            errors.append("no badge line under the name")

    # The names. check_sections already reports Status, Limits, Alternatives and FAQ.
    fixed = OPENING_SECTIONS + CLOSING_SECTIONS
    for name in ORDER_REQUIRED:
        if name not in names:
            errors.append(f"no '## {name}' section; the fixed names, in order: {', '.join(fixed)}")

    # The order. A section of the project's own has one slot: before 'Status'.
    highest, own_seen, closing_seen = -1, False, False
    for name in names:
        if name in fixed:
            position = fixed.index(name)
            if position < highest:
                errors.append(f"'{name}' is out of order; the order is: {', '.join(fixed)}")
            elif name in OPENING_SECTIONS and own_seen:
                errors.append(f"'{name}' comes after the project's own sections: it belongs before them")
            highest = max(highest, position)
            closing_seen = closing_seen or name in CLOSING_SECTIONS
        elif closing_seen:
            errors.append(f"'{name}' is a section of the project's own, after 'Status': "
                          "its one slot is before 'Status'")
        else:
            own_seen = True

    if names and names[0] != "What it is":
        errors.append(f"the first section is '{names[0]}', not 'What it is'")

    for name, body in sections:
        if not any(line.strip() for line in body):
            errors.append(f"'{name}' has nothing under it")

    return errors


def check_page(text, kind):
    """Every rule for one README in one list: this is what the real run prints."""
    return check(text, kind) + check_sections(text) + check_order(text)


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

# D50: one top of the page, and one fixed order for the sections. Written before check_order.
# Each failing case lists the error lines it must get, and it must get no other: a case that
# fails for another reason proves nothing about the rule it is named for.
_BANNER_LINE = "![A page it made](" + BANNER + ")"
_BADGE_LINE = "[![b](https://x/b.svg)](https://x)"
_ITALIC_LINE = "*HTML in, Word and PDF out.*"
_TOP = _BANNER_LINE + "\n\n# T\n\n" + _BADGE_LINE + "\n\n" + _ITALIC_LINE + "\n\n"
_ONE_SHAPE = ("What it is", "Get started", "Numbers", "Status", "Limits", "Alternatives", "FAQ",
              "Documentation", "Licence")
_NO_OWN = tuple(name for name in _ONE_SHAPE if name != "Numbers")


def _page(names=_ONE_SHAPE, top=_TOP):
    return top + "".join(f"## {name}\n\nx\n\n" for name in names)


def _swap(old, *new):
    """The one shape with one section replaced by others, or by none."""
    i = _ONE_SHAPE.index(old)
    return _ONE_SHAPE[:i] + new + _ONE_SHAPE[i + 1:]


def _moved(names, name, before):
    """names with one section moved to sit just before another."""
    rest = [n for n in names if n != name]
    i = rest.index(before)
    return tuple(rest[:i] + [name] + rest[i:])


_OWN_AFTER = "is a section of the project's own, after 'Status'"
_NOT_FIRST = "the banner (" + BANNER + ") is not the first thing in the file"

_ORDER_CASES = [
    ("the one shape passes", _page(), ()),
    ("'How to use it' may follow 'Get started'",
     _page(_swap("Get started", "Get started", "How to use it")), ()),
    ("a page with no section of its own passes", _page(_NO_OWN), ()),
    ("'How to use it' before 'Get started' fails", _page(_swap("Get started", "How to use it", "Get started")),
     ("'Get started' is out of order",)),
    ("'How to use it' after a section of its own fails",
     _page(_swap("Numbers", "Numbers", "How to use it")),
     ("'How to use it' comes after the project's own sections",)),
    ("two sections of its own share the one slot", _page(_swap("Numbers", "Numbers", "Proof")), ()),
    ("an HTML banner passes",
     _page(top=_TOP.replace(_BANNER_LINE, "<p align=\"center\"><img src=\"" + BANNER + "\"></p>")), ()),
    ("an absolute banner URL passes",
     _page(top=_TOP.replace("(" + BANNER, "(https://raw.githubusercontent.com/o/r/main/" + BANNER)), ()),
    ("HTML badges pass", _page(top=_TOP.replace(_BADGE_LINE, "<a href=\"https://x\"><img src=\"https://x/b.svg\"></a>")), ()),
    ("a heading inside a ``` block is not a section", _page() + "```\n## Zebra\n```\n", ()),
    ("a heading inside a ~~~ block is not a section", _page() + "~~~\n## Zebra\n~~~\n", ()),
    ("a ``` line does not close a ~~~ block", _page() + "~~~\n```\n## Zebra\n~~~\n", ()),
    ("a level-3 heading is not a section", _page() + "### Every type in the package\n\nx\n", ()),
    ("a '# ' line inside a code block is not a name", _page(top=_TOP + "```\n# a comment\n```\n\n"), ()),
    ("the generated type list under Documentation, with Licence below it, passes",
     _page().replace("## Licence", TYPE_INDEX_MARKER + "\n### Word\n\n`DocxEditor`\n<!-- END GENERATED -->\n\n## Licence"),
     ()),

    ("'Status' before 'Get started' fails", _page(_moved(_NO_OWN, "Status", "Get started")),
     ("'Get started' is out of order",)),
    ("'FAQ' before 'Alternatives' fails", _page(_moved(_ONE_SHAPE, "FAQ", "Alternatives")),
     ("'Alternatives' is out of order",)),
    ("'Licence' before 'Documentation' fails", _page(_moved(_ONE_SHAPE, "Licence", "Documentation")),
     ("'Documentation' is out of order",)),
    ("'Get started' after a section of its own fails", _page(_moved(_ONE_SHAPE, "Numbers", "Get started")),
     ("'Get started' comes after the project's own sections",)),
    ("a section of its own after 'Status' fails", _page(_moved(_ONE_SHAPE, "Numbers", "Limits")),
     ("'Numbers' " + _OWN_AFTER,)),
    ("a section of its own at the end fails", _page(_ONE_SHAPE + ("Credits",)),
     ("'Credits' " + _OWN_AFTER,)),
    ("a level-2 heading in the generated type list fails",
     _page().replace("## Licence", TYPE_INDEX_MARKER + "\n## Word\n\n`DocxEditor`\n<!-- END GENERATED -->\n\n## Licence"),
     ("'Word' " + _OWN_AFTER,)),
    ("'Install' is not 'Get started'", _page(_swap("Get started", "Install")),
     ("no '## Get started' section",)),
    ("'Learn more' is not 'Documentation'", _page(_swap("Documentation", "Learn more")),
     ("no '## Documentation' section", "'Learn more' " + _OWN_AFTER)),
    ("a missing Licence fails", _page(_swap("Licence")), ("no '## Licence' section",)),
    # A renamed 'What it is' is a section of the project's own, so 'Get started' now follows one.
    ("'What it does' is not 'What it is'", _page(_swap("What it is", "What it does")),
     ("no '## What it is' section", "'Get started' comes after the project's own sections",
      "the first section is 'What it does', not 'What it is'")),
    ("a section of its own before 'What it is' fails", _page(_moved(_ONE_SHAPE, "Numbers", "What it is")),
     ("'What it is' comes after the project's own sections",
      "'Get started' comes after the project's own sections",
      "the first section is 'Numbers', not 'What it is'")),
    ("an empty section fails", _page().replace("## Limits\n\nx\n\n", "## Limits\n\n"),
     ("'Limits' has nothing under it",)),

    ("the icon above the banner fails", _page(top="![icon](assets/icon.png)\n\n" + _TOP), (_NOT_FIRST,)),
    ("a first picture that is not the banner fails", _page(top=_TOP.replace(BANNER, "assets/icon.png")),
     (_NOT_FIRST,)),
    ("the name above the banner fails",
     _page(top="# T\n\n" + _BANNER_LINE + "\n\n" + _BADGE_LINE + "\n\n" + _ITALIC_LINE + "\n\n"), (_NOT_FIRST,)),
    ("a code block above the banner fails", _page(top="```\nx\n```\n\n" + _TOP), (_NOT_FIRST,)),
    ("no '# Name' line fails", _page(top=_TOP.replace("# T\n\n", "")),
     ("expected one '# Name' line above the first section, found 0",)),
    ("two '# Name' lines fail", _page(top=_TOP.replace("# T\n\n", "# T\n\n# U\n\n")),
     ("expected one '# Name' line above the first section, found 2",)),
    ("no line in italics fails", _page(top=_TOP.replace(_ITALIC_LINE, _ITALIC_LINE.strip("*"))),
     ("no one-line description in italics under the name",)),
    ("a line in bold is not a line in italics", _page(top=_TOP.replace(_ITALIC_LINE, "*" + _ITALIC_LINE + "*")),
     ("no one-line description in italics under the name",)),
    ("the line in italics above the name is not under it",
     _page(top=_BANNER_LINE + "\n\n" + _ITALIC_LINE + "\n\n# T\n\n" + _BADGE_LINE + "\n\n"),
     ("no one-line description in italics under the name",)),
    ("no badge line fails", _page(top=_TOP.replace(_BADGE_LINE + "\n\n", "")),
     ("no badge line under the name",)),
    ("badges above the name are not under it",
     _page(top=_BANNER_LINE + "\n\n" + _BADGE_LINE + "\n\n# T\n\n" + _ITALIC_LINE + "\n\n"),
     ("no badge line under the name",)),
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
    for name, text, expected in _ORDER_CASES:
        errors = check_order(text)
        exact = len(errors) == len(expected) and all(any(want in e for e in errors) for want in expected)
        if not exact:
            bad += 1
            print(f"::error::self-test case '{name}': expected {list(expected) or 'pass'}, "
                  f"got {errors or 'pass'}")
    # The real run must use every rule. One page that passes them all, then one that breaks
    # one rule of each kind: a 26-word sentence, no Limits, a section of its own at the end.
    errors = check_page(_page(), "root")
    if errors:
        bad += 1
        print(f"::error::self-test: the one shape should pass every rule, got {errors}")
    errors = check_page(_page(_swap("Limits") + ("Credits",)) + _LONG26 + "\n", "root")
    for want in ("sentence has 26 words", "no '## Limits' section", "'Credits' " + _OWN_AFTER):
        if not any(want in e for e in errors):
            bad += 1
            print(f"::error::self-test: the real run should report \"{want}\", got {errors}")
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
        errors = check_page(text, kind)
        for line in errors:
            print(f"::error file={path}::{line}")
        failed |= bool(errors)
        if not errors:
            print(f"{path}: front page within shape")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
