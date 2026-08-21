"""Regenerate BACKLOG.html from the specs.

THE SPECS ARE THE SOURCE OF TRUTH; the index is a view over them. Each spec declares its own id,
status, section and headline in its front matter, and this reads those rather than any separate
data file - so there is exactly one place a ticket lives and the index cannot disagree with it.

It used to read window.__ROWS__ out of the previous BACKLOG.html. That stopped working the moment
the rows moved into the markup, which is the correct outcome: the old page WAS the database, and
a page that is its own database is the thing this refactor set out to end.

Usage:
    python scripts/gen-backlog-index.py
"""
import glob
import html as H
import io
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SPECS = os.path.join(ROOT, "docs", "superpowers", "specs")

ORDER = ["OPEN", "ACCEPTED", "DONE", "SUSPENDED", "DROPPED"]
MEANS = {
    "OPEN": "filed by the assistant; the maintainer has not ruled on it",
    "ACCEPTED": "reviewed and validated by the maintainer - the only status that means work to do",
    "DONE": "implemented and merged to main",
    "SUSPENDED": "closed by the assistant without work - declined, or measured and found sound",
    "DROPPED": "declined by the maintainer",
}

HEADING = re.compile(r"^#\s+([A-Z]+[0-9][A-Za-z0-9-]*)\s*$", re.M)
STATUS = re.compile(r"^\*\*Status:\*\*\s*(.+?)\s*·\s*(\S+)\s*$", re.M)
FILED = re.compile(r"^\*\*Filed:\*\*\s*(\S+)\s*$", re.M)
HEADLINE = re.compile(r"^\*\*Headline:\*\*\s*(.+?)\s*$", re.M)


def load_tickets():
    """Read every ticket spec. A spec that does not declare an id is not a ticket."""
    tickets = []
    for path in sorted(glob.glob(os.path.join(SPECS, "*.md"))):
        name = os.path.basename(path)
        if not re.match(r"^[A-Z]+[0-9][A-Za-z0-9-]*-", name):
            continue                                   # a dated design doc, not a ticket
        text = io.open(path, encoding="utf-8").read()
        m = HEADING.search(text)
        if not m:
            print("  skipped, no '# <ID>' heading: " + name)
            continue
        s = STATUS.search(text)
        tickets.append({
            "id": m.group(1),
            "section": s.group(1) if s else "",
            "status": s.group(2) if s else "",
            "era": (FILED.search(text).group(1) if FILED.search(text) else ""),
            "headline": (HEADLINE.search(text).group(1) if HEADLINE.search(text) else ""),
            "file": name,
        })
    return tickets


CSS = """
:root{--bg:#fbfaf8;--fg:#1c1a17;--mut:#6b665e;--line:#e2ded6;--card:#fff;
 --s-OPEN:#1f5fa8;--s-ACCEPTED:#4a97d8;--s-DONE:#12923f;--s-SUSPENDED:#9c9992;--s-DROPPED:#d13c2b}
@media(prefers-color-scheme:dark){:root:not([data-theme=light]){
 --bg:#161513;--fg:#ece8e1;--mut:#9c968c;--line:#302d28;--card:#1e1c19;
 --s-OPEN:#2c5dbd;--s-ACCEPTED:#359ccb;--s-DONE:#35a84f;--s-SUSPENDED:#96938c;--s-DROPPED:#e8604e}}
*{box-sizing:border-box}
body{margin:0;padding:2rem 1.25rem 4rem;background:var(--bg);color:var(--fg);
 font:15px/1.55 ui-sans-serif,-apple-system,"Segoe UI",Roboto,sans-serif}
.wrap{max-width:1180px;margin:0 auto}
h1{font-size:1.5rem;margin:0 0 .3rem}
.lede{color:var(--mut);margin:0 0 1.5rem;max-width:78ch}
.lede code{background:var(--card);border:1px solid var(--line);border-radius:4px;padding:.05em .35em}
.tools{display:flex;gap:.6rem;flex-wrap:wrap;margin:0 0 1.4rem}
input[type=search]{flex:1 1 16rem;padding:.5rem .7rem;border:1px solid var(--line);
 border-radius:7px;background:var(--card);color:var(--fg);font:inherit}
details.group{margin:0 0 1.1rem;border:1px solid var(--line);border-radius:10px;
 background:var(--card);overflow:hidden}
summary{cursor:pointer;padding:.65rem .9rem;font-weight:600;list-style:none;
 display:flex;align-items:baseline;gap:.55rem;flex-wrap:wrap}
summary::-webkit-details-marker{display:none}
summary::before{content:"\\25B8";display:inline-block;transition:transform .15s;color:var(--mut)}
details[open]>summary::before{transform:rotate(90deg)}
.count{color:var(--mut);font-weight:500}
.means{color:var(--mut);font-weight:400;font-size:.83rem}
table{width:100%;border-collapse:collapse;font-size:.9rem}
thead th{text-align:left;padding:.4rem .9rem;color:var(--mut);font-weight:600;
 border-top:1px solid var(--line);border-bottom:1px solid var(--line);font-size:.78rem;
 text-transform:uppercase;letter-spacing:.04em}
td{padding:.5rem .9rem;border-bottom:1px solid var(--line);vertical-align:top}
tr:last-child td{border-bottom:none}
td.id{white-space:nowrap;font-variant-numeric:tabular-nums}
td.id a{font-weight:700;text-decoration:none;color:inherit;
 border-bottom:2px solid currentColor;padding-bottom:1px}
td.sum{max-width:62ch;overflow-wrap:anywhere}
td.sec,td.era{color:var(--mut);white-space:nowrap;font-size:.85rem}
#g-open td.id a{color:var(--s-OPEN)}#g-accepted td.id a{color:var(--s-ACCEPTED)}
#g-done td.id a{color:var(--s-DONE)}#g-suspended td.id a{color:var(--s-SUSPENDED)}
#g-dropped td.id a{color:var(--s-DROPPED)}
tr[hidden]{display:none}
@media(max-width:720px){td.sec,td.era,thead th:nth-child(3),thead th:nth-child(4){display:none}}
"""


def esc(s):
    return H.escape(s or "", quote=True)


def main():
    tickets = load_tickets()
    if not tickets:
        print("REFUSING: no ticket specs found - an index generated from nothing would look "
              "exactly like a backlog with nothing in it")
        return 1

    unknown = sorted({t["status"] for t in tickets} - set(ORDER))
    if unknown:
        print("REFUSING: spec(s) declare a status outside the known five: " + ", ".join(unknown))
        return 1

    parts = []
    for status in ORDER:
        group = [t for t in tickets if t["status"] == status]
        if not group:
            continue
        parts.append('<details class="group" id="g-%s" open>' % status.lower())
        parts.append('<summary>%s <span class="count">(%d)</span> '
                     '<span class="means">%s</span></summary>'
                     % (status, len(group), esc(MEANS[status])))
        parts.append('<table><thead><tr><th>ID</th><th>Headline</th><th>Section</th>'
                     '<th>Filed</th></tr></thead><tbody>')
        for t in sorted(group, key=lambda x: (x["section"], x["id"])):
            head = t["headline"]
            short = head if len(head) <= 190 else head[:187] + "..."
            parts.append(
                '<tr data-id="%s" data-status="%s" data-section="%s">'
                '<td class="id"><a href="docs/superpowers/specs/%s">%s</a></td>'
                '<td class="sum" title="%s">%s</td>'
                '<td class="sec">%s</td><td class="era">%s</td></tr>'
                % (esc(t["id"]), esc(status), esc(t["section"]), esc(t["file"]), esc(t["id"]),
                   esc(head), esc(short), esc(t["section"]), esc(t["era"])))
        parts.append("</tbody></table></details>")

    total = len(tickets)
    doc = """<!doctype html>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>DocToolkit backlog</title>
<style>%s</style>
<div class="wrap">
<h1>DocToolkit backlog</h1>
<p class="lede">%d tickets. <strong>Every row here is an index entry, not the ticket.</strong>
The ticket is its file under <code>docs/superpowers/specs/</code> - click the ID, and that file is
what this page is generated FROM. This page and those files are <strong>gitignored</strong>: this
repository is public and the backlog is not. See <code>PUBLIC.md</code>.</p>
<div class="tools">
  <input type="search" id="q" placeholder="Filter by id, headline or section - rows render without JavaScript"
         aria-label="Filter tickets">
</div>
%s
</div>
<script>
// ADDITIVE ONLY. Every row above is already in the markup: if this block throws, is stripped, or
// never runs, the page still renders all %d tickets and `grep BACKLOG.html` still finds every id.
// A backlog page that renders empty looks exactly like a backlog with nothing in it.
try {
  var q = document.getElementById('q');
  var rows = Array.prototype.slice.call(document.querySelectorAll('tbody tr'));
  q.addEventListener('input', function () {
    var t = q.value.trim().toLowerCase();
    rows.forEach(function (r) {
      r.hidden = t !== '' && r.textContent.toLowerCase().indexOf(t) === -1;
    });
    document.querySelectorAll('details.group').forEach(function (g) {
      var vis = g.querySelectorAll('tbody tr:not([hidden])').length;
      g.hidden = t !== '' && vis === 0;
    });
  });
} catch (e) { /* the page is fully usable without this */ }
</script>
""" % (CSS, total, "\n".join(parts), total)

    io.open(os.path.join(ROOT, "BACKLOG.html"), "w", encoding="utf-8", newline="\n").write(doc)
    print("BACKLOG.html rebuilt from %d specs" % total)
    for status in ORDER:
        n = sum(1 for t in tickets if t["status"] == status)
        if n:
            print("   %-10s %d" % (status, n))
    return 0


sys.exit(main())
