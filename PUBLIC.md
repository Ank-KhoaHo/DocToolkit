# This repository is public

Which means the working material a maintainer keeps — design notes, decision records, a backlog —
is **not published here**. If you went looking for it, it is deliberately absent rather than
missing.

Nothing withheld is secret. It is working material for whoever develops the package, and a
consumer arriving from nuget.org has no use for it.

## What is published is what you need

- **[README.md](README.md)** — what the package does, the constraints it holds, measured
  real-world conversion rates
- **[package README](src/DocToolkit/README.md)** — every capability, every overload, the caveats,
  and the migration notes per release. This is what nuget.org renders
- **[Guides](https://ank-khoaho.github.io/DocToolkit/guides/getting-started.html)** and the
  **[API reference](https://ank-khoaho.github.io/DocToolkit/)**
- **[CONTRIBUTING.md](CONTRIBUTING.md)** — how to build it, and what will get a pull request
  rejected
- **[ROADMAP.md](ROADMAP.md)** — direction, and the things deliberately not coming
- **[SECURITY.md](SECURITY.md)** — disclosure, scope, and the documented limits that are not
  findings

The standing rule is that anything in the unpublished material which turns out to matter to a
consumer **gets repeated in a README or turned into a CI check** — and the check is the better
answer, because prose goes stale and a check does not.

## If something you need is not here

**Open an issue.** That is not a formality: the gap in the guides that prompted issue #321 was
found exactly that way, and the fix shipped the same day.
