<!--
Thank you for contributing. The checklist below is short on purpose: everything CI can check for
itself is left to CI. What is listed is what CI cannot catch, and each item is here because it has
gone wrong at least once.

Delete anything that does not apply.
-->

## What this changes, and why

<!-- The behaviour, not the diff. If it fixes something, what was the symptom? -->

## Checklist

- [ ] **Commit messages follow Conventional Commits** (`type(scope): description`). The type decides
      the version bump and the changelog section, so `fix:` for a test-only change publishes noise —
      `test:`, `ci:`, `docs:` and `chore:` are hidden and non-bumping.
- [ ] **No `Co-Authored-By` trailer.**
- [ ] **The PR title carries no Conventional Commit prefix.** Merge commits copy the title into the
      body, and release-please parses it — a prefixed title duplicates every changelog entry.
- [ ] If the **public API changed**, the approved-API file is updated and the diff in this PR is the
      change you meant. That diff is the only review this package's surface gets.
- [ ] If a **consumer-facing capability or limitation** changed, it is in `src/DocToolkit/README.md`
      — *not only* the root `README.md`, which is **not** what nuget.org renders.
- [ ] New tests were **watched to fail** before the code that makes them pass, and anything
      security- or silence-related was mutation-verified.

## The four constraints

Any change that breaks one of these makes the package pointless, and CI will reject it:

- [ ] No dependency that is not permissively licensed
- [ ] No native binaries; `dotnet restore` is still the whole install
- [ ] Still runs on Linux, Windows, macOS and arm64
- [ ] No new code path opens a socket by default
