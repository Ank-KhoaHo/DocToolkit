# More badges (downloads, coverage) — design

## Why

The README currently has CI and NuGet-version badges. Two more give adopters a faster read on the
package's health: download counts (adoption signal) and test coverage (quality signal). Coverage
specifically needs collection wired into CI first — it isn't just a markdown addition like the
downloads badge is.

## Downloads badges

No CI change. Shields.io serves these directly from the nuget.org API:

```
[![NuGet Downloads](https://img.shields.io/nuget/dt/Ank.DocToolkit.svg)](https://www.nuget.org/packages/Ank.DocToolkit/)
[![NuGet Downloads](https://img.shields.io/nuget/dt/Ank.DocToolkit.Extensions.DependencyInjection.svg)](https://www.nuget.org/packages/Ank.DocToolkit.Extensions.DependencyInjection/)
```

Placement mirrors the existing version badges: both in the top-level `README.md`, and the
extensions package's own badge in `src/DocToolkit.Extensions.DependencyInjection/README.md`.

## Coverage collection

Coverlet is already referenced (`coverlet.collector`) in both test projects — no new package
reference needed. `.github/workflows/ci.yml`'s existing `Test` step
(`build-test` job, both `matrix.os` legs) gets `--collect:"XPlat Code Coverage"` added:

```yaml
- name: Test
  run: >
    dotnet test ${{ env.SOLUTION }} -c Release --no-build
    --collect:"XPlat Code Coverage"
    --logger "trx;LogFileName=${{ matrix.os }}.trx"
    --results-directory ${{ github.workspace }}/TestResults
```

This produces one `coverage.cobertura.xml` per test project per target framework — up to four
files (two test projects × `net8.0`/`net10.0`) under `TestResults/`, alongside the trx logs
already collected there.

## Codecov upload

A new step, gated to the `ubuntu-latest` leg only — Windows exists in this matrix to catch
Linux-only assumptions, not to double-report identical coverage numbers:

```yaml
- name: Upload coverage to Codecov
  if: matrix.os == 'ubuntu-latest'
  uses: codecov/codecov-action@v5
  with:
    token: ${{ secrets.CODECOV_TOKEN }}
    fail_ci_if_error: false
```

**Token requirement is genuinely conditional, not something to assume either way** (verified
against Codecov's own docs, not memory — an earlier draft of this design incorrectly assumed
public repos are always tokenless): tokenless upload only applies to unprotected/fork-PR branches
and to organizations that have explicitly opted out of requiring a token; **existing** Codecov
orgs default to requiring one, **new** orgs default to not requiring one, and this repo's push
events come from `main` directly, not a fork PR. Since which default applies here isn't knowable
until the org exists on codecov.io, the step passes `token: ${{ secrets.CODECOV_TOKEN }}`
proactively — harmless (an empty/unset secret behaves as if the input were omitted) — with the
secret itself added only if the first real run's upload turns out to need it.

`fail_ci_if_error: false` because an upload hiccup shouldn't fail the build the badge itself is
supposed to be reporting on — the badge going stale is a soft signal, not a release-blocking one
(unlike the CHANGELOG guard, which blocks a real irreversible action).

Codecov auto-detects the cobertura files in the workspace; no need to enumerate them or merge
across TFMs ourselves — Codecov aggregates multiple reports for one commit natively (it's built
for exactly this, e.g. matrix builds).

**One manual step, once:** sign in to codecov.io with GitHub and enable the `DocToolkit` repo, so
there's somewhere to receive the first upload — same shape as the nuget.org Trusted Publishing
policy setup earlier in this project. **If the first post-merge CI run's upload step fails or the
badge doesn't populate**, the fix is almost certainly the token: go to the repo's settings on
codecov.io, copy the repository upload token, and add it as a `CODECOV_TOKEN` repository secret
in GitHub — the same diagnose-then-fix shape as the nuget release-extensions.yml OIDC failure
earlier in this project, not a sign the workflow step itself is broken.

## Badge markdown

Added to `README.md` next to the existing CI badge:

```
[![codecov](https://codecov.io/gh/Ank-KhoaHo/DocToolkit/branch/main/graph/badge.svg)](https://codecov.io/gh/Ank-KhoaHo/DocToolkit)
```

## Explicitly out of scope

No enforced minimum-coverage threshold (Codecov supports failing a PR below a target percentage).
This design only adds visibility, not a gate — a coverage gate is a separate decision worth its
own brainstorm if wanted later.

## Testing / validation

No unit tests apply. Validation is: push this to a branch, open a PR (or push to `main` directly,
since `ci.yml` also triggers on `main`), watch the `Test` step actually emit
`coverage.cobertura.xml` files, watch the Codecov upload step succeed, and confirm the badge
renders a real percentage (not "unknown") once codecov.io has processed the report.
