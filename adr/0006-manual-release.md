# 6. Releasing is a manual decision

**Status:** accepted

## Context

Release automation briefly auto-merged the Release PR, so every merge to `main` published.

It worked exactly as designed: **eleven versions in one day**, most carrying no library change at
all, because a docs fix and a CI tweak each consumed a version number.

A nuget.org version can be unlisted but **never deleted or reused**. That is not recoverable, merely
survivable.

## Decision

release-please computes the version and maintains a Release PR with the changelog already written.
Nothing publishes until a human merges it.

## Consequences

**The Release PR is open most of the time. That is the normal state, not a stuck pipeline** — it
accumulates commits until someone decides there is enough to release. Batching several merges into
one release is what a version number is for.

A related trap: only a commit type that computes a bump (`feat`, `fix`, `perf`, `build`, `revert`)
opens or updates the Release PR. A merge containing only `ci:`, `docs:`, `chore:` or `test:`
proposes nothing, and **silence after such a merge is healthy rather than broken**.

There is one asymmetry worth knowing: an auto-merged dependency update is merged with `GITHUB_TOKEN`,
and a push by that token does not trigger other workflows — so it does not immediately update the
Release PR. Nothing is lost, because release-please computes from commit history rather than from
events, and the next run sweeps it in. It is a delay, not a gap.

## What would change this

Very little. The failure this prevents is unrecoverable and the cost is one click. If publishing
ever needs to be automatic, the guard that must survive is the one refusing to publish a version
whose changelog entry is empty.
