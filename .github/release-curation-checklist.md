<!-- release-curation-checklist -->
## Before merging this

Merging tags the version and publishes both packages to nuget.org. **A published version can be
unlisted but never edited**, so everything below is cheaper now than at any point afterwards.

- [ ] **Read the changelog entry in the diff.** An entry that is empty, or that says nothing a
      consumer would act on, is a signal to let more changes accumulate rather than release now.

- [ ] **Does anything here CHANGE existing behaviour?** release-please files every `feat:` under
      **Added**, which reads as a new capability. A changed default, a different output, or a new
      exception from a call that used to succeed is none of those — and a consumer upgrading will
      not learn about it from an *Added* line. Move it to **Changed** by hand, with the one-line
      migration.

      Not hypothetical: **0.13.0** changed generated documents from whatever the reader's template
      chose to A4, and the entry said only "defaulting to A4", filed under *Added*.

- [ ] **Does the nuget.org `<Description>` still describe what ships?** Check it against the
      generated [`docfx/guides/capabilities.md`](../docfx/guides/capabilities.md). CI proves the
      field is non-empty; **nothing can prove it is true**. It described a smaller library than
      shipped for four releases.

- [ ] **Once you arm auto-merge, this PR is finished.** Pushing to it afterwards is a race that is
      lost silently — `git push` reports success either way, and the commit is orphaned on a dead
      branch. Anything further goes on a new branch.

---

**Editing `CHANGELOG.md` on this branch is all that is needed.** The GitHub Release body is derived
from that entry automatically after merge, so the two cannot drift apart — they did on **0.27.2**,
where the curated changelog and the published release notes disagreed until somebody noticed by hand.
