# Architecture decision records

Why this package is shaped the way it is.

**These exist because the reasoning was previously unreachable.** It lived in `CLAUDE.md` and
`docs/`, both of which are gitignored — so a contributor cloning this repository, or anyone
evaluating the package, saw the decisions and none of the arguments. Several of them look arbitrary
without the evidence, and one of them has already been re-litigated once.

Each record states the decision, what it costs, and **what would change it**. That last part is the
point: a decision with no stated conditions for revisiting is a rule, and rules outlive their
reasons.

| # | Decision |
|---|---|
| [0001](0001-four-constraints.md) | The four constraints, and why they are enforced by CI rather than intent |
| [0002](0002-no-native-binaries.md) | Reject any dependency that ships native binaries |
| [0003](0003-pin-sixlabors-fonts.md) | Pin `SixLabors.Fonts` to `[1.0.0]` |
| [0004](0004-html-to-pdf-via-docx.md) | HTML → PDF pivots through DOCX |
| [0005](0005-below-one-point-oh.md) | Stay below 1.0.0 permanently |
| [0006](0006-manual-release.md) | Releasing is a manual decision |
