# Repeating table rows — design

Backlog item **A4**, from `2026-08-03-enhancement-backlog.md`.

## Why

`DocxEditor.ReplaceText` substitutes scalars: `{{customer}}` becomes `Contoso Ltd`. It has no way
to express *one row per record*, which is the most common real-world Word-template need — invoice
line items, timesheet entries, order lines. A caller who needs that today has to build the table
themselves with the raw OpenXML SDK, which is precisely the work this package exists to remove.

This is also the first change in a while that a *user* of the package would notice. Two complete
infrastructure tracks shipped before it without adding a capability.

## Scope

**In:** one level of repetition over table rows, driven by placeholders already in the template.

**Out, deliberately:** nested or hierarchical data, conditionals (`{{#if}}`), repetition of
non-table content (paragraphs, sections), and expressions or formatting directives inside
placeholders. Those are how template engines grow without bound. Scalars plus one level of
repetition covers the stated need; anything past it should be its own decision, not a slippery
slope entered by accident.

## Template syntax

A row repeats when any of its cells contains a placeholder prefixed with the collection name:

| Description      | Qty          | Total          |
|------------------|--------------|----------------|
| `{{item.Desc}}`  | `{{item.Qty}}` | `{{item.Total}}` |

The placeholders *are* the loop declaration. There are no separate open/close markers, which means
nothing has to be stripped from the finished document and there is no unbalanced-marker error class
to design for.

Chosen over explicit `{{#items}}`/`{{/items}}` markers, which would additionally support records
spanning several rows — genuinely more expressive, and rejected as unnecessary for the need at hand.
Chosen over caller-supplied table and row indices, which require no template syntax at all but
break silently when a table is inserted above the target.

**This introduces a delimiter convention the rest of the editor does not have.** `ReplaceText` never
parses anything — it matches whatever literal keys the caller supplies, so `{{ }}` is convention,
not code. `FillRows` must *find* placeholders by pattern, so it has to know the delimiters, and it
hard-codes `{{` and `}}`. Making them configurable is deferred until someone needs it.

## Public API

```csharp
public static byte[] FillRows(
    byte[] docx,
    string collection,
    IEnumerable<IReadOnlyDictionary<string, string>> rows);

public static Task FillRowsAsync(
    Stream source,
    string collection,
    IEnumerable<IReadOnlyDictionary<string, string>> rows,
    Stream destination,
    CancellationToken ct = default);
```

Purely additive: no existing signature changes, so no consumer breaks. Both overloads follow the
house shape — `Stream source` where the `byte[]` overload took bytes, then `Stream destination`,
then `CancellationToken ct = default` — and a single `FillRowsCore` holds the one real
implementation that both call, so they cannot drift apart.

Row dictionary keys are **bare field names** (`Desc`), not full placeholders. The collection name is
already an argument; repeating it in every key of every row would duplicate it many times over.
This does diverge from `ReplaceText`, whose keys are the full placeholder text including braces, and
that difference must be stated plainly in both methods' doc comments rather than left to be
discovered.

### Composition, not combination

Row expansion and scalar replacement stay separate methods:

```csharp
byte[] filled = DocxEditor.FillRows(docx, "item", lineItems);
filled = DocxEditor.ReplaceText(filled, new() { ["{{customer}}"] = "Contoso Ltd" });
```

This costs a second open/save cycle. It buys a public API that is additive rather than overlapping,
and it follows the composition the codebase already prefers — `HtmlToPdfConverter` composes the
other two converters rather than reimplementing conversion inside itself.

Order matters and is documented: expand rows first, then replace scalars. The reverse also works
for values that appear only outside the table, but expanding first is the rule that always holds.

## Semantics

- A row is a **template row** if any of its cells contains `{{collection.` — every other row is
  untouched.
- Each record produces one clone, in order, inserted where the template row was. The template row
  is then removed.
- `{{collection.Field}}` resolves to the value at key `Field`.
- A placeholder with **no matching key resolves to empty**, not left visible. A half-filled document
  showing `{{item.Missing}}` to an end user is worse than a blank cell.
- Placeholders for **other prefixes are left alone**, so a second `FillRows` call fills a second
  table in the same document.
- An **empty collection removes the template row** — and removes the whole table if that row was
  the only one (see *Traps*).

Two cases the simple description leaves open, made explicit because either could reasonably be read
the other way:

- **Several template rows for the same collection each expand independently.** Two rows in the same
  table both containing `{{item.`, given three records, produce three clones of the first followed
  by three clones of the second — not a repeating two-row block. Treating them as a block is what
  open/close markers are for, and those were rejected above; independent expansion is the reading
  that follows from "a row repeats if it contains the prefix", with no extra concept.
- **A template row inside a nested table expands in its own right.** It is not swept into the outer
  row's expansion (see *Traps*), and it is not skipped either — the rule is applied uniformly at
  whatever depth the row sits.

## Implementation

Clone the template row per record, then reuse `RunTextSplicer` on the clone.

For each record: deep-clone the row, build a per-record dictionary mapping `{{collection.Field}}` to
each value, and run the existing splicer over each paragraph's `w:t` nodes within the clone.

The formatting guarantee then comes for free. The clone carries every run property, cell border and
shading; the splicer already writes back only to the runs a match actually overlaps, so runs outside
a match keep their text and formatting and a hyperlink inside a cell survives untouched — the same
properties `ReplaceText` has today, for the same reason, via the same code.

**The alternative was a dedicated row-substitution engine**, giving more control over row-specific
cases. Rejected: it would duplicate the offset-to-run mapping, which is the single trickiest piece
of code in the repo and the one `CLAUDE.md` warns explicitly against reimplementing — "don't
simplify it into a merge-everything-onto-run-0 loop; that silently flattens formatting and guts
hyperlinks." Two copies of that logic would eventually disagree, and the disagreement would be
silent.

## Traps

Each is the analogue of a hazard this codebase has already been bitten by.

**Nested tables must not be discovered by `Descendants`.** `body.Descendants<TableRow>()` also
yields rows of tables nested inside cells. `CLAUDE.md` documents the same shape of bug for
`Descendants<Paragraph>()` reaching into `w:txbxContent`, where the result was "schema-valid, no
exception, silent data loss." Row discovery therefore walks the **direct child rows of each table**
and recurses into nested tables deliberately, so an inner row is expanded in its own right rather
than as part of its container.

**Removing the last row may invalidate the table.** If the template row is a table's only row and
the collection is empty, deleting it appears to leave a `w:tbl` with no `w:tr` — which Word is
expected to treat as corrupt: a file that saves without error and fails to open. The design's answer
is to remove the whole table in that case. **This must be verified empirically during
implementation, not assumed** — write the zero-row table, open it, and see. If Word accepts it, the
simpler behaviour of leaving an empty table is fine.

**A placeholder split across runs inside a cell** is already handled, but only if the splicer is fed
each paragraph's `w:t` nodes rather than the cell's. That is how the existing code does it and the
row path must not "simplify" it.

## Error handling

Matching the existing API's contract exactly.

| Condition | Result |
|---|---|
| `docx`, `collection` or `rows` is null | `ArgumentNullException` |
| `docx` is empty, or `collection` is blank | `ArgumentException` |
| The package could not be opened or edited | `DocumentConversionException` |
| `ct` was cancelled | `OperationCanceledException` |
| **No template row found for the collection** | `DocumentConversionException` |

That last row is a deliberate choice. Passing ten line items to a document containing no `{{item.`
row would otherwise do nothing at all, successfully — and a silent no-op is the failure mode this
repo consistently refuses. A mismatch between the call and the template is a bug in one of them, so
it says so.

## Testing

Beyond the round-trip, the tests that earn their place:

- **Formatting survives.** A template row with a bold run, cell shading and custom borders produces
  clones carrying all three. Without this the feature has no advantage over hand-building rows.
- **A placeholder split across runs inside a cell** still substitutes — the `RunTextSplicer`
  guarantee, re-proved at row level.
- **A hyperlink in a cell** survives cloning and substitution.
- **Nested table.** A table inside a cell whose rows also contain `{{item.` must expand in its own
  right, and must **not** be swept into the outer row's expansion. Without this test, the
  `Descendants` trap returns as schema-valid data loss.
- **Two template rows for one collection** in the same table each expand independently, giving
  clones of the first followed by clones of the second.
- **Empty collection** removes the template row, and removes the whole table when that row was its
  only one.
- **Other prefixes untouched.** `{{payment.Total}}` survives `FillRows(…, "item", …)`, which is what
  makes a second table possible.
- **Missing key** yields an empty cell, not a visible `{{item.Missing}}`.
- **No template row** throws rather than silently succeeding.

### Two repo-specific obligations

1. **`FillRowsAsync` must be added to the name lists at the top of `StreamOverloadTests`.**
   `CLAUDE.md`: *"an overload missing from those lists is the only way to escape the whole suite."*
   Omitting it means the stream-shape guarantees — caller-owned streams never disposed or sought,
   forward-only sources honoured, genuinely async I/O — go unproven for this method.
2. **`FillRows` must be added to `AirGapGuardTests`.** That suite asserts zero connections across
   the *whole public API*, and `README.md` quantifies it as 35 guard tests. This method opens no
   sockets; the guard's value lies in being exhaustive rather than in any individual case.

## Success criteria

- A template row containing `{{item.X}}` placeholders produces one row per record, in order, each
  carrying the template row's formatting.
- A hyperlink and a split-run placeholder inside a template cell both survive.
- A nested table's rows are not swept into the outer expansion.
- An empty collection leaves a valid document that Word opens without repair.
- `FillRows` and `FillRowsAsync` produce identical output for identical input.
- Build stays at 0 warnings under `-warnaserror`; the whole suite passes on both target frameworks.
