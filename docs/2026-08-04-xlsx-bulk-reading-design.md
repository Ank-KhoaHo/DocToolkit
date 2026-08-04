# XLSX bulk reading — design

Backlog item **A3**, first slice. From `2026-08-03-enhancement-backlog.md`.

## Why

`WorkbookEditor` can create a single-sheet workbook, read one cell and write one cell. That is
enough to *produce* a spreadsheet and nothing like enough to *consume* one.

A caller handed an arbitrary `.xlsx` today cannot read it. `ReadCell` needs a sheet name and a cell
reference they already know, and there is no way to discover either — no sheet listing, no used
range, no bulk read. The only workable pattern is guessing cell references until one returns
something, which is not a pattern.

## Scope

A3 as written bundles four things with different shapes and different risks: bulk reading, bulk
writing, formulas, and CSV import/export. Designing them together would produce four specs wearing
one hat, and the formula and CSV questions — both arguably out of scope for a document toolkit —
would crowd out the part that carries the value.

**This spec is the reading slice only:** discover the sheets, read one whole sheet.

Bulk writing (append rows, multi-sheet create) is the natural next slice and gets its own spec.
Formulas and CSV are separate decisions, and the honest answer to both may be "no".

## Public API

```csharp
public static IReadOnlyList<string> SheetNames(byte[] xlsx);

public static Task<IReadOnlyList<string>> SheetNamesAsync(
    Stream source, CancellationToken ct = default);

public static IReadOnlyList<IReadOnlyList<string>> ReadSheet(byte[] xlsx, string sheetName);

public static Task<IReadOnlyList<IReadOnlyList<string>>> ReadSheetAsync(
    Stream source, string sheetName, CancellationToken ct = default);
```

Purely additive. Both pairs follow the house shape — a `byte[]` form and a `Stream` form, the async
one taking `Stream source` then `CancellationToken ct = default`, each pair sharing one `*Core` so
they cannot drift.

`ReadCell` already covers the single-cell case and is unchanged.

## Semantics

**Cells come back as strings**, produced by `.GetString()` — exactly what `ReadCell` uses. One
mental model across the whole XLSX surface, and no chance of `ReadCell` and `ReadSheet` disagreeing
about what a given cell says. A caller wanting a number parses it.

The alternative, `object?` with typed values, was considered and rejected: it preserves more but
diverges from `ReadCell`'s contract and pushes a type check into every call site.

**`SheetNames` returns names in workbook (tab) order, including hidden sheets.** Hiding a sheet is
a presentation choice, not a privacy boundary, and a caller who cannot see a hidden sheet in the
listing has no way to discover it exists.

**`ReadSheet` returns the used range anchored at A1.** If data starts at C3, `rows[2][2]` is C3 and
everything before it is `""`. This keeps `rows[r][c]` positionally meaningful — the alternative,
returning the used block from its own origin, forces every caller to track an offset separately or
silently misread positions.

- Every row is padded to the last used column, so all rows have the same length and `rows[r][c]`
  never throws on a short row.
- **Blank cells are `""`**, never null.
- **Entirely blank rows inside the range are kept**, as rows of `""`. Dropping them would shift
  every subsequent index and break the positional guarantee above.
- **An empty sheet returns an empty list**, not a single empty row.
- **A formula cell yields its cached value**, not the formula text. Nothing in this library
  evaluates formulas and that stays true.
- **A missing sheet throws `DocumentConversionException`**, matching `ReadCell` via the existing
  `Sheet()` helper.

## Implementation

ClosedXML's `Worksheets` collection and used-range accessors, reusing the `Open()` and `Sheet()`
helpers already in the file.

**Two alternatives were considered and rejected.** Reading the sheet XML directly with a SAX-style
reader would be faster and lighter on very large files, but means reimplementing shared-string
resolution and cell-type handling — the same duplication `CLAUDE.md` warns against for
`RunTextSplicer`. Lazy streaming via `IAsyncEnumerable` would genuinely help with large sheets and
speaks to backlog **A14**, but it is a different API shape from every other method here, and
adopting it for one operation would leave the surface inconsistent. Worth revisiting as its own
decision if A14 is ever addressed.

## Traps

**A sheet with no used range must be handled explicitly.** ClosedXML's used-range accessors return
null for an empty sheet, and code that assumes otherwise throws a `NullReferenceException` that gets
wrapped into a `DocumentConversionException` whose message points nowhere useful. The emptiness check
comes first and is unconditional — correct whether or not the null case is reachable on a given
input.

**Anchoring at A1 means not iterating `RangeUsed()`'s cells.** That range begins at the used origin,
so iterating it directly would put C3's value at `rows[0][0]` — precisely the behaviour this design
rules out. The loop runs from row 1 and column 1 to `LastRowUsed()` / `LastColumnUsed()`.

**ClosedXML counts formatted-but-empty cells as used.** A sheet where someone bolded column Z
reports a used range reaching column Z, so `ReadSheet` pads every row that far with `""`. That is
ClosedXML's definition of "used", not a defect, and matching it is better than inventing a second
definition — but it surprises people, so it is documented on the method.

**A1 anchoring has a memory cost.** A stray value at ZZ100000 produces a rectangular result of that
size, almost entirely empty strings. That is the price of positional indices being meaningful, and
it is the same class of concern as backlog **A14**. Stated rather than hidden.

## Error handling

| Condition | Applies to | Result |
|---|---|---|
| `xlsx` is null | both | `ArgumentNullException` |
| `xlsx` is empty | both | `ArgumentException` |
| `source` is not readable, or held no bytes | both async | `ArgumentException` |
| The workbook could not be opened | both | `DocumentConversionException` |
| `sheetName` is null | `ReadSheet` only | `ArgumentNullException` |
| `sheetName` is blank | `ReadSheet` only | `ArgumentException` |
| The sheet does not exist | `ReadSheet` only | `DocumentConversionException` |
| `ct` was cancelled | both async | `OperationCanceledException` |

`SheetNames` takes no sheet name, so the three sheet-name rows do not apply to it. Everything else
is `ReadCell`'s existing contract, reached through the same helpers.

## Testing

- **Data not starting at A1** returns rows anchored at A1. This is the behaviour most likely to be
  implemented the other way round by accident, and the result looks plausible either way.
- **Rows are rectangular**, including when the last row is shorter than the first.
- **Blank cells are `""`**, and an **empty sheet returns an empty list** rather than one empty row.
- **A formula cell yields its cached value** — pins the "nothing evaluates formulas" claim to a test
  rather than a sentence.
- **Hidden sheets appear in `SheetNames`**, in tab order.
- **A missing sheet throws**, matching `ReadCell`.
- **Round trip:** `Create` then `ReadSheet` returns what went in.
- **`byte[]` and `Stream` forms agree** for identical input.

### Repo-specific obligations

1. **`SheetNamesAsync` and `ReadSheetAsync` must be added to `StreamOverloadTests`' name lists** —
   both are source-readers with no destination, like `ReadCellAsync`. `CLAUDE.md`: an overload
   missing from those lists is the only way to escape the whole suite. Verify the result count rises
   rather than assuming the registration took.
2. **Both must be added to `AirGapGuardTests`**, which asserts zero connections across the whole
   public API.
3. **The B1 approved API files must be updated in the same PR.** Four new public methods means the
   approval test fails until the surface is re-approved — and that diff is the reviewable record of
   what was added. That is the guard working, not friction.

**Sequencing note:** the approval guard lands in its own pull request. This work must branch from a
`main` that already has it, or there will be no approved file to update.

## Success criteria

- A caller holding an arbitrary `.xlsx` can list its sheets and read one, without knowing anything
  about it in advance.
- A sheet whose data starts at C3 reads back with that value at `rows[2][2]`.
- `Create` followed by `ReadSheet` round-trips.
- `ReadCell` and `ReadSheet` never disagree about what a given cell contains.
- Build stays at 0 warnings under `-warnaserror --no-incremental`; the whole suite passes on both
  target frameworks.
