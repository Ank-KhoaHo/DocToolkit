# Spreadsheets

Creating an XLSX, reading and writing single cells, reading a whole sheet you know nothing about
in advance, and adding a pivot table or a chart to one.

```bash
dotnet run --project samples/Spreadsheets
```

Prints a cell before and after an edit, then the sheet names and the full grid, then adds a pivot
table and a chart over the same sales data.

## The non-obvious part

**`ReadSheet` anchors its result at A1, not at the first cell containing data.** If a sheet's data
starts at C3, that value is at `rows[2][2]` and everything before it is `""`. That keeps
`rows[r][c]` meaning what it looks like it means. Rows are padded to a uniform width, and entirely
blank rows inside the range are kept rather than dropped — dropping them would shift every later
index.

**Cells come back as strings**, produced the same way `ReadCell` produces them, so the two can
never disagree. **A formula cell yields its *computed* value, not a cached one** — the underlying
engine evaluates it, because a file this library writes carries no cached result. A reader that
only reads cached values sees an empty cell until Excel has opened and saved the file; this one
does not have that problem.

**A pivot table is the opposite case.** `AddPivotTable` writes the pivot's *definition*, and
nothing that writes a workbook — this method included — computes the aggregation. Reading the
pivot's own cells back with `ReadCell`/`ReadSheet` immediately after creating it returns empty
strings; open the result in Excel to see it populated.

**`ReadSheet` refuses a sheet spanning more than 2,000,000 cells.** The result is materialised
whole, so its cost tracks the *rectangle*, not how much of it holds data — one stray value in a far
corner of a sheet describes an enormous grid from a file only a few KB on disk. It throws
`DocumentConversionException` naming the actual extent rather than exhausting memory.

**`SheetNames` includes hidden sheets**, in tab order. Hiding a sheet is a presentation choice, not
a privacy boundary.
