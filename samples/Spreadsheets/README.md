# Spreadsheets

Creating an XLSX, reading and writing single cells, and reading a whole sheet you know nothing
about in advance.

```bash
dotnet run --project samples/Spreadsheets
```

Prints a cell before and after an edit, then the sheet names and the full grid.

## The non-obvious part

**`ReadSheet` anchors its result at A1, not at the first cell containing data.** If a sheet's data
starts at C3, that value is at `rows[2][2]` and everything before it is `""`. That keeps
`rows[r][c]` meaning what it looks like it means. Rows are padded to a uniform width, and entirely
blank rows inside the range are kept rather than dropped — dropping them would shift every later
index.

**Cells come back as strings**, produced the same way `ReadCell` produces them, so the two can
never disagree. A formula cell yields its **cached** value: nothing in this package evaluates
formulas.

**`ReadSheet` refuses a sheet spanning more than 2,000,000 cells.** The result is materialised
whole, so its cost tracks the *rectangle*, not how much of it holds data — one stray value in a far
corner of a sheet describes an enormous grid from a file only a few KB on disk. It throws
`DocumentConversionException` naming the actual extent rather than exhausting memory.

**`SheetNames` includes hidden sheets**, in tab order. Hiding a sheet is a presentation choice, not
a privacy boundary.
