# DOCX templating

Filling a Word template: **scalar placeholders** with `ReplaceText`, and **one table row per
record** with `FillRows`.

```bash
dotnet run --project samples/DocxTemplating
```

Prints the extracted text of the filled document, how many line items one template row became, and
whether any placeholder survived.

## The non-obvious part

**Call `FillRows` before `ReplaceText` — the rule is narrower than "always", but it's the safe
default.**

`FillRows` clones only the *template row* and substitutes into each clone alone (see `ExpandRow`
in `DocxEditor.cs`). If a scalar placeholder lives **inside** that row, substituting it first means
every clone inherits the already-filled text — so rows-first is the order that's safe in general.
In this sample `{{customer}}` sits in the `<h1>` above the table, outside the row `FillRows` ever
touches, so the two operations here happen to produce byte-identical output in either order — the
rule just doesn't bite this particular template.

Both operations are still shown in one sample because they're both templating and the ordering
rule is worth knowing before it bites you on a template that *does* put a scalar inside a
repeating row.

Worth knowing: a placeholder is often several `<w:t>` runs in the underlying XML, because Word
splits text as you type. Both methods handle that — a naive per-run `string.Replace` would not.
