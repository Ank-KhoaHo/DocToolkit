# DOCX templating

Filling a Word template: **scalar placeholders** with `ReplaceText`, and **one table row per
record** with `FillRows`.

```bash
dotnet run --project samples/DocxTemplating
```

Prints the extracted text of the filled document, how many line items one template row became, and
whether any placeholder survived.

## The non-obvious part

**Call `FillRows` before `ReplaceText`, not the other way round.**

`FillRows` clones the template row once per record. Any document-level scalar you substituted
*first* gets cloned along with it, so `{{customer}}` ends up repeated inside every line item
instead of appearing once in the heading. Rows first, then scalars.

That ordering rule is why both operations are shown in one sample: separated, the rule would live
in a folder that does not contain the code it constrains.

Worth knowing: a placeholder is often several `<w:t>` runs in the underlying XML, because Word
splits text as you type. Both methods handle that — a naive per-run `string.Replace` would not.
