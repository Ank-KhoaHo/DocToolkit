# Presentations

Reading a PowerPoint file: slide count, text extraction, and placeholder replacement.

```bash
dotnet run --project samples/Presentations
```

Prints the slide count and the first slide's text, before and after a replacement.

## The non-obvious part

**`ExtractText` returns one entry per slide, in deck order** — not one blob of text.

**PowerPoint splits words across runs**, exactly as Word does. A single visible `{{who}}` is often
several `<a:t>` elements in the underlying XML, so a naive per-run `string.Replace` would miss it.
`ReplaceText` maps matches back onto the individual runs they overlap, which is what preserves
per-run formatting.

**There is no "create a PPTX from scratch" method**, so this sample reads a committed fixture —
the test project's `sample.pptx`, borrowed rather than duplicated.
