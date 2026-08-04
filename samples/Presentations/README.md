# Presentations

Reading a PowerPoint file: slide count, text extraction, and placeholder replacement.

```bash
dotnet run --project samples/Presentations
```

Prints the slide count and the first text body's text, before and after a replacement.

## The non-obvious part

**`ExtractText` returns one entry per text-bearing body, not one entry per slide.** A body is any
shape's `<p:txBody>`, a shape nested in a group, or a table cell's `<a:txBody>` — a title slide
alone is already two bodies, and a 2x2 table adds four more. Bodies come back in **deck order**
(`SlidesInDeckOrder` guarantees it), but **`.Count` is not the slide count** — call `SlideCount()`
for that.

**PowerPoint splits words across runs**, exactly as Word does. A single visible `{{who}}` is often
several `<a:t>` elements in the underlying XML, so a naive per-run `string.Replace` would miss it.
`ReplaceText` maps matches back onto the individual runs they overlap, which is what preserves
per-run formatting.

**There is no "create a PPTX from scratch" method**, so this sample reads a committed fixture —
the test project's `sample.pptx`, borrowed rather than duplicated.
