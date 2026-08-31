# Presentations

Building a PowerPoint file from scratch, then reading it back: slide count, text extraction,
placeholder replacement, adding a chart, and reading a SmartArt diagram's text.

```bash
dotnet run --project samples/Presentations
```

Creates a two-slide deck from a typed model, prints its slide count and first text body before and
after a replacement, adds a chart and a SmartArt diagram, then writes the result — replacement,
chart and SmartArt all present — as `deck.pptx` next to the built binary.

## The non-obvious part

**`ExtractText` returns one entry per text-bearing body, not one entry per slide.** A body is any
shape's `<p:txBody>`, a shape nested in a group, or a table cell's `<a:txBody>` — a title slide
alone is already two bodies, and a 2x2 table adds four more. Bodies come back in **deck order**
(`SlidesInDeckOrder` guarantees it), but **`.Count` is not the slide count** — call `SlideCount()`
for that. This sample shows it directly: `Create` emits a title shape and a content shape per
slide, so its two slides print `Bodies : 4`.

**PowerPoint splits words across runs**, exactly as Word does. A single visible `{{who}}` is often
several `<a:t>` elements in the underlying XML, so a naive per-run `string.Replace` would miss it.
`ReplaceText` maps matches back onto the individual runs they overlap, which is what preserves
per-run formatting. This sample does *not* exercise that case: a deck `Create` just built has one
run per line, so its `{{who}}` is a single `<a:t>`. The split case is covered by the test suite,
against decks PowerPoint itself wrote.

**`Create` takes data, not a template.** A `PptxSlide` is a title and its bullet lines; there is no
source file to edit and nothing to copy. That is why this sample needs no fixture — it used to
borrow the test project's `sample.pptx`, before `Create` existed.

**There is no `AddSmartArt` on this library's own API.** Creating one through the usual
`byte[]`-in/`byte[]`-out shape is measured to have a rendering gap — content added that way does
not currently reach `PptxToPdfConverter`'s output — so it is left out rather than shipped with a
silent surprise. This sample builds its SmartArt-bearing deck with `OfficeIMO.PowerPoint` directly
instead, the same library `ReadSmartArt` itself reads with — the same escape hatch a real caller
has, not a shortcut unique to this sample.
