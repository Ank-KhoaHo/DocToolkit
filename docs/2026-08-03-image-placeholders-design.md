# Image placeholders — design

Backlog item **A5**, from `2026-08-03-enhancement-backlog.md`.

## Why

`DocxEditor` can now fill scalars (`ReplaceText`) and repeat table rows (`FillRows`). A template
still cannot carry a logo, a signature or a QR code, which is the remaining third of "fill a Word
template" — the use case the DOCX path is strongest at. An invoice with line items but no company
logo is not a finished invoice.

This completes a describable capability rather than adding a third unrelated feature.

## Scope

**In:** replacing a text placeholder with an inline image, in the body, headers, footers, footnotes
and endnotes, for PNG and JPEG.

**Out, deliberately:** image *extraction* or replacement of images already in the document, floating
or text-wrapped positioning, cropping, rotation, alt-text beyond a default, and any format needing
a decoder. Each is a separate decision, and none is needed to put a logo in a header.

## The constraint that shapes this

**Every image in a DOCX carries an explicit size** (`wp:extent`, in EMUs). Something must supply
that number, and the obvious answer — an image-decoding library — is expensive here. `SixLabors.
ImageSharp` is the natural candidate and its later majors moved to the same revenue-gated Six Labors
Split License that `SixLabors.Fonts` is pinned at `[1.0.x]` to avoid. Taking it would mean a second
permanent CI guard against a licence the project exists to stay clear of.

**PNG and JPEG dimensions are readable from the file header in managed code.** PNG stores width and
height as big-endian `uint32` at a fixed offset inside the IHDR chunk. JPEG needs a short walk over
segment markers to the Start-Of-Frame, where height and width follow the precision byte.

Both were verified against real bytes before this was written, not assumed:

```
JPEG (docProps/thumbnail.jpeg from the existing pptx fixture): 256 x 144
PNG  (hand-built IHDR):                                          2 x 3
```

That is the same trade the project already made when it replaced ShapeCrawler with raw
`DocumentFormat.OpenXml`: write the managed code, keep the dependency graph clean.

## Public API

```csharp
public static byte[] ReplaceImage(
    byte[] docx,
    string placeholder,
    byte[] image,
    double? widthPoints = null,
    double? heightPoints = null);

public static Task ReplaceImageAsync(
    Stream source,
    string placeholder,
    byte[] image,
    Stream destination,
    double? widthPoints = null,
    double? heightPoints = null,
    CancellationToken ct = default);
```

Purely additive. Both overloads share a single `ReplaceImageCore`, so they cannot drift.

**`placeholder` is the literal text including braces** — `"{{logo}}"` — matching `ReplaceText`
rather than `FillRows`. `FillRows` uses bare field names only because the collection name is already
an argument there and repeating it would be redundant; no such redundancy exists here.

**Sizing is in points**, not EMUs or pixels, because points are what a document author thinks in.
Omit both and the intrinsic pixel size is used at 96 DPI. Give one and the other scales to preserve
the aspect ratio. Give both and the image is stretched to fit — distortion is the caller's choice,
not an error.

**One image per call.** Several images means several calls, each reopening the package. That cost is
real and small against how many images a template actually holds, and it avoids adding a public
options record and a second way to express "an image". `FillRows` composes with `ReplaceText` the
same way.

## Semantics

- The image replaces the placeholder text **inline, within its own paragraph**, so alignment,
  indentation and surrounding text survive.
- **Only the matched span is removed.** Text sharing a run with the placeholder stays, and keeps its
  formatting — `Signed: {{signature}} (authorised)` becomes `Signed: ` + image + ` (authorised)`,
  not a paragraph containing only an image. This is the same principle `RunTextSplicer` enforces for
  text, applied to an element insertion.
- **PNG and JPEG only.** Anything else throws, naming what was actually found rather than saying
  "invalid image". GIF and BMP can be added later without changing the signature.
- **Every occurrence** of the placeholder gets an image, including across headers and footers.
- **No occurrence throws** `DocumentConversionException` — consistent with `FillRows`, and for the
  same reason: a call that matches nothing is a bug in the call or the template, not a no-op.
- Each inserted image gets a `wp:docPr` **name** derived from the placeholder with its braces and
  surrounding whitespace stripped — `{{logo}}` becomes `logo` — so the accessibility pane and the
  selection pane show something meaningful rather than "Picture 1". The `descr` (alt text) gets the
  same value. Callers wanting real alt text can set it afterwards; that is out of scope here.

## Implementation

Add an `ImagePart` to the part that owns the paragraph, build the DrawingML inline element with
`DocumentFormat.OpenXml`, and splice it in where the placeholder text was.

**Legacy VML (`w:pict`) was considered and rejected.** Its XML is far shorter, and `DocxFixtures`
already builds VML shapes for the text-box tests so there is precedent in the repo. But VML is
deprecated, Word emits DrawingML for anything modern, and a signature written as VML would look
subtly unlike one a human inserted.

**This cannot reuse `RunTextSplicer`, unlike `FillRows`.** That helper maps match offsets back onto
runs and writes *text*. Here the matched span must be removed and an *element* inserted at that
position, so the splicing is genuinely new code — built on the same principle of never flattening
runs that a match does not overlap.

## Traps

Every one of these produces a document that **opens**, which is why each gets a test.

**The image part must belong to the part that owns the paragraph.** A logo in a header needs
`HeaderPart.AddImagePart`, not `MainDocumentPart.AddImagePart`. Get it wrong and the relationship ID
resolves in the wrong scope: Word opens the file and shows nothing where the image should be. This
is where most of the header/footer complexity lives, and it is invisible to any test that only reads
text back.

**`wp:docPr/@id` must be unique across the whole document.** Duplicates make Word declare the file
corrupt and offer to repair it. The implementation scans existing `DocProperties` for the maximum
and counts up from there rather than starting at 1 and hoping.

**The content type must match the actual bytes.** `ImagePartType.Png` holding JPEG data yields a
part Word cannot render — no error, just a blank frame. Format is detected from the magic bytes, not
from any filename or caller assertion.

**The placeholder can be split across runs.** Word splits a visible word across `w:t` elements
routinely; the same problem `RunTextSplicer` exists for, needing a different solution here.

**The arithmetic is easy to get subtly wrong.** 1 point = 12,700 EMU. 1 pixel at 96 DPI = 9,525 EMU.
A factor-of-ten error produces an image that is merely the wrong size, which no schema check catches.

## Error handling

| Condition | Result |
|---|---|
| `docx`, `placeholder` or `image` is null | `ArgumentNullException` |
| `docx` or `image` is empty, or `placeholder` is blank | `ArgumentException` |
| `widthPoints` or `heightPoints` is zero or negative | `ArgumentOutOfRangeException` |
| Image is neither PNG nor JPEG | `DocumentConversionException`, naming the detected format |
| The package could not be opened or edited | `DocumentConversionException` |
| **No occurrence of the placeholder** | `DocumentConversionException` |
| `ct` was cancelled | `OperationCanceledException` |

## Testing

- **Schema validity via `OpenXmlValidator`** on every produced document. A4 taught this directly: a
  fixture built schema-invalid tables while every text assertion passed, so extracted text proves
  nothing about whether Word will open the file.
- **Intrinsic sizing from a PNG header, and from a JPEG header** — separate parsers, both proved.
- **One dimension given preserves the aspect ratio**; both given produces exactly those.
- **An image placed in a header is owned by the `HeaderPart`**, asserted by inspecting which part
  holds it. Without this the primary use case can regress invisibly.
- **A placeholder split across runs** is still matched.
- **Multiple occurrences each get an image, with distinct `docPr` IDs** — asserted directly, because
  duplicate IDs are what triggers Word's repair prompt.
- **An unsupported format throws**, and **no match throws**.
- **`ReplaceImage` and `ReplaceImageAsync` produce identical output** for identical input.

### Fixtures, without an image library

The tests need real PNG and JPEG bytes and cannot generate them with a decoder.

- **JPEG:** `tests/DocToolkit.Tests/assets/sample.pptx` already contains `docProps/thumbnail.jpeg` —
  3,935 bytes, 256 × 144, verified. Extract it at test time. Real bytes, already in the repo, no new
  binary fixture to commit and no `.gitattributes` question.
- **PNG:** hand-built — signature, IHDR, a stored-deflate IDAT, IEND, with a small CRC32 helper.
  Deliberately an unusual size (2 × 3) so an intrinsic-sizing assertion cannot pass by coincidence.

### Two repo-specific obligations

1. **`ReplaceImageAsync` must be added to the name lists at the top of `StreamOverloadTests`**, and
   the result count checked to rise. `CLAUDE.md`: an overload missing from those lists is the only
   way to escape the whole suite. Verified for `FillRowsAsync` by watching 83 → 90.
2. **`ReplaceImage` must be added to `AirGapGuardTests`.** It takes bytes and must never fetch
   anything — and that suite's value is being exhaustive across the public API, which `README.md`
   quantifies.

## Success criteria

- A `{{logo}}` placeholder in a header becomes an image owned by the `HeaderPart`, at its intrinsic
  size, in a document Word opens without repair.
- Supplying only `widthPoints` scales the height proportionally.
- Two occurrences produce two images with different `docPr` IDs.
- A GIF throws, naming GIF.
- `ReplaceImage` and `ReplaceImageAsync` agree.
- Build stays at 0 warnings under `-warnaserror`; the whole suite passes on both target frameworks.
