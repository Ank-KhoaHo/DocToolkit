---
description: Read and write XLSX and PPTX from .NET, and render either to PDF, without Excel or PowerPoint installed.
---

# Read and write XLSX and PPTX in .NET without Excel or PowerPoint

@DocToolkit.WorkbookEditor covers XLSX and @DocToolkit.PresentationEditor covers PPTX. Both follow
the same conventions as the Word surface: static methods, `byte[]` by default, `Stream` and path
overloads where you need them.

## Creating a workbook

[!code-csharp[](../../../samples/Spreadsheets/Program.cs#create)]

Cell values are `object?`. Numbers are written as numbers and strings as strings, so a column of
totals arrives in Excel as something you can sum rather than as text that merely looks numeric.
`ReadCell` returns a `string`, because that is what almost every caller does with it next.

> [!NOTE]
> **It is the cell's value as text, not the cell as Excel displays it.** A cell holding `1200`
> under a `#,##0.00` number format reads back as `1200`, not `1,200.00` — the number format is a
> presentation instruction stored beside the value, and this returns the value. The text also
> follows the calling thread's `CurrentCulture`, which is why the exporters below deliberately do
> not use it.

## Reading a workbook you were handed

You do not need to know the shape in advance.

[!code-csharp[](../../../samples/Spreadsheets/Program.cs#read)]

`SheetNames` tells you what is in the file and `ReadSheet` returns the used range as rows of
strings. Together they are enough to walk an upload you have never seen.

## Several sheets, and formulas between them

[!code-csharp[](../../../samples/Spreadsheets/Program.cs#multi-sheet)]

Two details in there are worth calling out.

**A cell holding an @DocToolkit.XlsxFormula is written as a formula**, not as text that starts with
`=`. `XlsxFormula.From("SUM(Q1!B2:B4)")` is how you say "this is a computation" — and note the
formula has no leading `=`, which the writer adds.

**Reading that formula back through this library gives you the computed value.** The underlying
engine evaluates it, because a file written this way carries no cached result. A reader that only
reads cached values — which is most of them — sees an empty cell until Excel has opened and saved
the file. This is a genuinely surprising interop failure, and it is worth knowing which side of it
you are on.

**Two operations exist for the reader that is neither this library nor Excel.**
@DocToolkit.WorkbookEditor.InspectFormulas* reports which formulas the underlying engine actually
understands, before you trust a value — each carries its own reason when it does not.
@DocToolkit.WorkbookEditor.EvaluateFormulas* writes the computed value into the file itself, for a
reader that will not recalculate on its own — a spreadsheet viewer that only trusts a cached value,
or a parser that reads the XML directly.

[!code-csharp[](../../../samples/Spreadsheets/Program.cs#formula-evaluate)]

`AppendRows` adds after the sheet's last used row and leaves every other sheet untouched, which is
the operation you want for a log or an export that accumulates.

## Making a generated sheet look like a report

A sheet written from data is correct and unreadable: no header emphasis, columns too narrow for
their contents, and the header scrolling out of sight on the first flick of the wheel.
`WorkbookEditor.Format` fixes those, and more besides — a number format or an explicit width per
column, a freeze at any position, an autofilter, conditional formats and data validations.

The count that used to sit in that sentence said "eight" and then listed five, which is the shape of
mistake this documentation keeps making: a number in prose is a claim nothing verifies. The settings
are enumerable from @DocToolkit.XlsxFormat itself, so they are named here and counted nowhere.

**The boundary is a CLOSED vocabulary rather than a small one.** Six rule conditions
([XlsxRuleKind](xref:DocToolkit.XlsxRuleKind)), five validation kinds ([XlsxValidationKind](xref:DocToolkit.XlsxValidationKind)) and four
highlights ([XlsxHighlight](xref:DocToolkit.XlsxHighlight)) — each enumerable, measured and guaranteed.
`XlsxHighlight` names an *intent* rather than a colour on purpose: a colour picker cannot be
enumerated, and the moment one exists the boundary is gone. If what you need falls outside a closed
set — arbitrary fonts, borders, colour scales — use ClosedXML directly rather than through a thinner
API.

The snippet below starts from that preset and adds five of them. Every setting is a `With…` call on
the same immutable object, so they compose in any order — with one rule worth knowing, visible in
the snippet: an explicit width is applied *after* auto-fit, so naming one wins for that column while
every other column is still sized to its contents.

[!code-csharp[](../../../samples/Spreadsheets/Program.cs#format)]

```text
Formatted    : 7,880 bytes (was 7,351)
```

@DocToolkit.XlsxFormat is **immutable**, like @DocToolkit.PageSetup — every `With…` returns a new
instance, so `XlsxFormat.Report` is safe to read from anywhere. `Report` is the combination you
almost always want (bold header, frozen header, auto-fit); `XlsxFormat.None` is the empty one to
build up from.

Two things worth knowing before you reach for it:

**It applies to a sheet that already exists.** `Format` is not an argument to `Create` or
`AppendRows` — it is a separate call taking a workbook and a sheet name. That is what lets it
compose with all of them, and with a file somebody else made.

**Auto-fit does not always widen.** It fits the column to its *content*, and against short values
that is narrower than Excel's 8.43-character default. A column of two-digit numbers gets narrower,
which is correct and is not what "auto-fit" makes most people picture.

The number-format string is Excel's own (`"#,##0.00"`, `"0%"`, `"yyyy-mm-dd"`), keyed by column
letter. It changes how Excel *displays* the cell and not what is stored in it — so `ReadCell` and
the exporters below still see `1200`, per the note further up.

## A real Excel table, rather than a range that looks like one

Banding a range and putting a filter on it makes something that *looks* like Excel's "Format as
Table". A table (a **ListObject**) is the real thing: a named object Excel treats as a unit, which
structured references like `Revenue[Revenue]` can point at.

[!code-csharp[](../../../samples/Spreadsheets/Program.cs#table)]

**Two limitations, both measured rather than assumed.**

`WithTable` and `WithAutoFilter` cannot target overlapping ranges. A table carries its own
autofilter, so applying the sheet-wide one on top throws
@DocToolkit.DocumentConversionException rather than silently picking one. Use one or the other over
the same cells — which is why the snippet above is its own `Format` call rather than another line
on the previous one.

**`WithTable` does not make `AppendRows` keep the table current**, and this is worth stating plainly
because the opposite is the intuitive guess. `AppendRows` writes a raw cell value at the sheet's
last used row; it has no awareness of any table, and ClosedXML does not retroactively absorb an
adjacent write into a table's range. Measured directly: appending after a table leaves the table's
own range and row count unchanged, with the new row sitting *next to* it rather than inside it.
Recreate the table over the new range if you need it to cover the new rows.

[XlsxTableStyle](xref:DocToolkit.XlsxTableStyle) is four tiers — `None`, `Light`, `Medium`, `Dark` —
for the same closed-vocabulary reason `XlsxHighlight` names an intent instead of a colour.

## Print setup, merged cells, hyperlinks and comments

The rest of a worksheet's own furniture, all through the same immutable `XlsxFormat`.

[!code-csharp[](../../../samples/Spreadsheets/Program.cs#worksheet-furniture)]

@DocToolkit.XlsxPageSetup is a **different type** from @DocToolkit.PageSetup, and that is deliberate
rather than duplication: the DOCX one describes paper for a document, while a worksheet's print
setup is about which cells print and which rows repeat at the top of each printed page. The two
formats do not agree on what a page is, so they do not share a type — the same reasoning that keeps
`PdfMetadata` separate from `DocumentMetadata`.

That matters most for PDF. A workbook this package writes carried no page setup at all before this
existed, so `XlsxToPdfConverter` inherited whatever the reader defaulted to — the same class of
defect as the missing `w:sectPr` that 0.13.0 fixed for DOCX.

Two rules the factories enforce rather than document:

- A hyperlink's URL must be **absolute**. A relative one is refused at
  [XlsxHyperlink.To](xref:DocToolkit.XlsxHyperlink.To(System.String,System.String)) rather than written and left silently broken for
  whoever opens the file.
- A range or cell must **not** name a sheet. `"Other!A1"` is refused rather than quietly retargeted,
  because `Format`'s own `sheetName` argument already chose the sheet, and honouring both would mean
  deciding which of two contradictory instructions wins.

## Named ranges and images

Both are workbook *edits* rather than presentation, so they are their own
@DocToolkit.WorkbookEditor methods rather than `XlsxFormat` members.

[!code-csharp[](../../../samples/Spreadsheets/Program.cs#defined-name-and-image)]

**The sheet name is always single-quoted in the reference `AddDefinedName` writes**, whether or not
it needs to be. Measured: a sheet name containing a space, left unquoted, raises no error at write
time — the defined name is simply *absent* when the file is reopened, with nothing telling the
caller why. Quoting a name that does not need it is harmless, so there is no conditional here to get
wrong.

**Image sizes are pixels**, matching `AddChart` rather than the points the DOCX and PPTX drawing
model uses, because pixels are what ClosedXML's own picture sizing takes. Give neither dimension and
the image's intrinsic size is used; give one and the other scales to preserve the aspect ratio; give
both and you get exactly those, distortion accepted as your choice.

Format is decided by **magic bytes, never a filename** — PNG and JPEG only. A file named `.png` that
holds JPEG bytes is read as the JPEG it actually is, which is the opposite of the silent blank frame
a filename-trusting reader produces.

## Handing a sheet to something that is not Excel

@DocToolkit.XlsxToCsvConverter and @DocToolkit.XlsxToHtmlConverter take one sheet, by name.

[!code-csharp[](../../../samples/Spreadsheets/Program.cs#export)]

```text
As CSV       : Region,Revenue / EMEA,1200 / APAC,980 / AMER,1450 / 
As HTML      : 222 chars, starts "<table>
```

**One sheet at a time is the API, not a limitation to work around.** A workbook is not one table,
and neither CSV nor an HTML `<table>` has any way to say "and now a different sheet". Call it once
per name from `SheetNames`.

**Both are culture-invariant, deliberately, and this one is not a preference.** A machine set to
`de-DE` formats `1234.5` as `1234,5` — and a decimal comma inside a comma-delimited file is not a
differently-formatted CSV, it is a corrupt one, with a row that silently gained a column. So the
exporters format numbers, dates (ISO 8601) and booleans invariantly regardless of the calling
thread, which is the same line `SetCell` already holds on the way in.

The CSV is RFC 4180 and quotes **only when it has to** — a value containing a comma, a quote or a
newline — so the common case stays diffable. The HTML is a `<table>` fragment with a `<thead>`, not
a document: there is no `<html>` or `<body>` around it, because the caller is embedding it in a
page they already have. Every cell is escaped.

## Pivot tables

`WorkbookEditor.AddPivotTable` creates a pivot table from existing sheet data:

[!code-csharp[](../../../samples/Spreadsheets/Program.cs#pivot)]

```text
Pivot cell D1 right after creation: ""
```

**The result grid is empty until Excel opens and recalculates it.** That is a harder version of
the formula caveat above (*Several sheets, and formulas between them*): a formula's value **is**
computed by `ReadCell`/`ReadSheet` on read, and, since `XlsxToPdfConverter` started calling
`Calculate()` before rendering, is now computed there too. A pivot aggregation is not, on any of
those paths: nothing in this library evaluates one. Reading the pivot's own cells back with
`ReadCell`/`ReadSheet` immediately after this call returns empty strings, and `XlsxToPdfConverter`
renders nothing where the pivot's results would be. Open the result in Excel (or an equivalent) to
see it populated.

## Presentations

`PresentationEditor.Create` builds a deck from a typed model — no template file involved.

[!code-csharp[](../../../samples/Presentations/Program.cs#create)]

[`PptxSlide.Titled`](xref:DocToolkit.PptxSlide.Titled*) gives you a title and any number of bullets, emitted as real title and
body placeholders rather than free-floating text boxes. That matters if anyone opens the deck in
PowerPoint afterwards: placeholders inherit the theme, respond to layout changes, and appear in the
outline view.

> [!WARNING]
> `ExtractText` returns one entry per text-bearing body, **not one per slide**. `Create` emits two
> shapes per slide — a title and a content placeholder — so a two-slide deck reports four bodies.
> Use `SlideCount` for the slide count; `ExtractText(...).Count` is not it.

`ReplaceText` works the same way it does for Word documents, so a deck can be a template too:

```csharp
byte[] edited = PresentationEditor.ReplaceText(pptx, new Dictionary<string, string>
{
    ["{{who}}"] = "World",
});
```

### Charts

`WorkbookEditor.AddChart` and `PresentationEditor.AddChart` create charts, sharing one
`ChartType`/`ChartData` model:

[!code-csharp[](../../../samples/Spreadsheets/Program.cs#chart)]

```text
With chart   : <varies> bytes (does not touch the sheet's cell data)
```

[!code-csharp[](../../../samples/Presentations/Program.cs#chart)]

```text
With chart   : <varies> bytes (reaches PptxToPdfConverter's output)
```

DOCX chart creation is not included in this version — `OfficeIMO.Word`'s chart API has a
structurally different shape from the one Excel and PowerPoint share, and forcing it into the same
model would under- or over-serve one side. Word charts may get their own API in a future version.

Both calls above reach the render: `PptxToPdfConverter` and `XlsxToPdfConverter` carry the chart
through, title and category labels included — see *Rendering either one to PDF* below for what is
measured and what is not.

### SmartArt is read, not (yet) written

`ReadSmartArt` returns each diagram's node texts on a slide, one entry per diagram, each entry
already newline-joined:

[!code-csharp[](../../../samples/Presentations/Program.cs#smartart)]

```text
SmartArt     : 1 diagram(s) on slide 1
Diagram text : "Plan / Build / Ship"
In ExtractText too: True
```

**A SmartArt diagram's text lives in a different OOXML construct entirely** — a diagram data
part, not a text-bearing shape body — which is why it needs its own method rather than showing up
in `ReadSlide`. `ExtractText` includes it too, appended after that slide's ordinary text, for the
same reason.

There is no `AddSmartArt` on this package's own API yet: creating one through the usual
`byte[]`-in/`byte[]`-out shape is measured to have a rendering gap — content added that way does
not currently reach `PptxToPdfConverter`'s output, so it is left out rather than shipped with a
silent surprise. The sample above builds its demonstration deck with `OfficeIMO.PowerPoint`
directly for exactly that reason — it is the same escape hatch a caller reaches for, not a shortcut
unique to this guide. A deck authored in PowerPoint itself reads back correctly either way.

## Rendering either one to PDF

@DocToolkit.XlsxToPdfConverter and @DocToolkit.PptxToPdfConverter mirror
@DocToolkit.DocxToPdfConverter exactly — the same three members, the same behaviour.

```csharp
byte[] fromSheet = XlsxToPdfConverter.Convert(xlsx);
byte[] fromDeck  = PptxToPdfConverter.Convert(pptx);

XlsxToPdfConverter.ConvertFile("report.xlsx", "report.pdf");
await PptxToPdfConverter.ConvertAsync(source, destination, ct);
```

The same fidelity limit applies as everywhere else: features the rendering engine cannot represent
— conditional formatting, some shape effects — are **dropped silently**, with no warning channel on
the public API. **A chart is not one of those drops, when it was added by this library.** A chart
created with `WorkbookEditor.AddChart` or `PresentationEditor.AddChart` (see *Charts* above) renders
correctly here, title and category labels included, measured directly rather than assumed. A chart
authored some other way — directly in Excel or PowerPoint, or through `OfficeIMO` — and merely
present in the source file is a different, unmeasured case; if that is your situation, render it to
an image yourself and place that instead.

**A formula cell renders its computed value here too, not its source text.** `XlsxToPdfConverter`
calls `Calculate()` before rendering, the same evaluation `ReadCell`/`ReadSheet` already do on
read — so `=SUM(A1:B1)` renders as `42`, not as the literal text `SUM(A1:B1)`. A pivot table's
result grid is a different, unfixed case: see *Pivot tables* above for why nothing renders there
until Excel has recalculated the file.

### Legacy `.ppt` decks

`PptxToPdfConverter.Convert` also reads **PowerPoint 97-2003 binary decks**. No separate call and no
conversion step — hand it the bytes.

**It succeeds on 60.2% of them**, measured over 88 real decks from a government crawl, which is a
lower bar than the OOXML path and is published rather than rounded up to "supported". The refusals
are dominated by one upstream limitation, and **none of them produces a corrupt PDF** — a deck that
cannot be read is refused, not rendered blank.

So it is worth pointing at an archive of old decks, and worth checking the result rather than
assuming it. `XlsxToPdfConverter` has no equivalent: a legacy `.xls` workbook is **refused**, with a
message saying so.

## Document properties

What a file manager shows in its properties panel, and what a search indexer reads.
@DocToolkit.DocumentMetadata is shared across XLSX, PPTX and DOCX — the same type comes back from
@DocToolkit.WorkbookEditor.ReadMetadata* and @DocToolkit.PresentationEditor.ReadMetadata*, since
the three formats' property bags are identical in shape. See
[Document properties](word-documents.md#document-properties) for the DOCX side, and note it is
deliberately **not** shared with the PDF-specific `PdfMetadata` — the two ecosystems use the word
`Creator` for two different things, and forcing one type to cover both would collide them.

[!code-csharp[](../../../samples/Spreadsheets/Program.cs#metadata)]

```text
After retitling: title "Superseded", creator still "Contoso Finance"
```

[!code-csharp[](../../../samples/Presentations/Program.cs#metadata)]

```text
After retitling: title "Superseded", creator still "Contoso Finance"
```

**Every property is `null` when absent, not empty**, and a `null` passed to `WithMetadata` leaves
whatever the document already had alone — retitling a workbook does not silently erase its author.
Pass an empty string to actually clear a value.

## Choosing a starting point

| You have | Use |
|---|---|
| A file somebody made in Excel or PowerPoint | `SetCell` / `AppendRows` / `ReplaceText` — edit in place, keep the formatting |
| Data, and no file | `Create` — describe the content, let the library write the file |
| A file you did not write and have never seen | `SheetNames` + `ReadSheet`, or `SlideCount` + `ExtractText` |
