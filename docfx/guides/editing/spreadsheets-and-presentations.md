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

```csharp
var withPivot = WorkbookEditor.AddPivotTable(
    workbook, "Sales", "A1:C10", "E1", "RegionSummary",
    rowFields: new[] { "Region" },
    dataFields: new[] { new PivotDataField("Amount", PivotFunction.Sum) });
```

**The result grid is empty until Excel opens and recalculates it.** That is a harder version of
the formula caveat above (*Several sheets, and formulas between them*): a formula's value **is**
computed by `ReadCell`/`ReadSheet` on read, because this library's own engine evaluates it — a
pivot table's is not, because nothing that writes a workbook, this method included, computes a
pivot aggregation. Reading the pivot's own cells back with `ReadCell`/`ReadSheet` immediately after
this call returns empty strings, and `XlsxToPdfConverter` renders nothing where the pivot's results
would be, for the identical reason it renders a formula's literal text rather than its computed
value. Open the result in Excel (or an equivalent) to see it populated.

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

### SmartArt is read, not (yet) written

`ReadSmartArt` returns the text of every SmartArt diagram on a slide, one entry per diagram:

```csharp
IReadOnlyList<IReadOnlyList<string>> diagrams = PresentationEditor.ReadSmartArt(pptx, index: 1);
```

**A SmartArt diagram's text lives in a different OOXML construct entirely** — a diagram data
part, not a text-bearing shape body — which is why it needs its own method rather than showing up
in `ReadSlide`. `ExtractText` includes it too, appended after that slide's ordinary text, for the
same reason.

There is no `AddSmartArt` here yet: creating one through this package's usual `byte[]`-in/`byte[]`-out
shape is measured to have a rendering gap — content added that way does not currently reach
`PptxToPdfConverter`'s output, so it is left out rather than shipped with a silent surprise. A deck
authored in PowerPoint itself, or built with `OfficeIMO.PowerPoint` directly, reads back correctly.

### Charts

`WorkbookEditor.AddChart` and `PresentationEditor.AddChart` create charts in an existing workbook
or presentation, sharing one `ChartType`/`ChartData` model:

```csharp
var data = new ChartData(
    new[] { "North", "South" },
    new[] { new ChartSeries("Total", new double[] { 1200, 980 }) });

var xlsx = WorkbookEditor.AddChart(
    workbook, "Sheet1", "B2", ChartType.ColumnClustered, data, title: "Regional Totals");

var pptx = PresentationEditor.AddChart(
    presentation, 1, ChartType.ColumnClustered, data, title: "Regional Totals");
```

DOCX chart creation is not included in this version — `OfficeIMO.Word`'s chart API has a
structurally different shape from the one Excel and PowerPoint share, and forcing it into the same
model would under- or over-serve one side. Word charts may get their own API in a future version.

Unlike the SmartArt case above, this one **does** reach the render: `PptxToPdfConverter` and
`XlsxToPdfConverter` both carry the chart through, title and category labels included — see
*Rendering either one to PDF* below for what is measured and what is not.

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

## Choosing a starting point

| You have | Use |
|---|---|
| A file somebody made in Excel or PowerPoint | `SetCell` / `AppendRows` / `ReplaceText` — edit in place, keep the formatting |
| Data, and no file | `Create` — describe the content, let the library write the file |
| A file you did not write and have never seen | `SheetNames` + `ReadSheet`, or `SlideCount` + `ExtractText` |
