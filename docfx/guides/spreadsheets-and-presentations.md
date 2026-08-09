# Spreadsheets and presentations

@DocToolkit.WorkbookEditor covers XLSX and @DocToolkit.PresentationEditor covers PPTX. Both follow
the same conventions as the Word surface: static methods, `byte[]` by default, `Stream` and path
overloads where you need them.

## Creating a workbook

[!code-csharp[](../../samples/Spreadsheets/Program.cs#create)]

Cell values are `object?`. Numbers are written as numbers and strings as strings, so a column of
totals arrives in Excel as something you can sum rather than as text that merely looks numeric.
`ReadCell` returns a `string` — the cell's value as displayed — because that is what almost every
caller does with it next.

## Reading a workbook you were handed

You do not need to know the shape in advance.

[!code-csharp[](../../samples/Spreadsheets/Program.cs#read)]

`SheetNames` tells you what is in the file and `ReadSheet` returns the used range as rows of
strings. Together they are enough to walk an upload you have never seen.

## Several sheets, and formulas between them

[!code-csharp[](../../samples/Spreadsheets/Program.cs#multi-sheet)]

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

## Presentations

`PresentationEditor.Create` builds a deck from a typed model — no template file involved.

[!code-csharp[](../../samples/Presentations/Program.cs#create)]

@DocToolkit.PptxSlide.Titled gives you a title and any number of bullets, emitted as real title and
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
— charts, conditional formatting, some shape effects — are **dropped silently**. There is no
warning channel on the public API. The PDF is valid; it just will not have the chart in it. If a
chart is the point of the document, render it to an image yourself and place that.

## Choosing a starting point

| You have | Use |
|---|---|
| A file somebody made in Excel or PowerPoint | `SetCell` / `AppendRows` / `ReplaceText` — edit in place, keep the formatting |
| Data, and no file | `Create` — describe the content, let the library write the file |
| A file you did not write and have never seen | `SheetNames` + `ReadSheet`, or `SlideCount` + `ExtractText` |
