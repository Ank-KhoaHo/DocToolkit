# HTML conversion

Converting **HTML to DOCX**, **HTML to PDF**, and **DOCX to PDF**.

```bash
dotnet run --project samples/HtmlConversion
```

Prints the byte count each conversion produced, then does the same on a landscape US Letter page,
and finishes by handing a text file to a DOCX reader to show what a failure looks like.

## The non-obvious part

**HTML → PDF pivots through DOCX.** No permissively-licensed, NuGet-only, Linux-safe library
renders HTML to PDF directly — the only free renderers *are* browsers, and a browser is a native
binary this package will not take on. `HtmlToPdfConverter` therefore composes the other two
converters rather than doing anything of its own.

That is why all three conversions live in one sample: split across three folders, the relationship
between them is invisible.

**`PageSetup` is immutable, and the sample checks it rather than asserting it.** `Landscape()` and
`WithMargins()` return new instances, so `PageSetup.Letter` is still portrait after the sample has
derived a landscape page from it — which is what makes those static properties safe to share
across a running application. The last line of that section prints the check.

**Every failure arrives as one exception type.** `DocumentConversionException` wraps whatever the
underlying library threw, so a caller writes one `catch` instead of one per dependency. The
original is kept as `InnerException` — here, a `FileFormatException`.
