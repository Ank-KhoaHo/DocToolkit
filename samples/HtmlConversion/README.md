# HTML conversion

Converting **HTML to DOCX**, **HTML to PDF**, and **DOCX to PDF**.

```bash
dotnet run --project samples/HtmlConversion
```

Prints the byte count each conversion produced.

## The non-obvious part

**HTML → PDF pivots through DOCX.** No permissively-licensed, NuGet-only, Linux-safe library
renders HTML to PDF directly — the only free renderers *are* browsers, and a browser is a native
binary this package will not take on. `HtmlToPdfConverter` therefore composes the other two
converters rather than doing anything of its own.

That is why all three conversions live in one sample: split across three folders, the relationship
between them is invisible.
