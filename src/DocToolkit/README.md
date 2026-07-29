# DocToolkit

Convert HTML to DOCX and PDF, and open/edit DOCX, XLSX and PPTX from .NET.

**Pure managed.** No native binaries, no browser, no LibreOffice, no Office interop.
Works after `dotnet restore` alone, and runs on Linux.

## Install

```bash
dotnet add package DocToolkit
```

Targets `net8.0` and `net10.0`.

## Usage

```csharp
using DocToolkit;

// HTML -> DOCX
byte[] docx = await HtmlToDocxConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

// HTML -> PDF (pivots through DOCX internally)
byte[] pdf = await HtmlToPdfConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 18,100.00</p>");

// DOCX -> PDF
byte[] rendered = DocxToPdfConverter.Convert(docx);

// Fill a DOCX template
byte[] filled = DocxEditor.ReplaceText(docx, new Dictionary<string, string>
{
    ["{{customer}}"] = "Contoso Ltd",
});
string text = DocxEditor.ExtractText(filled);

// Spreadsheets
byte[] xlsx = WorkbookEditor.Create("Sales", new[]
{
    new object?[] { "Region", "Total" },
    new object?[] { "North", 1200 },
});
string cell = WorkbookEditor.ReadCell(xlsx, "Sales", "B2");
byte[] updated = WorkbookEditor.SetCell(xlsx, "Sales", "B2", 1500);

// Presentations
int slides = PresentationEditor.SlideCount(pptx);
IReadOnlyList<string> slideText = PresentationEditor.ExtractText(pptx);
byte[] editedPptx = PresentationEditor.ReplaceText(pptx, new Dictionary<string, string>
{
    ["{{title}}"] = "Q3 Results",
});
```

## Why HTML to PDF goes through DOCX

No permissively-licensed, NuGet-only library renders HTML to PDF on Linux: the only free
renderers are browsers, and a browser is a native binary. Pivoting through DOCX keeps the
whole chain pure managed.

## Licence

MIT. See `THIRD-PARTY-NOTICES.txt` for dependency attribution.
