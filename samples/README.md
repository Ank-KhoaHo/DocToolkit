# Samples

Two runnable, standalone projects, each referencing the published NuGet packages (not this
repo's source) — the same restore an external consumer would get.

## ConsoleSample

Exercises all five core `Ank.DocToolkit` capabilities in one run: HTML→DOCX, HTML→PDF,
DOCX→PDF, DOCX template fill + text extraction, XLSX create/read/edit, and PPTX read.

```bash
dotnet run --project samples/ConsoleSample
```

Prints a short report for each step (byte counts, extracted text, a before/after cell value).

## MinimalApiSample

An ASP.NET Core minimal API demonstrating `services.AddDocToolkit()` — one endpoint per
injected interface. `byte[]` request/response fields are base64-encoded JSON strings, using
ASP.NET Core's built-in handling — no custom (de)serialization needed.

```bash
dotnet run --project samples/MinimalApiSample --urls http://127.0.0.1:5299
```

Then, in another terminal:

```bash
curl -X POST http://127.0.0.1:5299/html-to-docx \
  -H "Content-Type: application/json" \
  -d '{"html":"<h1>Hello</h1>"}' \
  -o output.docx
```

Other endpoints (`/html-to-pdf`, `/docx-to-pdf`, `/docx/extract-text`, `/xlsx/read-cell`,
`/pptx/slide-count`) take a `{"bytes":"<base64>"}` body (`/xlsx/read-cell` also takes `sheet`
and `cell` fields) — see `samples/MinimalApiSample/Program.cs` for the exact request shape of
each endpoint.
