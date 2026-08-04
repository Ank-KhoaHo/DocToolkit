# Minimal API

An ASP.NET Core minimal API demonstrating `services.AddDocToolkit()` — one endpoint per injected
interface.

```bash
dotnet run --project samples/MinimalApi --urls http://127.0.0.1:5299
```

Then, in another terminal:

```bash
curl -X POST http://127.0.0.1:5299/html-to-docx \
  -H "Content-Type: application/json" \
  -d '{"html":"<h1>Hello</h1>"}' \
  -o output.docx
```

`/html-to-pdf` takes the same `{"html":"..."}` body. The remaining endpoints (`/docx-to-pdf`,
`/docx/extract-text`, `/xlsx/read-cell`, `/pptx/slide-count`) take `{"bytes":"<base64>"}` instead —
`/xlsx/read-cell` also takes `sheet` and `cell`. See `Program.cs` for each endpoint's exact shape.

## The non-obvious part

**`byte[]` fields are base64-encoded JSON strings**, using ASP.NET Core's built-in handling. No
custom serialization is needed in either direction.

**This project references only `Ank.DocToolkit.Extensions.DependencyInjection`**, never the core
package directly. That is deliberate: it proves the core package arrives transitively, exactly as
it would in a consumer's project. Adding an explicit core reference would make the build pass even
if that transitive dependency broke.

**`AllowRemoteImageDownload` is configured once** at `AddDocToolkit(...)`, rather than being
decided per call as the static API does.
