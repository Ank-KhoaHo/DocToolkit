# Razor to PDF

Turning a **Razor view you already have** into a PDF or a Word document, from an MVC controller.

```bash
dotnet run --project samples/RazorPdf
```

Then:

| URL | What you get |
|---|---|
| `/invoice` | The view, rendered as a web page the ordinary way |
| `/invoice/pdf` | The same view as a PDF (~179 KB, US Letter) |
| `/invoice/docx` | The same view as a Word document, table intact |

One template, three outputs, no headless browser anywhere in the deployment.

## The non-obvious part

**Rendering a view to a `string` is not something MVC gives you.** Views render into the response
stream, so getting the HTML in hand means driving `IRazorViewEngine` yourself. That is
`RazorViewRenderer.cs` — about twenty lines, and the same twenty lines whatever you want the HTML
for: email bodies, snapshots, or a document.

Two details in it are worth knowing before you copy it:

- **`GetView` by explicit path, not `FindView` by name.** `FindView` resolves through the route
  values of the action currently executing. There is no action executing when a background job
  wants a document, and relying on one couples your rendering to your routing.
- **The `ActionContext` is a stand-in.** Rendering demands one, but nothing in a document template
  reads the request, so an empty `DefaultHttpContext` carrying the application's `RequestServices`
  is enough. That is also exactly what lets the same renderer run from a worker with no request in
  flight.

**Keep the CSS inline.** `<link rel="stylesheet">` is never fetched — nothing in this library opens
a socket you did not ask it to — so an external stylesheet silently does nothing. Inline `<style>`
and `style=` attributes are honoured, which is what the view uses.

**Layout is bounded by what Word can express**, because HTML → PDF pivots through DOCX. Tables,
headings, lists, inline styling and images survive; flexbox, grid, floats and absolute positioning
do not. An invoice or a statement is well inside that; a marketing page is not.

## How this differs from `MinimalApi`

`MinimalApi` answers "how do I wire this into DI and return a file". This one answers "how do I use
the template I already maintain". The DI difference is only that dependencies arrive by constructor
rather than as endpoint parameters — both resolve the same singletons.
