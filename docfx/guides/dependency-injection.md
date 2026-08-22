---
description: Register DocToolkit with services.AddDocToolkit() for ASP.NET Core and worker services, one interface per capability.
---

# Register DocToolkit with ASP.NET Core dependency injection

The core library is static classes, on purpose: for most callers there is nothing to configure and
nothing to inject. But a static call is awkward to substitute in a test, and application-wide
settings have to be passed at every call site.

`Ank.DocToolkit.Extensions.DependencyInjection` fixes both without changing the core.

```bash
dotnet add package Ank.DocToolkit.Extensions.DependencyInjection
```

It brings the core package with it transitively — you do not need to reference both.

## Registering

[!code-csharp[](../../samples/MinimalApi/Program.cs#register)]

One call registers every interface the package ships, each a thin wrapper over the matching
static class:

| Interface | Wraps |
|---|---|
| @DocToolkit.Extensions.DependencyInjection.IHtmlToDocxConverter | `HtmlToDocxConverter` |
| @DocToolkit.Extensions.DependencyInjection.IHtmlToPdfConverter | `HtmlToPdfConverter` |
| @DocToolkit.Extensions.DependencyInjection.IDocxToPdfConverter | `DocxToPdfConverter` |
| @DocToolkit.Extensions.DependencyInjection.IXlsxToPdfConverter | `XlsxToPdfConverter` |
| @DocToolkit.Extensions.DependencyInjection.IPptxToPdfConverter | `PptxToPdfConverter` |
| @DocToolkit.Extensions.DependencyInjection.IDocxToHtmlConverter | `DocxToHtmlConverter` |
| @DocToolkit.Extensions.DependencyInjection.IDocxToMarkdownConverter | `DocxToMarkdownConverter` |
| @DocToolkit.Extensions.DependencyInjection.IMarkdownToDocxConverter | `MarkdownToDocxConverter` |
| @DocToolkit.Extensions.DependencyInjection.IMarkdownToPdfConverter | `MarkdownToPdfConverter` |
| @DocToolkit.Extensions.DependencyInjection.IXlsxToCsvConverter | `XlsxToCsvConverter` |
| @DocToolkit.Extensions.DependencyInjection.IXlsxToHtmlConverter | `XlsxToHtmlConverter` |
| @DocToolkit.Extensions.DependencyInjection.IDocToDocxConverter | `DocToDocxConverter` |
| @DocToolkit.Extensions.DependencyInjection.IDocxEditor | `DocxEditor` |
| @DocToolkit.Extensions.DependencyInjection.IWorkbookEditor | `WorkbookEditor` |
| @DocToolkit.Extensions.DependencyInjection.IPresentationEditor | `PresentationEditor` |
| @DocToolkit.Extensions.DependencyInjection.IPdfEditor | `PdfEditor` |

All are registered as **singletons**, which is safe because none of them hold state — every
operation takes its input and returns its output. They are also registered with `TryAdd`, so your
own implementation of any of these interfaces, registered first, wins.

## Injecting

[!code-csharp[](../../samples/MinimalApi/Program.cs#inject)]

Constructor injection works the same way in a controller, a worker service, or anywhere else the
container reaches.

## Configuration

@DocToolkit.Extensions.DependencyInjection.DocToolkitOptions is where an application-wide setting
goes, so it is not repeated at every call site.

```csharp
builder.Services.AddDocToolkit(options =>
{
    options.AllowRemoteImageDownload = true;
    options.RemoteImage.AllowedHosts.Add("assets.contoso.example");
    options.RemoteImage.Timeout = TimeSpan.FromSeconds(5);
    options.RemoteImage.MaxBytesPerImage = 2 * 1024 * 1024;
});
```

`RemoteImage` is the same @DocToolkit.RemoteImageOptions the static API takes — see
[Images the HTML points at](html-to-word-and-pdf.md#images-the-html-points-at) for what the
allow-list actually guarantees. `AllowRemoteImageDownload` is the master switch; it is `false`
unless you set it.

Binding from configuration works as usual:

```csharp
builder.Services.Configure<DocToolkitOptions>(builder.Configuration.GetSection("DocToolkit"));
builder.Services.AddDocToolkit();
```

### Setting the paper once

@DocToolkit.Extensions.DependencyInjection.DocToolkitOptions.Page is the page every producer
lays out on when a call does not name one. Default `PageSetup.A4`, so leaving it alone
changes nothing.

```csharp
services.AddDocToolkit(o => o.Page = PageSetup.Letter);
```

It reaches both HTML converters and `IDocxEditor.Create`. An explicit page argument still
wins — a call that names a page is answering a narrower question than the configuration was.

Setting it alongside `AllowRemoteImageDownload` keeps both. That combination was impossible
before 0.18.0: the remote-image path could only lay out on A4, so enabling downloads would
have silently discarded the configured paper.

### Options are read per call, not captured

The services consume options through `IOptionsMonitor` and read them **on each call**, not once at
construction. A singleton that captured `IOptions` at startup would freeze whatever configuration
was present then — and since these services are singletons, that freeze would last the lifetime of
the process. Reading through the monitor means an `appsettings.json` change with `reloadOnChange`
takes effect on the next conversion, with no restart.

The practical consequence: turning remote images off in a running service actually turns them off.

## Background services

A `BackgroundService` is registered as a singleton, and the standard advice is that a singleton
cannot take a dependency on anything scoped — inject `IServiceScopeFactory`, open a scope per unit
of work, resolve from that. That advice is correct, and it exists because `DbContext` and most
repository types are scoped.

It does not apply to these. They are stateless singletons, so they inject straight into the
worker's constructor:

[!code-csharp[](../../samples/WorkerService/ReportWorker.cs#worker)]

Wrapping them in a scope would be ceremony implying a lifetime problem that is not there.

`IOptionsMonitor` is the one that earns its keep in a long-running process: because the services
read options per call, a configuration change reaches a worker that has been up for weeks without
a restart.

The [WorkerService sample](https://github.com/Ank-KhoaHo/DocToolkit/tree/main/samples/WorkerService)
is the whole thing running.

## Testing

The reason to inject rather than call statically. Substitute the interface:

```csharp
var converter = Substitute.For<IHtmlToPdfConverter>();
converter.ConvertAsync(Arg.Any<string>(), Arg.Any<CancellationToken>())
         .Returns(Task.FromResult(new byte[] { 1, 2, 3 }));

var sut = new InvoiceService(converter);
```

That keeps a test that is about *your* logic from spending real time rendering a real PDF. Tests
that are about the conversion should call the real thing — it needs no fixtures and no network.

## Which package do I want?

| Situation | Package |
|---|---|
| A console app, a script, a library | `Ank.DocToolkit` — call the static classes |
| ASP.NET Core, worker service, anything with a container | Both — register with `AddDocToolkit()` |
| You want one place to configure remote images | Both — that is what `DocToolkitOptions` is for |

The extensions package adds no capability. Everything it exposes is available statically; it exists
for substitutability and configuration.
