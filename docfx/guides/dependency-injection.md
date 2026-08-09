# Dependency injection

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

One call registers ten interfaces, each a thin wrapper over the matching static class:

| Interface | Wraps |
|---|---|
| @DocToolkit.Extensions.DependencyInjection.IHtmlToDocxConverter | `HtmlToDocxConverter` |
| @DocToolkit.Extensions.DependencyInjection.IHtmlToPdfConverter | `HtmlToPdfConverter` |
| @DocToolkit.Extensions.DependencyInjection.IDocxToPdfConverter | `DocxToPdfConverter` |
| @DocToolkit.Extensions.DependencyInjection.IXlsxToPdfConverter | `XlsxToPdfConverter` |
| @DocToolkit.Extensions.DependencyInjection.IPptxToPdfConverter | `PptxToPdfConverter` |
| @DocToolkit.Extensions.DependencyInjection.IDocxToHtmlConverter | `DocxToHtmlConverter` |
| @DocToolkit.Extensions.DependencyInjection.IDocxToMarkdownConverter | `DocxToMarkdownConverter` |
| @DocToolkit.Extensions.DependencyInjection.IDocxEditor | `DocxEditor` |
| @DocToolkit.Extensions.DependencyInjection.IWorkbookEditor | `WorkbookEditor` |
| @DocToolkit.Extensions.DependencyInjection.IPresentationEditor | `PresentationEditor` |

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

### Options are read per call, not captured

The services consume options through `IOptionsMonitor` and read them **on each call**, not once at
construction. A singleton that captured `IOptions` at startup would freeze whatever configuration
was present then — and since these services are singletons, that freeze would last the lifetime of
the process. Reading through the monitor means an `appsettings.json` change with `reloadOnChange`
takes effect on the next conversion, with no restart.

The practical consequence: turning remote images off in a running service actually turns them off.

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
