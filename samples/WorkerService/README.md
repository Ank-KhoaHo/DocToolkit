# Worker service

Generating documents from a **background job** — the nightly-export shape — with
`Microsoft.NET.Sdk.Worker` and `AddDocToolkit()`.

```bash
dotnet run --project samples/WorkerService
```

Writes `revenue.xlsx` and `revenue.pdf` into `bin/.../reports/`, logs where they went, and stops
the host so the command returns.

## The non-obvious part

**A `BackgroundService` can inject these directly. It does not need a scope.**

The standard advice for worker services is that a hosted service is a singleton and therefore
cannot take a dependency on anything scoped — you inject `IServiceScopeFactory`, open a scope per
unit of work, and resolve from that. The advice is correct, and it exists because `DbContext` and
most repository types are scoped: capturing one in a singleton either throws at startup or, worse,
silently hands you a single instance for the lifetime of the process.

It does not apply here. Everything `AddDocToolkit()` registers is a **stateless singleton** — each
method takes its input and returns its output and holds nothing between calls — so they inject
straight into the worker's constructor. Wrapping them in a scope would be ceremony that suggests a
lifetime problem which does not exist.

**`IOptionsMonitor` is the part that earns its keep here.** The services read their options on
every call rather than capturing them at construction. In a request that lasts 40 ms that is a
detail; in a process that runs for weeks it is the difference between a configuration change taking
effect and requiring a restart. Turning remote images off in a running exporter actually turns them
off.

## Why one package reference is pinned and the other floats

`Ank.DocToolkit.Extensions.DependencyInjection` uses `Version="*"`, like every sample — that is
what makes these projects a canary for a breaking change in *this* repository's package.

`Microsoft.Extensions.Hosting` is pinned to `8.0.1`. The Worker SDK does not carry the hosting
types the way the Web SDK does, so a worker project has to name the package — and floating
somebody else's package would turn an unrelated upstream release into a red build on a sample,
which is noise rather than signal.
