using DocToolkit.Extensions.DependencyInjection;
using WorkerService;

var builder = Host.CreateApplicationBuilder(args);

// Identical to the ASP.NET Core registration - AddDocToolkit knows nothing about the host it is
// in. Everything it registers is a stateless singleton, which is the detail that matters below.
builder.Services.AddDocToolkit(options =>
{
    // Bound from appsettings.json in a real service. Set here so the sample shows where an
    // application-wide policy goes: once, at startup, rather than at every call site.
    options.AllowRemoteImageDownload = false;
});

builder.Services.AddHostedService<ReportWorker>();

await builder.Build().RunAsync();
