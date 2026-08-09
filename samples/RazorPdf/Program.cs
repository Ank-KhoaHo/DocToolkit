using DocToolkit.Extensions.DependencyInjection;
using RazorPdf;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllersWithViews();
builder.Services.AddDocToolkit();

// The renderer holds no state; it resolves the view engine per call. Singleton is fine.
builder.Services.AddSingleton<RazorViewRenderer>();

var app = builder.Build();

app.MapControllers();

app.MapGet("/", () => Results.Redirect("/invoice"));

app.Run();
