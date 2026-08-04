using DocToolkit.Extensions.DependencyInjection;

var builder = WebApplication.CreateBuilder(args);
builder.Services.AddDocToolkit();

var app = builder.Build();

app.MapPost("/html-to-docx", async (IHtmlToDocxConverter converter, HtmlRequest request) =>
{
    byte[] docx = await converter.ConvertAsync(request.Html);
    return Results.File(docx, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "output.docx");
});

app.MapPost("/html-to-pdf", async (IHtmlToPdfConverter converter, HtmlRequest request) =>
{
    byte[] pdf = await converter.ConvertAsync(request.Html);
    return Results.File(pdf, "application/pdf", "output.pdf");
});

app.MapPost("/docx-to-pdf", (IDocxToPdfConverter converter, FileRequest request) =>
{
    byte[] pdf = converter.Convert(request.Bytes);
    return Results.File(pdf, "application/pdf", "output.pdf");
});

app.MapPost("/docx/extract-text", (IDocxEditor editor, FileRequest request) =>
{
    string text = editor.ExtractText(request.Bytes);
    return Results.Text(text);
});

app.MapPost("/xlsx/read-cell", (IWorkbookEditor editor, CellRequest request) =>
{
    string value = editor.ReadCell(request.Bytes, request.Sheet, request.Cell);
    return Results.Text(value);
});

app.MapPost("/pptx/slide-count", (IPresentationEditor editor, FileRequest request) =>
{
    int count = editor.SlideCount(request.Bytes);
    return Results.Text(count.ToString());
});

app.Run();

internal sealed record HtmlRequest(string Html);
internal sealed record FileRequest(byte[] Bytes);
internal sealed record CellRequest(byte[] Bytes, string Sheet, string Cell);
