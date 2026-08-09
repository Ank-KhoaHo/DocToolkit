using DocToolkit;
using DocToolkit.Extensions.DependencyInjection;
using Microsoft.AspNetCore.Mvc;
using RazorPdf.Models;

namespace RazorPdf.Controllers;

/// <summary>
/// One Razor template, three outputs: the page itself, a PDF, and a Word document.
/// </summary>
/// <remarks>
/// The template is the point. An invoice layout that already exists as a view — reviewed, styled,
/// and maintained by whoever owns the branding — becomes a PDF without being rewritten into a
/// document-building API, and without a headless browser in the deployment.
///
/// Dependencies arrive by constructor here rather than as endpoint parameters, which is the only
/// real difference from the MinimalApi sample. Both resolve the same singletons.
/// </remarks>
[Route("invoice")]
public sealed class InvoiceController(
    RazorViewRenderer renderer,
    IHtmlToPdfConverter pdf,
    IHtmlToDocxConverter docx) : Controller
{
    private const string ViewPath = "~/Views/Invoice/Document.cshtml";

    private static readonly Invoice Sample = new(
        "INV-2026-0042",
        "Contoso Ltd",
        new DateOnly(2026, 8, 9),
        [
            new InvoiceLine("Widget", 2, 9.99m),
            new InvoiceLine("Gadget", 5, 9.00m),
            new InvoiceLine("Doohickey", 1, 7.50m),
        ]);

    /// <summary>The ordinary path: MVC renders the view to the response.</summary>
    [HttpGet("")]
    public IActionResult Page() => View(ViewPath, Sample);

    /// <summary>The same view, rendered to HTML first, then converted.</summary>
    [HttpGet("pdf")]
    public async Task<IActionResult> Pdf(CancellationToken ct)
    {
        var html = await renderer.RenderAsync(ViewPath, Sample);

        // Letter rather than the A4 default, to show where the choice goes. Landscape() and
        // WithMargins() are available here too - PageSetup is immutable, so these statics are safe
        // to read from a request handler.
        var bytes = await pdf.ConvertAsync(html, PageSetup.Letter, ct);

        return File(bytes, "application/pdf", $"{Sample.Number}.pdf");
    }

    /// <summary>And the same HTML as a Word document, for someone who needs to edit it.</summary>
    [HttpGet("docx")]
    public async Task<IActionResult> Docx(CancellationToken ct)
    {
        var html = await renderer.RenderAsync(ViewPath, Sample);
        var bytes = await docx.ConvertAsync(html, ct);

        return File(
            bytes,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            $"{Sample.Number}.docx");
    }
}
