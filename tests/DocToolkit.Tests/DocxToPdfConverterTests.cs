using System.Globalization;
using DocToolkit;
using Xunit;

namespace DocToolkit.Tests;

public class DocxToPdfConverterTests
{
    [Fact]
    public async Task Convert_ProducesAPdfContainingTheSourceText()
    {
        var docx = await HtmlToDocxConverter.ConvertAsync(
            "<h1>Invoice INV-42</h1><p>Total due: 18,100.00</p>");

        var pdf = DocxToPdfConverter.Convert(docx);

        Assert.True(PdfProbe.IsPdf(pdf));
        var text = PdfProbe.ExtractText(pdf);
        Assert.Contains("Invoice INV-42", text);
        Assert.Contains("18,100.00", text);
    }

    [Fact]
    public async Task Convert_PaginatesLongDocuments()
    {
        var rows = string.Concat(Enumerable.Range(1, 60).Select(i =>
            $"<tr><td>Line item {i} with a reasonably long description</td><td>{i * 950}</td></tr>"));
        var html = $"<h1>Big</h1><table border=\"1\">{rows}</table><p>END-MARKER</p>";

        var pdf = DocxToPdfConverter.Convert(await HtmlToDocxConverter.ConvertAsync(html));

        Assert.True(PdfProbe.PageCount(pdf) > 1, "expected the document to span multiple pages");
        Assert.Contains("END-MARKER", PdfProbe.ExtractText(pdf));
        Assert.DoesNotContain(PdfProbe.TextYPositions(pdf), y => y < 0);
    }

    [Fact]
    public void Convert_RejectsEmptyInput()
    {
        Assert.Throws<ArgumentException>(() => DocxToPdfConverter.Convert(Array.Empty<byte>()));
    }

    // ---- the render thread (A123): see DocxToPdfConverter.Render -----------------------------------

    [Fact]
    public void OnRenderThread_RunsTheWorkOnItsOwnThread_AndReturnsItsResult()
    {
        var caller = Environment.CurrentManagedThreadId;

        var (thread, name) = DocxToPdfConverter.OnRenderThread(
            () => (Environment.CurrentManagedThreadId, Thread.CurrentThread.Name));

        Assert.NotEqual(caller, thread);
        Assert.Equal("DocToolkit PDF render", name);
    }

    /// <summary>
    /// The refusal ladder matches on the exception the renderer threw, so the hop must hand back
    /// that same instance, with the render thread's frames still in its stack trace.
    /// </summary>
    [Fact]
    public void OnRenderThread_RethrowsTheSameException_WithItsOriginalStackTrace()
    {
        var thrown = new ArgumentException("Parameter 'text' cannot be empty or whitespace.");

        var caught = Assert.Throws<ArgumentException>(
            () => DocxToPdfConverter.OnRenderThread<int>(() => throw thrown));

        Assert.Same(thrown, caught);
        Assert.Contains("<OnRenderThread>b__", caught.StackTrace!, StringComparison.Ordinal);
    }

    /// <summary>
    /// <b>With the execution context suppressed</b>, because a new thread normally inherits the
    /// cultures through it anyway, and then this test could not tell the explicit hand-over from
    /// none.
    /// </summary>
    [Fact]
    public void OnRenderThread_CarriesTheCallersCultures_EvenWithoutTheExecutionContext()
    {
        var (culture, uiCulture) = (CultureInfo.CurrentCulture, CultureInfo.CurrentUICulture);
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("de-DE");
            CultureInfo.CurrentUICulture = new CultureInfo("fr-FR");

            (string, string) seen;
            using (ExecutionContext.SuppressFlow())
                seen = DocxToPdfConverter.OnRenderThread(
                    () => (CultureInfo.CurrentCulture.Name, CultureInfo.CurrentUICulture.Name));

            Assert.Equal(("de-DE", "fr-FR"), seen);
        }
        finally
        {
            CultureInfo.CurrentCulture = culture;
            CultureInfo.CurrentUICulture = uiCulture;
        }
    }
}
