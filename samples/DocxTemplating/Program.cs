using DocToolkit;

Console.WriteLine("DOCX templating");
Console.WriteLine("===============");

// --- Scalars ------------------------------------------------------------------------------
// ReplaceText handles a placeholder even when Word has split it across several runs, which it
// routinely does - {{customer}} is often three separate <w:t> elements.

#region scalars
byte[] template = await HtmlToDocxConverter.ConvertAsync("<p>Customer: {{customer}}</p>");
byte[] filled = DocxEditor.ReplaceText(template, new Dictionary<string, string>
{
    ["{{customer}}"] = "Contoso Ltd",
});
#endregion

Console.WriteLine($"\nScalar fill  : \"{DocxEditor.ExtractText(filled).Trim()}\"");

// --- Repeating rows -----------------------------------------------------------------------
// A whole invoice: one template row becomes one row per line item, each keeping the template
// row's formatting.

byte[] invoiceTemplate = await HtmlToDocxConverter.ConvertAsync(
    """
    <h1>Invoice for {{customer}}</h1>
    <table border="1">
      <tr><th>Description</th><th>Qty</th><th>Total</th></tr>
      <tr><td>{{item.Desc}}</td><td>{{item.Qty}}</td><td>{{item.Total}}</td></tr>
    </table>
    """);

// ROWS FIRST, THEN SCALARS - the safe order in general, since expanding clones the template row
// and any scalar already inside it would be duplicated into every line. Here {{customer}} sits in
// the <h1> outside the row FillRows clones, so this particular ordering doesn't actually change
// the output - see README.md for why it's still the sample's default.
#region rows
byte[] withRows = DocxEditor.FillRows(invoiceTemplate, "item", new[]
{
    new Dictionary<string, string> { ["Desc"] = "Widget",    ["Qty"] = "2", ["Total"] = "19.98" },
    new Dictionary<string, string> { ["Desc"] = "Gadget",    ["Qty"] = "5", ["Total"] = "45.00" },
    new Dictionary<string, string> { ["Desc"] = "Doohickey", ["Qty"] = "1", ["Total"] = "7.50" },
});

byte[] invoice = DocxEditor.ReplaceText(withRows, new Dictionary<string, string>
{
    ["{{customer}}"] = "Contoso Ltd",
});
#endregion

string invoiceText = DocxEditor.ExtractText(invoice);
string[] descriptions = { "Widget", "Gadget", "Doohickey" };
int lineCount = descriptions.Count(invoiceText.Contains);

Console.WriteLine($"Line items   : {lineCount} rows from one template row");
Console.WriteLine($"Customer set : {invoiceText.Contains("Contoso Ltd")}");
Console.WriteLine($"Placeholders left over: {invoiceText.Contains("{{item.")}");

// --- When there is no template at all --------------------------------------------------------
// Templating starts from a file somebody made in Word. When the document's shape comes from your
// data instead, describe it as blocks and skip the round trip through HTML entirely.

#region blocks
byte[] report = DocxEditor.Create(
    new[]
    {
        DocxBlock.Heading("Quarterly report", 1),
        DocxBlock.Paragraph("Revenue by region, in thousands."),
        DocxBlock.Table(
            new[] { "Region", "Q1", "Q2" },
            new[]
            {
                new object?[] { "EMEA", 1200, 1310 },
                new object?[] { "APAC", 980, 1040 },
            }),
    },
    PageSetup.A4.WithMargins(54));
#endregion

Console.WriteLine($"\nFrom blocks  : {report.Length:N0} bytes, no template file involved");

// --- Publishing the result somewhere that is not Word ----------------------------------------
// The same filled document, as HTML for a web page and as Markdown for a diff-able record. Both
// read the DOCX; neither needs Word, a browser or a network.

#region export
string html = DocxToHtmlConverter.Convert(invoice);
string markdown = DocxToMarkdownConverter.Convert(invoice);
#endregion

Console.WriteLine($"As HTML      : {html.Length:N0} chars, has a <table>: {html.Contains("<table", StringComparison.OrdinalIgnoreCase)}");
Console.WriteLine($"As Markdown  : {markdown.Length:N0} chars, first line \"{markdown.Split('\n')[0].Trim()}\"");

// --- Headers and footers -----------------------------------------------------------------
// A header belongs to the PAGE, so it goes on the PageSetup - which is why every producer
// honours it without a new overload. The page number is a real field: written as text it would
// be fixed at generation time, correct on one page and wrong on every other.

#region headers
PageSetup withRunningHead = PageSetup.A4
    .WithHeader(DocxHeader.Text("Contoso Ltd"))
    .WithFooter(DocxHeader.Of(
        HeaderAlignment.Right,
        DocxHeaderSegment.Text("Page "),
        DocxHeaderSegment.PageNumber,
        DocxHeaderSegment.Text(" of "),
        DocxHeaderSegment.PageCount));

byte[] withHeader = DocxEditor.Create(
    new[] { DocxBlock.Paragraph("Body text.") }, withRunningHead);
#endregion

// Calling WithFirstPage is the switch that turns on a distinct first page, and null means BLANK
// on page one rather than 'use the ordinary one'. That is what makes a title page carrying
// nothing across it expressible at all.

#region headers-first-page
PageSetup withTitlePage = withRunningHead.WithFirstPage(
    header: null,
    footer: DocxHeader.Text("Confidential", HeaderAlignment.Center));

byte[] withCover = DocxEditor.Create(
    new[] { DocxBlock.Paragraph("Body text.") }, withTitlePage);
#endregion

string headerText = DocxEditor.ExtractText(withHeader, includeHeadersAndFooters: true);

Console.WriteLine($"\nWith header  : {withHeader.Length:N0} bytes");
Console.WriteLine($"Title page   : {withCover.Length:N0} bytes (page one carries no header)");
bool headerSurvived = headerText.Contains("Contoso Ltd", StringComparison.Ordinal);

Console.WriteLine($"Header found : {headerSurvived}");

// --- Metadata -----------------------------------------------------------------------------
// What a file manager shows in its properties panel, and what a search indexer reads - shared
// across DOCX, XLSX and PPTX; see the Spreadsheets sample for the same type on a workbook.

#region metadata
byte[] withMetadata = DocxEditor.WithMetadata(invoice, new DocumentMetadata
{
    Title = "Invoice for Contoso Ltd",
    Creator = "Billing",
});

DocumentMetadata readBack = DocxEditor.ReadMetadata(withMetadata);
#endregion

Console.WriteLine($"\nTitle          : {readBack.Title}");
Console.WriteLine($"Creator        : {readBack.Creator}");

byte[] retitled = DocxEditor.WithMetadata(withMetadata, new DocumentMetadata { Title = "Superseded" });

Console.WriteLine($"After retitling: title \"{DocxEditor.ReadMetadata(retitled).Title}\", "
    + $"creator still \"{DocxEditor.ReadMetadata(retitled).Creator}\"");

Console.WriteLine("\nDone.");
