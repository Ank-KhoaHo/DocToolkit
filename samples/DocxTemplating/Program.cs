using DocToolkit;

Console.WriteLine("DOCX templating");
Console.WriteLine("===============");

// --- Scalars ------------------------------------------------------------------------------
// ReplaceText handles a placeholder even when Word has split it across several runs, which it
// routinely does - {{customer}} is often three separate <w:t> elements.

byte[] template = await HtmlToDocxConverter.ConvertAsync("<p>Customer: {{customer}}</p>");
byte[] filled = DocxEditor.ReplaceText(template, new Dictionary<string, string>
{
    ["{{customer}}"] = "Contoso Ltd",
});

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

// ROWS FIRST, THEN SCALARS. Expanding clones the template row, so any scalar substituted
// beforehand is duplicated into every line. This ordering is the whole reason these two
// operations are demonstrated together - see README.md.
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

string invoiceText = DocxEditor.ExtractText(invoice);
string[] descriptions = { "Widget", "Gadget", "Doohickey" };
int lineCount = descriptions.Count(invoiceText.Contains);

Console.WriteLine($"Line items   : {lineCount} rows from one template row");
Console.WriteLine($"Customer set : {invoiceText.Contains("Contoso Ltd")}");
Console.WriteLine($"Placeholders left over: {invoiceText.Contains("{{item.")}");

Console.WriteLine("\nDone.");
