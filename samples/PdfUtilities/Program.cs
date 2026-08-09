using DocToolkit;

// The only part of this library that READS a PDF. Everything else renders into one and never looks
// at the result, which is why these operations arrived last.
//
// Nothing here re-renders. Pages move between documents as they are, so the fidelity caveats that
// apply to the converters do not apply to anything below - a page that came out of the renderer
// looking right still looks right after being merged, extracted and stamped.

Console.WriteLine("PDF utilities");
Console.WriteLine("=============");

byte[] cover = await HtmlToPdfConverter.ConvertAsync("<h1>Cover</h1><p>Contoso Ltd</p>");
byte[] invoice = await HtmlToPdfConverter.ConvertAsync("<h1>Invoice</h1><p>Total: 72.45</p>");
byte[] terms = await HtmlToPdfConverter.ConvertAsync("<h1>Terms</h1><p>Net 30.</p>");

// --- Merge, then take a page back out ---------------------------------------------------------

#region merge
byte[] bundle = PdfEditor.Merge([cover, invoice, terms]);

// firstPage is 1-based, the way a reader numbers pages rather than the way an array indexes them.
byte[] justTheInvoice = PdfEditor.ExtractPages(bundle, firstPage: 2, count: 1);
#endregion

Console.WriteLine($"\nThree documents : {PdfEditor.PageCount(cover)} + {PdfEditor.PageCount(invoice)} + {PdfEditor.PageCount(terms)} pages");
Console.WriteLine($"Merged          : {PdfEditor.PageCount(bundle)} pages, {bundle.Length:N0} bytes");
Console.WriteLine($"Page 2 alone    : {PdfEditor.PageCount(justTheInvoice)} page, {justTheInvoice.Length:N0} bytes");

// A range that is not entirely inside the document is refused rather than silently clamped, so a
// slice that runs off the end is a bug you hear about instead of a short document you do not.
try
{
    PdfEditor.ExtractPages(bundle, firstPage: 3, count: 2);
}
catch (ArgumentOutOfRangeException ex)
{
    Console.WriteLine($"Refused         : {ex.Message.Split('(')[0].Trim()}");
}

// --- Metadata ---------------------------------------------------------------------------------
// What a file manager shows in its properties panel, and what a search indexer reads.

#region metadata
byte[] stamped = PdfEditor.WithMetadata(bundle, new PdfMetadata
{
    Title = "Invoice INV-2026-0042",
    Author = "Contoso Ltd",
});

PdfMetadata read = PdfEditor.ReadMetadata(stamped);
#endregion

Console.WriteLine($"\nTitle           : {read.Title}");
Console.WriteLine($"Author          : {read.Author}");

// null means ABSENT, not blank - the distinction anything merging metadata from several sources
// depends on, since "no subject" should take a fallback and "a subject set to empty" should not.
Console.WriteLine($"Subject         : {read.Subject?.ToString() ?? "(null - never set)"}");

// And a null property leaves what the document already had alone, so stamping one field cannot
// quietly erase another.
byte[] retitled = PdfEditor.WithMetadata(stamped, new PdfMetadata { Title = "Superseded" });

Console.WriteLine($"After retitling : title \"{PdfEditor.ReadMetadata(retitled).Title}\", "
    + $"author still \"{PdfEditor.ReadMetadata(retitled).Author}\"");

Console.WriteLine($"Pages unchanged : {PdfEditor.PageCount(retitled)}");

Console.WriteLine("\nDone.");
