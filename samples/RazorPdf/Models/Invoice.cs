namespace RazorPdf.Models;

public sealed record InvoiceLine(string Description, int Quantity, decimal UnitPrice)
{
    public decimal Total => Quantity * UnitPrice;
}

public sealed record Invoice(string Number, string Customer, DateOnly Issued, IReadOnlyList<InvoiceLine> Lines)
{
    public decimal Total => Lines.Sum(line => line.Total);
}
