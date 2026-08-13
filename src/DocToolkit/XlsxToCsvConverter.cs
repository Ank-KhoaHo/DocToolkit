using System.Text;

namespace DocToolkit;

/// <summary>
/// Exports one sheet of a workbook as CSV (RFC 4180).
/// </summary>
/// <remarks>
/// <b>The output does not depend on the machine's regional settings.</b> That is a correctness
/// requirement here rather than a preference: <c>ReadSheet</c> renders a number through
/// <see cref="System.Globalization.CultureInfo.CurrentCulture"/>, which turns <c>1234.5</c> into
/// <c>1234,5</c> on a German machine — a decimal comma inside a comma-delimited file. Measured
/// across en-US, de-DE and fr-FR before this class was written. Numbers are invariant, dates are
/// ISO 8601, and a formula cell exports its computed <b>value</b>.
///
/// <b>One sheet, named explicitly.</b> CSV has no concept of a workbook, so there is nothing
/// sensible to do with the other sheets and no default worth guessing —
/// <see cref="WorkbookEditor.SheetNames(byte[])"/> lists them.
/// </remarks>
public static class XlsxToCsvConverter
{
    private const string FailureMessage = "Failed to convert XLSX to CSV.";

    /// <summary>Exports <paramref name="sheetName"/> as CSV.</summary>
    /// <param name="xlsx">The workbook to read.</param>
    /// <param name="sheetName">The sheet to export.</param>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="xlsx"/> is empty, or <paramref name="sheetName"/> is blank.
    /// </exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, or the sheet does not exist.
    /// </exception>
    public static string Convert(byte[] xlsx, string sheetName)
        => Format(WorkbookEditor.ReadSheetInvariant(xlsx, sheetName));

    /// <summary>
    /// Reads a workbook from <paramref name="source"/> and exports <paramref name="sheetName"/> as
    /// CSV.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and is neither disposed, closed nor
    /// sought, so it may be forward-only.
    /// </summary>
    /// <param name="source">The stream the workbook is read from.</param>
    /// <param name="sheetName">The sheet to export.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="sheetName"/>
    /// is blank.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The workbook could not be opened, or the sheet does not exist.
    /// </exception>
    public static async Task<string> ConvertAsync(
        Stream source, string sheetName, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ct.ThrowIfCancellationRequested();

        using var xlsx = await StreamPipeline
            .DrainAsync(source, "XLSX content was empty.", nameof(source), FailureMessage, ct)
            .ConfigureAwait(false);

        return Convert(xlsx.ToArray(), sheetName);
    }

    /// <summary>
    /// Renders a grid as RFC 4180 CSV: <c>CRLF</c> between records, and a field quoted only when
    /// it has to be.
    /// </summary>
    /// <remarks>
    /// <c>CRLF</c> rather than <c>\n</c> because RFC 4180 specifies it, and because a bare
    /// <c>\n</c> inside a quoted field would otherwise be indistinguishable from a record
    /// separator to a reader that splits on it.
    /// </remarks>
    private static string Format(IReadOnlyList<IReadOnlyList<string>> rows)
    {
        var csv = new StringBuilder();

        foreach (var row in rows)
        {
            for (var i = 0; i < row.Count; i++)
            {
                if (i > 0) csv.Append(',');
                Append(csv, row[i]);
            }

            csv.Append("\r\n");
        }

        return csv.ToString();
    }

    /// <summary>
    /// Quotes a field only when it contains a comma, a quote or a line break, and escapes a quote
    /// by doubling it.
    /// </summary>
    /// <remarks>
    /// Quoting everything unconditionally would also be valid CSV and is tempting because it needs
    /// no decision — but it makes every numeric column arrive as text in Excel and in most
    /// importers, which is a worse default for the format's main use.
    /// </remarks>
    private static void Append(StringBuilder csv, string field)
    {
        var mustQuote = field.AsSpan().IndexOfAny(',', '"', '\n') >= 0 || field.Contains('\r');

        if (!mustQuote)
        {
            csv.Append(field);
            return;
        }

        csv.Append('"');
        foreach (var c in field)
        {
            if (c == '"') csv.Append('"');
            csv.Append(c);
        }

        csv.Append('"');
    }
}
