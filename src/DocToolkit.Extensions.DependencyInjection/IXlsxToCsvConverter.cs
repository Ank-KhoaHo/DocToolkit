namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Exports one sheet of a workbook as CSV (RFC 4180). Registered by
/// <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.
/// </summary>
/// <remarks>
/// <b>The output does not depend on the machine's regional settings</b>, and that is a correctness
/// requirement rather than a preference: <see cref="IWorkbookEditor.ReadSheet(byte[], string)"/>
/// renders a number through the current culture, which turns <c>1234.5</c> into <c>1234,5</c> on a
/// German machine - a decimal comma inside a comma-delimited file. Numbers are invariant, dates are
/// ISO 8601, and a formula cell exports its computed <b>value</b>.
///
/// <b>One sheet, named explicitly.</b> CSV has no concept of a workbook, so there is nothing
/// sensible to do with the other sheets - <see cref="IWorkbookEditor.SheetNames(byte[])"/> lists
/// them.
/// </remarks>
public interface IXlsxToCsvConverter
{
    /// <summary>Exports <paramref name="sheetName"/> as CSV.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty, or the sheet name is blank.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The workbook could not be opened, or the sheet does not exist.
    /// </exception>
    string Convert(byte[] xlsx, string sheetName);

    /// <summary>
    /// Reads a workbook from <paramref name="source"/> and exports <paramref name="sheetName"/> as
    /// CSV. <paramref name="source"/> is read to its end and is not disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or the sheet name is blank.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The workbook could not be opened, or the sheet does not exist.
    /// </exception>
    Task<string> ConvertAsync(Stream source, string sheetName, CancellationToken ct = default);
}
