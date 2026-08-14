namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Exports one sheet of a workbook as an HTML table. Registered by
/// <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.
/// </summary>
/// <remarks>
/// <b>A fragment, not a document</b> - a bare <c>&lt;table&gt;</c> with no <c>&lt;html&gt;</c>
/// wrapper, deliberately the opposite of <see cref="IDocxToHtmlConverter"/>. A sheet is a component
/// of a page rather than a page.
///
/// <b>Cell text is culture-invariant</b>, matching <see cref="IXlsxToCsvConverter"/> so the two
/// exporters cannot disagree about what a cell says. <b>Every cell is escaped</b> - a workbook is
/// untrusted input.
/// </remarks>
public interface IXlsxToHtmlConverter
{
    /// <summary>Exports <paramref name="sheetName"/> as an HTML <c>&lt;table&gt;</c> fragment.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="xlsx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty, or the sheet name is blank.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The workbook could not be opened, or the sheet does not exist.
    /// </exception>
    string Convert(byte[] xlsx, string sheetName);

    /// <summary>
    /// Reads a workbook from <paramref name="source"/> and exports <paramref name="sheetName"/> as
    /// an HTML <c>&lt;table&gt;</c> fragment. <paramref name="source"/> is read to its end and is
    /// not disposed, closed or sought.
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
