namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Creates, reads and edits Excel (.xlsx) workbooks. Registered by <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.</summary>
public interface IWorkbookEditor
{
    /// <summary>Creates a workbook with one sheet populated from <paramref name="rows"/>.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="rows"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="sheetName"/> is blank, or a row is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be built.</exception>
    byte[] Create(string sheetName, IEnumerable<IEnumerable<object?>> rows);

    /// <summary>Reads a cell as a string. <paramref name="cellRef"/> is an A1-style reference.</summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty, or a name is blank.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be opened, the sheet does not exist, or the reference is not valid.</exception>
    string ReadCell(byte[] xlsx, string sheetName, string cellRef);

    /// <summary>Sets a cell and returns the updated workbook bytes.</summary>
    /// <exception cref="ArgumentNullException">Any argument other than <paramref name="value"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="xlsx"/> is empty, or a name is blank.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be opened, the sheet does not exist, or the reference is not valid.</exception>
    byte[] SetCell(byte[] xlsx, string sheetName, string cellRef, object? value);

    /// <summary>
    /// Builds a workbook with one sheet populated from <paramref name="rows"/> and writes it to
    /// <paramref name="destination"/>. See <see cref="Create"/> for the exact typing and culture
    /// rules applied to each cell. <paramref name="destination"/> is <b>written</b> and is not
    /// disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="rows"/> or <paramref name="destination"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="sheetName"/> is blank, a row is null, or <paramref name="destination"/> is
    /// not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">The workbook could not be built or written.</exception>
    Task CreateAsync(
        string sheetName, IEnumerable<IEnumerable<object?>> rows, Stream destination,
        CancellationToken ct = default);

    /// <summary>
    /// Reads a workbook from <paramref name="source"/> and returns a cell as a string.
    /// <paramref name="cellRef"/> is an A1-style reference. <paramref name="source"/> is
    /// <b>read</b> to its end and is neither disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or a name is blank.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    Task<string> ReadCellAsync(Stream source, string sheetName, string cellRef, CancellationToken ct = default);

    /// <summary>
    /// Reads a workbook from <paramref name="source"/>, sets one cell, and writes the result to
    /// <paramref name="destination"/>. <paramref name="cellRef"/> is an A1-style reference.
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought.
    /// </summary>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="source"/> or <paramref name="destination"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, a name is blank, or
    /// <paramref name="destination"/> is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// The workbook could not be opened, the sheet does not exist, or the reference is not valid.
    /// </exception>
    Task SetCellAsync(
        Stream source, string sheetName, string cellRef, object? value, Stream destination,
        CancellationToken ct = default);
}
