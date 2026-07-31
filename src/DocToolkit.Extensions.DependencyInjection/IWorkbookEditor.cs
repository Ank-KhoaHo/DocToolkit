namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Creates, reads and edits Excel (.xlsx) workbooks. Registered by <c>AddDocToolkit</c> (see AddDocToolkit).</summary>
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
}
