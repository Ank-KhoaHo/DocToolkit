namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IWorkbookEditor"/>, delegating to <see cref="DocToolkit.WorkbookEditor"/>.</summary>
internal sealed class WorkbookEditorService : IWorkbookEditor
{
    public byte[] Create(string sheetName, IEnumerable<IEnumerable<object?>> rows)
        => DocToolkit.WorkbookEditor.Create(sheetName, rows);

    public string ReadCell(byte[] xlsx, string sheetName, string cellRef)
        => DocToolkit.WorkbookEditor.ReadCell(xlsx, sheetName, cellRef);

    public byte[] SetCell(byte[] xlsx, string sheetName, string cellRef, object? value)
        => DocToolkit.WorkbookEditor.SetCell(xlsx, sheetName, cellRef, value);

    public Task CreateAsync(
        string sheetName, IEnumerable<IEnumerable<object?>> rows, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.CreateAsync(sheetName, rows, destination, ct);

    public Task<string> ReadCellAsync(Stream source, string sheetName, string cellRef, CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.ReadCellAsync(source, sheetName, cellRef, ct);

    public Task SetCellAsync(
        Stream source, string sheetName, string cellRef, object? value, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.SetCellAsync(source, sheetName, cellRef, value, destination, ct);
}
