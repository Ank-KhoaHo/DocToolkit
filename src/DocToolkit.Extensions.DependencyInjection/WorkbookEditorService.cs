namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IWorkbookEditor"/>, delegating to <see cref="DocToolkit.WorkbookEditor"/>.</summary>
internal sealed class WorkbookEditorService : IWorkbookEditor
{
    public byte[] Create(string sheetName, IEnumerable<IEnumerable<object?>> rows)
        => DocToolkit.WorkbookEditor.Create(sheetName, rows);

    public string ReadCell(byte[] xlsx, string sheetName, string cellRef)
        => DocToolkit.WorkbookEditor.ReadCell(xlsx, sheetName, cellRef);

    public IReadOnlyList<string> SheetNames(byte[] xlsx)
        => DocToolkit.WorkbookEditor.SheetNames(xlsx);

    public IReadOnlyList<IReadOnlyList<string>> ReadSheet(byte[] xlsx, string sheetName)
        => DocToolkit.WorkbookEditor.ReadSheet(xlsx, sheetName);

    public byte[] SetCell(byte[] xlsx, string sheetName, string cellRef, object? value)
        => DocToolkit.WorkbookEditor.SetCell(xlsx, sheetName, cellRef, value);

    public Task CreateAsync(
        string sheetName, IEnumerable<IEnumerable<object?>> rows, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.CreateAsync(sheetName, rows, destination, ct);

    public Task<string> ReadCellAsync(Stream source, string sheetName, string cellRef, CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.ReadCellAsync(source, sheetName, cellRef, ct);

    public Task<IReadOnlyList<string>> SheetNamesAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.SheetNamesAsync(source, ct);

    public Task<IReadOnlyList<IReadOnlyList<string>>> ReadSheetAsync(
        Stream source, string sheetName, CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.ReadSheetAsync(source, sheetName, ct);

    public Task SetCellAsync(
        Stream source, string sheetName, string cellRef, object? value, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.SetCellAsync(source, sheetName, cellRef, value, destination, ct);

    public byte[] Create(IEnumerable<DocToolkit.XlsxSheet> sheets)
        => DocToolkit.WorkbookEditor.Create(sheets);

    public byte[] AppendRows(byte[] xlsx, string sheetName, IEnumerable<IEnumerable<object?>> rows)
        => DocToolkit.WorkbookEditor.AppendRows(xlsx, sheetName, rows);

    public Task CreateAsync(
        IEnumerable<DocToolkit.XlsxSheet> sheets, Stream destination, CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.CreateAsync(sheets, destination, ct);

    public Task AppendRowsAsync(
        Stream source, string sheetName, IEnumerable<IEnumerable<object?>> rows, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.AppendRowsAsync(source, sheetName, rows, destination, ct);
}
