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
}
