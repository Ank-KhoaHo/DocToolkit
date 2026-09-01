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

    public byte[] Format(byte[] xlsx, string sheetName, DocToolkit.XlsxFormat format)
        => DocToolkit.WorkbookEditor.Format(xlsx, sheetName, format);

    public Task FormatAsync(
        Stream source, string sheetName, DocToolkit.XlsxFormat format, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.FormatAsync(source, sheetName, format, destination, ct);

    public byte[] Protect(byte[] xlsx, string password)
        => DocToolkit.WorkbookEditor.Protect(xlsx, password);

    public byte[] Unprotect(byte[] xlsx, string password)
        => DocToolkit.WorkbookEditor.Unprotect(xlsx, password);

    public bool IsProtected(byte[] xlsx) => DocToolkit.WorkbookEditor.IsProtected(xlsx);

    public Task ProtectAsync(Stream source, Stream destination, string password, CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.ProtectAsync(source, destination, password, ct);

    public Task UnprotectAsync(Stream source, Stream destination, string password, CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.UnprotectAsync(source, destination, password, ct);

    public DocToolkit.DocumentSignatureInfo InspectSignatures(byte[] xlsx)
        => DocToolkit.WorkbookEditor.InspectSignatures(xlsx);

    public Task<DocToolkit.DocumentSignatureInfo> InspectSignaturesAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.InspectSignaturesAsync(source, ct);

    public DocToolkit.DocumentSignatureValidationReport ValidateSignatures(byte[] xlsx, DocToolkit.DocumentSignatureValidationOptions? options = null)
        => DocToolkit.WorkbookEditor.ValidateSignatures(xlsx, options);

    public Task<DocToolkit.DocumentSignatureValidationReport> ValidateSignaturesAsync(
        Stream source, DocToolkit.DocumentSignatureValidationOptions? options = null, CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.ValidateSignaturesAsync(source, options, ct);

    public DocToolkit.DocumentMetadata ReadMetadata(byte[] xlsx)
        => DocToolkit.WorkbookEditor.ReadMetadata(xlsx);

    public byte[] WithMetadata(byte[] xlsx, DocToolkit.DocumentMetadata metadata)
        => DocToolkit.WorkbookEditor.WithMetadata(xlsx, metadata);

    public DocToolkit.XlsxFormulaInspection InspectFormulas(byte[] xlsx)
        => DocToolkit.WorkbookEditor.InspectFormulas(xlsx);

    public byte[] EvaluateFormulas(byte[] xlsx)
        => DocToolkit.WorkbookEditor.EvaluateFormulas(xlsx);

    public byte[] AddChart(
        byte[] xlsx, string sheetName, string cellRef, DocToolkit.ChartType type, DocToolkit.ChartData data,
        string title = "", int widthPixels = 640, int heightPixels = 360)
        => DocToolkit.WorkbookEditor.AddChart(xlsx, sheetName, cellRef, type, data, title, widthPixels, heightPixels);

    public Task AddChartAsync(
        Stream source, string sheetName, string cellRef, DocToolkit.ChartType type, DocToolkit.ChartData data, Stream destination,
        string title = "", int widthPixels = 640, int heightPixels = 360, CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.AddChartAsync(source, sheetName, cellRef, type, data, destination, title, widthPixels, heightPixels, ct);

    public byte[] AddPivotTable(
        byte[] xlsx, string sheetName, string sourceRange, string destinationCell, string name,
        IEnumerable<string> rowFields, IEnumerable<DocToolkit.PivotDataField> dataFields,
        IEnumerable<string>? columnFields = null, IEnumerable<string>? pageFields = null,
        bool showRowGrandTotals = true, bool showColumnGrandTotals = true)
        => DocToolkit.WorkbookEditor.AddPivotTable(
            xlsx, sheetName, sourceRange, destinationCell, name, rowFields, dataFields,
            columnFields, pageFields, showRowGrandTotals, showColumnGrandTotals);

    public Task AddPivotTableAsync(
        Stream source, string sheetName, string sourceRange, string destinationCell, string name,
        IEnumerable<string> rowFields, IEnumerable<DocToolkit.PivotDataField> dataFields, Stream destination,
        IEnumerable<string>? columnFields = null, IEnumerable<string>? pageFields = null,
        bool showRowGrandTotals = true, bool showColumnGrandTotals = true, CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.AddPivotTableAsync(
            source, sheetName, sourceRange, destinationCell, name, rowFields, dataFields, destination,
            columnFields, pageFields, showRowGrandTotals, showColumnGrandTotals, ct);

    public byte[] AddDefinedName(byte[] xlsx, string name, string sheetName, string range)
        => DocToolkit.WorkbookEditor.AddDefinedName(xlsx, name, sheetName, range);

    public Task AddDefinedNameAsync(
        Stream source, string name, string sheetName, string range, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.AddDefinedNameAsync(source, name, sheetName, range, destination, ct);

    public byte[] AddImage(
        byte[] xlsx, string sheetName, string cellRef, byte[] image,
        int? widthPixels = null, int? heightPixels = null)
        => DocToolkit.WorkbookEditor.AddImage(xlsx, sheetName, cellRef, image, widthPixels, heightPixels);

    public Task AddImageAsync(
        Stream source, string sheetName, string cellRef, byte[] image, Stream destination,
        int? widthPixels = null, int? heightPixels = null, CancellationToken ct = default)
        => DocToolkit.WorkbookEditor.AddImageAsync(
            source, sheetName, cellRef, image, destination, widthPixels, heightPixels, ct);
}
