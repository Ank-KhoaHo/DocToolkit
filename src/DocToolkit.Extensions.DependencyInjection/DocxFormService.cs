namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IDocxForm"/>, delegating to <see cref="DocToolkit.DocxForm"/>.</summary>
internal sealed class DocxFormService : IDocxForm
{
    public DocToolkit.DocxFormReport Inspect(
        byte[] docx, DocToolkit.DocxFormKey key = DocToolkit.DocxFormKey.TagThenAlias)
        => DocToolkit.DocxForm.Inspect(docx, key);

    public Task<DocToolkit.DocxFormReport> InspectAsync(
        Stream source, DocToolkit.DocxFormKey key = DocToolkit.DocxFormKey.TagThenAlias,
        CancellationToken ct = default)
        => DocToolkit.DocxForm.InspectAsync(source, key, ct);

    public DocToolkit.DocxFormValidation Validate(
        byte[] docx, IReadOnlyDictionary<string, DocToolkit.DocxFormValue> values,
        DocToolkit.DocxFormKey key = DocToolkit.DocxFormKey.TagThenAlias)
        => DocToolkit.DocxForm.Validate(docx, values, key);

    public Task<DocToolkit.DocxFormValidation> ValidateAsync(
        Stream source, IReadOnlyDictionary<string, DocToolkit.DocxFormValue> values,
        DocToolkit.DocxFormKey key = DocToolkit.DocxFormKey.TagThenAlias,
        CancellationToken ct = default)
        => DocToolkit.DocxForm.ValidateAsync(source, values, key, ct);

    public byte[] Fill(
        byte[] docx, IReadOnlyDictionary<string, DocToolkit.DocxFormValue> values,
        DocToolkit.DocxFormKey key = DocToolkit.DocxFormKey.TagThenAlias)
        => DocToolkit.DocxForm.Fill(docx, values, key);

    public Task FillAsync(
        Stream source, Stream destination,
        IReadOnlyDictionary<string, DocToolkit.DocxFormValue> values,
        DocToolkit.DocxFormKey key = DocToolkit.DocxFormKey.TagThenAlias,
        CancellationToken ct = default)
        => DocToolkit.DocxForm.FillAsync(source, destination, values, key, ct);
}
