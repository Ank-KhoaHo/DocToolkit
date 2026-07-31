namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IDocxEditor"/>, delegating to <see cref="DocToolkit.DocxEditor"/>.</summary>
internal sealed class DocxEditorService : IDocxEditor
{
    public byte[] ReplaceText(byte[] docx, IReadOnlyDictionary<string, string> replacements)
        => DocToolkit.DocxEditor.ReplaceText(docx, replacements);

    public string ExtractText(byte[] docx) => DocToolkit.DocxEditor.ExtractText(docx);

    public string ExtractText(byte[] docx, bool includeHeadersAndFooters)
        => DocToolkit.DocxEditor.ExtractText(docx, includeHeadersAndFooters);
}
