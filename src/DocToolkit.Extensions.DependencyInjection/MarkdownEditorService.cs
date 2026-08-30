namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IMarkdownEditor"/>, delegating to <see cref="DocToolkit.MarkdownEditor"/>.</summary>
internal sealed class MarkdownEditorService : IMarkdownEditor
{
    public IReadOnlyDictionary<string, object> ReadFrontMatter(string markdown) =>
        DocToolkit.MarkdownEditor.ReadFrontMatter(markdown);

    public DocToolkit.MarkdownHeading? FindHeading(
        string markdown, string headingText,
        StringComparison comparison = StringComparison.OrdinalIgnoreCase) =>
        DocToolkit.MarkdownEditor.FindHeading(markdown, headingText, comparison);

    public int TableCount(string markdown) => DocToolkit.MarkdownEditor.TableCount(markdown);

    public IReadOnlyList<IReadOnlyList<string>> ReadTable(string markdown, int index) =>
        DocToolkit.MarkdownEditor.ReadTable(markdown, index);

    public string ReplaceSection(
        string markdown, string headingText, string newContent,
        StringComparison comparison = StringComparison.OrdinalIgnoreCase) =>
        DocToolkit.MarkdownEditor.ReplaceSection(markdown, headingText, newContent, comparison);
}
