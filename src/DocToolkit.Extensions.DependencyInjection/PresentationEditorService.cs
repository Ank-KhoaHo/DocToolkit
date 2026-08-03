namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IPresentationEditor"/>, delegating to <see cref="DocToolkit.PresentationEditor"/>.</summary>
internal sealed class PresentationEditorService : IPresentationEditor
{
    public int SlideCount(byte[] pptx) => DocToolkit.PresentationEditor.SlideCount(pptx);

    public IReadOnlyList<string> ExtractText(byte[] pptx) => DocToolkit.PresentationEditor.ExtractText(pptx);

    public byte[] ReplaceText(byte[] pptx, IReadOnlyDictionary<string, string> replacements)
        => DocToolkit.PresentationEditor.ReplaceText(pptx, replacements);

    public Task<int> SlideCountAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.SlideCountAsync(source, ct);

    public Task<IReadOnlyList<string>> ExtractTextAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.ExtractTextAsync(source, ct);

    public Task ReplaceTextAsync(
        Stream source, IReadOnlyDictionary<string, string> replacements, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.PresentationEditor.ReplaceTextAsync(source, replacements, destination, ct);
}
