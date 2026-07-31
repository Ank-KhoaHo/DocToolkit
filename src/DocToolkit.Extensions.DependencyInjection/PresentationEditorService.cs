namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IPresentationEditor"/>, delegating to <see cref="DocToolkit.PresentationEditor"/>.</summary>
internal sealed class PresentationEditorService : IPresentationEditor
{
    public int SlideCount(byte[] pptx) => DocToolkit.PresentationEditor.SlideCount(pptx);

    public IReadOnlyList<string> ExtractText(byte[] pptx) => DocToolkit.PresentationEditor.ExtractText(pptx);

    public byte[] ReplaceText(byte[] pptx, IReadOnlyDictionary<string, string> replacements)
        => DocToolkit.PresentationEditor.ReplaceText(pptx, replacements);
}
