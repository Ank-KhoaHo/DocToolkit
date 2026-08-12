namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IPresentationEditor"/>, delegating to <see cref="DocToolkit.PresentationEditor"/>.</summary>
internal sealed class PresentationEditorService : IPresentationEditor
{
    public byte[] Create(IEnumerable<DocToolkit.PptxSlide> slides)
        => DocToolkit.PresentationEditor.Create(slides);

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

    public Task CreateAsync(
        IEnumerable<DocToolkit.PptxSlide> slides, Stream destination, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.CreateAsync(slides, destination, ct);

    public byte[] ReplaceImage(byte[] pptx, string placeholder, byte[] image)
        => DocToolkit.PresentationEditor.ReplaceImage(pptx, placeholder, image);

    public Task ReplaceImageAsync(
        Stream source, string placeholder, byte[] image, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.PresentationEditor.ReplaceImageAsync(source, placeholder, image, destination, ct);
}
