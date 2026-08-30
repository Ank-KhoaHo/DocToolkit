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

    public byte[] Protect(byte[] pptx, string password)
        => DocToolkit.PresentationEditor.Protect(pptx, password);

    public byte[] Unprotect(byte[] pptx, string password)
        => DocToolkit.PresentationEditor.Unprotect(pptx, password);

    public bool IsProtected(byte[] pptx) => DocToolkit.PresentationEditor.IsProtected(pptx);

    public Task ProtectAsync(Stream source, Stream destination, string password, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.ProtectAsync(source, destination, password, ct);

    public Task UnprotectAsync(Stream source, Stream destination, string password, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.UnprotectAsync(source, destination, password, ct);

    public IReadOnlyList<string> ReadSlide(byte[] pptx, int index)
        => DocToolkit.PresentationEditor.ReadSlide(pptx, index);

    public Task<IReadOnlyList<string>> ReadSlideAsync(Stream source, int index, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.ReadSlideAsync(source, index, ct);

    public byte[] RemoveSlides(byte[] pptx, IEnumerable<int> indices)
        => DocToolkit.PresentationEditor.RemoveSlides(pptx, indices);

    public Task RemoveSlidesAsync(
        Stream source, IEnumerable<int> indices, Stream destination, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.RemoveSlidesAsync(source, indices, destination, ct);

    public byte[] ReorderSlides(byte[] pptx, IEnumerable<int> order)
        => DocToolkit.PresentationEditor.ReorderSlides(pptx, order);

    public Task ReorderSlidesAsync(
        Stream source, IEnumerable<int> order, Stream destination, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.ReorderSlidesAsync(source, order, destination, ct);

    public byte[] InsertSlides(byte[] pptx, int atIndex, IEnumerable<DocToolkit.PptxSlide> slides)
        => DocToolkit.PresentationEditor.InsertSlides(pptx, atIndex, slides);

    public Task InsertSlidesAsync(
        Stream source, int atIndex, IEnumerable<DocToolkit.PptxSlide> slides, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.PresentationEditor.InsertSlidesAsync(source, atIndex, slides, destination, ct);
}
