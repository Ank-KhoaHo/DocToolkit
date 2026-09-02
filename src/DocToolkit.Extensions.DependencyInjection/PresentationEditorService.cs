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

    public DocToolkit.DocumentSignatureInfo InspectSignatures(byte[] pptx)
        => DocToolkit.PresentationEditor.InspectSignatures(pptx);

    public Task<DocToolkit.DocumentSignatureInfo> InspectSignaturesAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.InspectSignaturesAsync(source, ct);

    public DocToolkit.DocumentSignatureValidationReport ValidateSignatures(byte[] pptx, DocToolkit.DocumentSignatureValidationOptions? options = null)
        => DocToolkit.PresentationEditor.ValidateSignatures(pptx, options);

    public Task<DocToolkit.DocumentSignatureValidationReport> ValidateSignaturesAsync(
        Stream source, DocToolkit.DocumentSignatureValidationOptions? options = null, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.ValidateSignaturesAsync(source, options, ct);

    public DocToolkit.DocumentMetadata ReadMetadata(byte[] pptx)
        => DocToolkit.PresentationEditor.ReadMetadata(pptx);

    public byte[] WithMetadata(byte[] pptx, DocToolkit.DocumentMetadata metadata)
        => DocToolkit.PresentationEditor.WithMetadata(pptx, metadata);

    public IReadOnlyList<string> ReadSmartArt(byte[] pptx, int index)
        => DocToolkit.PresentationEditor.ReadSmartArt(pptx, index);

    public Task<IReadOnlyList<string>> ReadSmartArtAsync(Stream source, int index, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.ReadSmartArtAsync(source, index, ct);

    public byte[] AddChart(
        byte[] pptx, int slideIndex, DocToolkit.ChartType type, DocToolkit.ChartData data, string title = "",
        double leftPoints = 0, double topPoints = 0, double widthPoints = 432, double heightPoints = 252)
        => DocToolkit.PresentationEditor.AddChart(pptx, slideIndex, type, data, title, leftPoints, topPoints, widthPoints, heightPoints);

    public Task<byte[]> AddChartAsync(
        Stream source, int slideIndex, DocToolkit.ChartType type, DocToolkit.ChartData data, string title = "",
        double leftPoints = 0, double topPoints = 0, double widthPoints = 432, double heightPoints = 252,
        CancellationToken ct = default)
        => DocToolkit.PresentationEditor.AddChartAsync(source, slideIndex, type, data, title, leftPoints, topPoints, widthPoints, heightPoints, ct);

    public Task<bool> IsProtectedAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.IsProtectedAsync(source, ct);

    public Task<DocToolkit.DocumentMetadata> ReadMetadataAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.ReadMetadataAsync(source, ct);

    public Task WithMetadataAsync(Stream source, DocToolkit.DocumentMetadata metadata, Stream destination, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.WithMetadataAsync(source, metadata, destination, ct);

    public string ReadNotes(byte[] pptx, int slideIndex)
        => DocToolkit.PresentationEditor.ReadNotes(pptx, slideIndex);

    public Task<string> ReadNotesAsync(Stream source, int slideIndex, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.ReadNotesAsync(source, slideIndex, ct);

    public byte[] SetNotes(byte[] pptx, int slideIndex, string notes)
        => DocToolkit.PresentationEditor.SetNotes(pptx, slideIndex, notes);

    public Task SetNotesAsync(Stream source, int slideIndex, string notes, Stream destination, CancellationToken ct = default)
        => DocToolkit.PresentationEditor.SetNotesAsync(source, slideIndex, notes, destination, ct);
}
