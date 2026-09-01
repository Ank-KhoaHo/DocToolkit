using Microsoft.Extensions.Options;

namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>Default <see cref="IDocxEditor"/>, delegating to <see cref="DocToolkit.DocxEditor"/>.</summary>
internal sealed class DocxEditorService : IDocxEditor
{
    private readonly IOptionsMonitor<DocToolkitOptions> _options;

    // The only option this service reads is Page, and it reads it PER CALL for the same reason
    // the converters do: these are singletons, so a captured IOptions<T>.Value would freeze
    // whatever configuration existed at startup for the life of the process.
    public DocxEditorService(IOptionsMonitor<DocToolkitOptions> options) => _options = options;

    // Create is a producer, so DocToolkitOptions.Page applies to it as much as to the HTML
    // converters. Leaving it out would make the option true of two producers out of three - the
    // kind of inconsistency a consumer discovers one document at a time.
    public byte[] Create(IEnumerable<DocToolkit.DocxBlock> blocks)
        => DocToolkit.DocxEditor.Create(blocks, _options.CurrentValue.Page);

    public byte[] ReplaceText(byte[] docx, IReadOnlyDictionary<string, string> replacements)
        => DocToolkit.DocxEditor.ReplaceText(docx, replacements);

    public byte[] FillRows(
        byte[] docx, string collection,
        IEnumerable<IReadOnlyDictionary<string, string>> rows)
        => DocToolkit.DocxEditor.FillRows(docx, collection, rows);

    public byte[] ReplaceImage(
        byte[] docx, string placeholder, byte[] image,
        double? widthPoints = null, double? heightPoints = null)
        => DocToolkit.DocxEditor.ReplaceImage(docx, placeholder, image, widthPoints, heightPoints);

    public string ExtractText(byte[] docx) => DocToolkit.DocxEditor.ExtractText(docx);

    public string ExtractText(byte[] docx, bool includeHeadersAndFooters)
        => DocToolkit.DocxEditor.ExtractText(docx, includeHeadersAndFooters);

    public Task ReplaceTextAsync(
        Stream source, IReadOnlyDictionary<string, string> replacements, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.DocxEditor.ReplaceTextAsync(source, replacements, destination, ct);

    public Task<string> ExtractTextAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.DocxEditor.ExtractTextAsync(source, ct);

    public Task<string> ExtractTextAsync(Stream source, bool includeHeadersAndFooters, CancellationToken ct = default)
        => DocToolkit.DocxEditor.ExtractTextAsync(source, includeHeadersAndFooters, ct);

    public Task CreateAsync(
        IEnumerable<DocToolkit.DocxBlock> blocks, Stream destination, CancellationToken ct = default)
        => DocToolkit.DocxEditor.CreateAsync(blocks, _options.CurrentValue.Page, destination, ct);

    public Task FillRowsAsync(
        Stream source, string collection,
        IEnumerable<IReadOnlyDictionary<string, string>> rows, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.DocxEditor.FillRowsAsync(source, collection, rows, destination, ct);

    public Task ReplaceImageAsync(
        Stream source, string placeholder, byte[] image, Stream destination,
        double? widthPoints = null, double? heightPoints = null,
        CancellationToken ct = default)
        => DocToolkit.DocxEditor.ReplaceImageAsync(
            source, placeholder, image, destination, widthPoints, heightPoints, ct);

    public byte[] Create(System.Collections.Generic.IEnumerable<DocToolkit.DocxBlock> blocks, DocToolkit.PageSetup page)
        => DocToolkit.DocxEditor.Create(blocks, page);

    public Task CreateAsync(System.Collections.Generic.IEnumerable<DocToolkit.DocxBlock> blocks, DocToolkit.PageSetup page, Stream destination, CancellationToken ct = default)
        => DocToolkit.DocxEditor.CreateAsync(blocks, page, destination, ct);

    public int TableCount(byte[] docx) => DocToolkit.DocxEditor.TableCount(docx);

    public Task<int> TableCountAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.DocxEditor.TableCountAsync(source, ct);

    public IReadOnlyList<IReadOnlyList<string>> ReadTable(byte[] docx, int index)
        => DocToolkit.DocxEditor.ReadTable(docx, index);

    public Task<IReadOnlyList<IReadOnlyList<string>>> ReadTableAsync(Stream source, int index, CancellationToken ct = default)
        => DocToolkit.DocxEditor.ReadTableAsync(source, index, ct);

    public byte[] Protect(byte[] docx, string password)
        => DocToolkit.DocxEditor.Protect(docx, password);

    public byte[] Unprotect(byte[] docx, string password)
        => DocToolkit.DocxEditor.Unprotect(docx, password);

    public bool IsProtected(byte[] docx) => DocToolkit.DocxEditor.IsProtected(docx);

    public Task ProtectAsync(Stream source, Stream destination, string password, CancellationToken ct = default)
        => DocToolkit.DocxEditor.ProtectAsync(source, destination, password, ct);

    public Task UnprotectAsync(Stream source, Stream destination, string password, CancellationToken ct = default)
        => DocToolkit.DocxEditor.UnprotectAsync(source, destination, password, ct);

    public byte[] AddFootnote(byte[] docx, string placeholder, string footnoteText)
        => DocToolkit.DocxEditor.AddFootnote(docx, placeholder, footnoteText);

    public Task AddFootnoteAsync(
        Stream source, string placeholder, string footnoteText, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.DocxEditor.AddFootnoteAsync(source, placeholder, footnoteText, destination, ct);

    public byte[] AddEndnote(byte[] docx, string placeholder, string endnoteText)
        => DocToolkit.DocxEditor.AddEndnote(docx, placeholder, endnoteText);

    public Task AddEndnoteAsync(
        Stream source, string placeholder, string endnoteText, Stream destination,
        CancellationToken ct = default)
        => DocToolkit.DocxEditor.AddEndnoteAsync(source, placeholder, endnoteText, destination, ct);

    public byte[] AddTableOfContents(byte[] docx, string placeholder, int minLevel = 1, int maxLevel = 3)
        => DocToolkit.DocxEditor.AddTableOfContents(docx, placeholder, minLevel, maxLevel);

    public Task AddTableOfContentsAsync(
        Stream source, string placeholder, Stream destination,
        int minLevel = 1, int maxLevel = 3, CancellationToken ct = default)
        => DocToolkit.DocxEditor.AddTableOfContentsAsync(source, placeholder, destination, minLevel, maxLevel, ct);

    public DocToolkit.DocumentSignatureInfo InspectSignatures(byte[] docx)
        => DocToolkit.DocxEditor.InspectSignatures(docx);

    public Task<DocToolkit.DocumentSignatureInfo> InspectSignaturesAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.DocxEditor.InspectSignaturesAsync(source, ct);

    public DocToolkit.DocumentSignatureValidationReport ValidateSignatures(byte[] docx, DocToolkit.DocumentSignatureValidationOptions? options = null)
        => DocToolkit.DocxEditor.ValidateSignatures(docx, options);

    public Task<DocToolkit.DocumentSignatureValidationReport> ValidateSignaturesAsync(
        Stream source, DocToolkit.DocumentSignatureValidationOptions? options = null, CancellationToken ct = default)
        => DocToolkit.DocxEditor.ValidateSignaturesAsync(source, options, ct);

    public DocToolkit.DocumentMetadata ReadMetadata(byte[] docx)
        => DocToolkit.DocxEditor.ReadMetadata(docx);

    public byte[] WithMetadata(byte[] docx, DocToolkit.DocumentMetadata metadata)
        => DocToolkit.DocxEditor.WithMetadata(docx, metadata);

    public Task<bool> IsProtectedAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.DocxEditor.IsProtectedAsync(source, ct);

    public Task<DocToolkit.DocumentMetadata> ReadMetadataAsync(Stream source, CancellationToken ct = default)
        => DocToolkit.DocxEditor.ReadMetadataAsync(source, ct);

    public Task WithMetadataAsync(Stream source, DocToolkit.DocumentMetadata metadata, Stream destination, CancellationToken ct = default)
        => DocToolkit.DocxEditor.WithMetadataAsync(source, metadata, destination, ct);
}
