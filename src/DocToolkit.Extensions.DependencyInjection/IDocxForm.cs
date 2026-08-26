namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Reads, checks and fills the content controls a Word document carries — the format's own answer to
/// a fill-in form. Registered by <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.
/// </summary>
/// <remarks>
/// <b>One of three template models, not a replacement for the other two.</b> <see cref="IDocxEditor"/>
/// fills <c>{{placeholder}}</c> text and <see cref="IDocxMailMerge"/> fills <c>MERGEFIELD</c>
/// instructions; which you need is decided by whoever authored the document.
///
/// <b>Only the document BODY is read or written.</b> A control in a header or footer is invisible
/// here — <see cref="IDocxMailMerge"/> does reach headers.
/// </remarks>
public interface IDocxForm
{
    /// <summary>
    /// Reads the content controls in <paramref name="docx"/>'s body, and what they hold. <b>Not
    /// necessarily every control</b> — see <see cref="DocToolkit.DocxFormReport.Fields"/>.
    /// </summary>
    /// <param name="docx">The document to read.</param>
    /// <param name="key">Which name identifies a control.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be opened or read.</exception>
    DocToolkit.DocxFormReport Inspect(
        byte[] docx, DocToolkit.DocxFormKey key = DocToolkit.DocxFormKey.TagThenAlias);

    /// <inheritdoc cref="Inspect(byte[], DocToolkit.DocxFormKey)"/>
    /// <param name="source">The document to read. Read to its end; never disposed or sought.</param>
    /// <param name="key">Which name identifies a control.</param>
    /// <param name="ct">Cancels before the document is read.</param>
    Task<DocToolkit.DocxFormReport> InspectAsync(
        Stream source, DocToolkit.DocxFormKey key = DocToolkit.DocxFormKey.TagThenAlias,
        CancellationToken ct = default);

    /// <summary>
    /// Checks <paramref name="values"/> against the controls in <paramref name="docx"/> without
    /// writing anything. <b>Not a promise that <see cref="Fill(byte[],
    /// IReadOnlyDictionary{string, DocToolkit.DocxFormValue}, DocToolkit.DocxFormKey)"/> will
    /// succeed</b> — image bytes are not decoded here.
    /// </summary>
    /// <param name="docx">The document to check against.</param>
    /// <param name="values">The values to check.</param>
    /// <param name="key">Which name identifies a control.</param>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a value is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be opened or read.</exception>
    DocToolkit.DocxFormValidation Validate(
        byte[] docx, IReadOnlyDictionary<string, DocToolkit.DocxFormValue> values,
        DocToolkit.DocxFormKey key = DocToolkit.DocxFormKey.TagThenAlias);

    /// <inheritdoc cref="Validate(byte[], IReadOnlyDictionary{string, DocToolkit.DocxFormValue}, DocToolkit.DocxFormKey)"/>
    /// <param name="source">The document to check against. Read to its end; never disposed or sought.</param>
    /// <param name="values">The values to check.</param>
    /// <param name="key">Which name identifies a control.</param>
    /// <param name="ct">Cancels before the document is read.</param>
    Task<DocToolkit.DocxFormValidation> ValidateAsync(
        Stream source, IReadOnlyDictionary<string, DocToolkit.DocxFormValue> values,
        DocToolkit.DocxFormKey key = DocToolkit.DocxFormKey.TagThenAlias,
        CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="docx"/> with each named control set to its value.
    /// </summary>
    /// <remarks>
    /// <b>Lenient about a MISSING value</b> — a control with no entry keeps its own existing text.
    /// <b>Not lenient about a value that does not fit a typed control</b>, and the three typed kinds
    /// disagree: a drop-down value outside its list throws, while a bad date or boolean is silently
    /// skipped. Run <c>Validate</c> first; it reports all three the same way.
    /// </remarks>
    /// <param name="docx">The document to fill.</param>
    /// <param name="values">The value for each control.</param>
    /// <param name="key">Which name identifies a control.</param>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a value is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// It could not be read or written, or a value did not fit a typed control.
    /// </exception>
    byte[] Fill(
        byte[] docx, IReadOnlyDictionary<string, DocToolkit.DocxFormValue> values,
        DocToolkit.DocxFormKey key = DocToolkit.DocxFormKey.TagThenAlias);

    /// <inheritdoc cref="Fill(byte[], IReadOnlyDictionary{string, DocToolkit.DocxFormValue}, DocToolkit.DocxFormKey)"/>
    /// <param name="source">The document to fill. Read to its end; never disposed or sought.</param>
    /// <param name="destination">Receives the filled document. Written; never disposed or sought.</param>
    /// <param name="values">The value for each control.</param>
    /// <param name="key">Which name identifies a control.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    Task FillAsync(
        Stream source, Stream destination,
        IReadOnlyDictionary<string, DocToolkit.DocxFormValue> values,
        DocToolkit.DocxFormKey key = DocToolkit.DocxFormKey.TagThenAlias,
        CancellationToken ct = default);
}
