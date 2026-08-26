namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Fills a Word mail-merge template — a document carrying <c>MERGEFIELD</c> instructions — from a set
/// of named values. Registered by <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.
/// </summary>
/// <remarks>
/// <b>Not <see cref="IDocxEditor"/>'s placeholders and not <see cref="IDocxForm"/>'s content
/// controls.</b> The difference is who authored the template: <c>{{placeholder}}</c> is a convention
/// this library invented, a <c>MERGEFIELD</c> is what Word writes from <i>Insert → Merge Field</i>,
/// and a content control is a named region Word protects. A caller has whichever one their document
/// was built with.
/// </remarks>
public interface IDocxMailMerge
{
    /// <summary>Reads what <paramref name="docx"/> asks for, without merging anything.</summary>
    /// <param name="docx">The template to read.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be opened or read.</exception>
    DocToolkit.DocxMailMergeTemplate InspectTemplate(byte[] docx);

    /// <inheritdoc cref="InspectTemplate(byte[])" path="/summary|/remarks|/exception"/>
    /// <param name="source">The template to read. Read to its end; never disposed or sought.</param>
    /// <param name="ct">Cancels before the document is read.</param>
    Task<DocToolkit.DocxMailMergeTemplate> InspectTemplateAsync(
        Stream source, CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="docx"/> with every merge field filled.
    /// </summary>
    /// <remarks>
    /// <b>Refuses to produce a document with an unfilled field</b>, naming every one. Measured: an
    /// unfilled field survives as a live field and the document reads <c>«Balance»</c> — valid,
    /// opening cleanly, and looking finished. Use
    /// <see cref="MergeWithReport(byte[], IReadOnlyDictionary{string, string})"/> when you want it
    /// anyway.
    /// </remarks>
    /// <param name="docx">The template to fill.</param>
    /// <param name="values">The value for each field, matched case-insensitively.</param>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a value is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">
    /// A field received no value, or the document could not be read or written.
    /// </exception>
    byte[] Merge(byte[] docx, IReadOnlyDictionary<string, string> values);

    /// <inheritdoc cref="Merge(byte[], IReadOnlyDictionary{string, string})" path="/summary|/remarks|/exception"/>
    /// <param name="source">The template to fill. Read to its end; never disposed or sought.</param>
    /// <param name="destination">Receives the filled document. Written; never disposed or sought.</param>
    /// <param name="values">The value for each field, matched case-insensitively.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    Task MergeAsync(
        Stream source, Stream destination, IReadOnlyDictionary<string, string> values,
        CancellationToken ct = default);

    /// <summary>
    /// A copy of <paramref name="docx"/> with every merge field filled, <b>together with what
    /// happened to each one</b>. Always produces a document, complete or not.
    /// </summary>
    /// <param name="docx">The template to fill.</param>
    /// <param name="values">The value for each field, matched case-insensitively.</param>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a value is null.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be read or written.</exception>
    DocToolkit.DocxMailMergeResult MergeWithReport(
        byte[] docx, IReadOnlyDictionary<string, string> values);

    /// <inheritdoc cref="MergeWithReport(byte[], IReadOnlyDictionary{string, string})" path="/summary|/remarks|/exception"/>
    /// <remarks>
    /// Returns the report alone, because the document went to <paramref name="destination"/>.
    /// </remarks>
    /// <param name="source">The template to fill. Read to its end; never disposed or sought.</param>
    /// <param name="destination">Receives the filled document. Written; never disposed or sought.</param>
    /// <param name="values">The value for each field, matched case-insensitively.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    Task<DocToolkit.DocxMailMergeReport> MergeWithReportAsync(
        Stream source, Stream destination, IReadOnlyDictionary<string, string> values,
        CancellationToken ct = default);
}
