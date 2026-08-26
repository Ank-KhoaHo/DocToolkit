namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Lists what a document CONTAINS that <see cref="IDocxToPdfConverter"/> may not represent, so a
/// caller converting third-party documents knows which ones need a human to look at them. Registered
/// by <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.
/// </summary>
/// <remarks>
/// <b>An inventory of the input, never a claim about what was lost.</b> Nothing here converts
/// anything, so nothing here can know what the renderer did — it answers "is there anything in this
/// file worth a second look?", which stays true whatever a future renderer improves.
/// </remarks>
public interface IDocxToPdfPreflight
{
    /// <summary>
    /// Lists the constructs in <paramref name="docx"/> that the PDF renderer may not represent.
    /// </summary>
    /// <param name="docx">The document to inspect.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">It could not be opened or read.</exception>
    DocToolkit.DocxToPdfPreflightReport Inspect(byte[] docx);

    /// <inheritdoc cref="Inspect(byte[])"/>
    /// <param name="source">The document to inspect. Read to its end; never disposed or sought.</param>
    /// <param name="ct">Cancels before the document is read.</param>
    Task<DocToolkit.DocxToPdfPreflightReport> InspectAsync(
        Stream source, CancellationToken ct = default);
}
