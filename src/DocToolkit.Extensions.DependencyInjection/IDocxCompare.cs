namespace DocToolkit.Extensions.DependencyInjection;

/// <summary>
/// Compares two versions of a Word document and returns the later one with the differences marked
/// as tracked changes. Registered by
/// <see cref="ServiceCollectionExtensions.AddDocToolkit"/>.
/// </summary>
/// <remarks>
/// <b>The result is an ordinary .docx carrying revisions</b>, not a report — so
/// <see cref="IDocxReview"/> reads and resolves it without knowing where it came from.
/// <c>Inspect</c> lists the revisions on the report it returns, and <c>AcceptRevisions</c> /
/// <c>RejectRevisions</c> resolve them.
///
/// <para>
/// <b>Three limits, repeated here because a DI consumer reads them at the call site and nowhere
/// else.</b> Only <b>paragraph text</b> is compared; <b>tables are reported rather than diffed</b>;
/// and a <b>formatting-only change is never detected at all</b>. A comparison that quietly
/// mis-marked a table would be worse than one that says what it did not look at.
/// </para>
///
/// <para>
/// <b>Separate from <see cref="IDocxEditor"/> deliberately</b>, on the same grounds as
/// <see cref="IDocxReview"/>: that interface is about one document's content, and this is about the
/// difference between two of them.
/// </para>
///
/// <para>
/// <b>There are no <c>Stream</c> members here, and that is not an omission to fill in later.</b> A
/// comparison reads <b>two</b> documents, while the <c>Stream</c> shape this library uses
/// throughout is <c>source</c>/<c>destination</c>. Two sources and one destination is a shape
/// nothing else has, and inventing it in the DI layer rather than in the core package would invert
/// the direction every other member follows — core grows the overload, the mirror follows one
/// release later. If those overloads are wanted, they belong in core first.
/// </para>
/// </remarks>
public interface IDocxCompare
{
    /// <summary>
    /// Returns <paramref name="revised"/> with its differences from <paramref name="original"/>
    /// marked as tracked insertions and deletions.
    /// </summary>
    /// <remarks>
    /// <b>This overload discards the report of what was NOT compared.</b> Prefer
    /// <see cref="CompareWithReport(byte[], byte[], string)"/> unless you already know the
    /// documents contain no tables and you do not care about formatting — a caller who cannot see
    /// what the comparison skipped has a verdict covering less than they think.
    ///
    /// <b>Comparing a document with itself produces no revisions</b>, rather than a document marked
    /// entirely rewritten.
    /// </remarks>
    /// <param name="original">The earlier version.</param>
    /// <param name="revised">The later version, which the result is built from.</param>
    /// <param name="author">The name recorded against each revision.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">Either document is empty, or <paramref name="author"/> is blank.</exception>
    /// <exception cref="DocToolkit.DocumentConversionException">Either package could not be opened or edited.</exception>
    byte[] Compare(byte[] original, byte[] revised, string author);

    /// <summary>
    /// Compares two documents and reports what the comparison did not look at.
    /// </summary>
    /// <inheritdoc cref="Compare(byte[], byte[], string)" path="/param|/exception"/>
    /// <returns>
    /// The marked-up document, with a warning for every construct present but not compared.
    /// <c>HasLoss</c> is true whenever anything was skipped, which is the signal that the verdict
    /// covers less than the document.
    /// </returns>
    DocToolkit.ConversionResult<byte[]> CompareWithReport(byte[] original, byte[] revised, string author);
}
