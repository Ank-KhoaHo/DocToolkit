using OfficeIMOFormKey = OfficeIMO.Word.WordContentControlFormKey;
using OfficeIMOIssue = OfficeIMO.Word.WordContentControlFormIssue;
using OfficeIMOIssueKind = OfficeIMO.Word.WordContentControlFormIssueKind;
using OfficeIMOWordDocument = OfficeIMO.Word.WordDocument;

namespace DocToolkit;

/// <summary>
/// Reads, checks and fills the content controls a Word document carries — the format's own answer to
/// a fill-in form.
/// </summary>
/// <remarks>
/// <b>This library now has three template models, and a caller has whichever one their document was
/// authored with.</b> They are not interchangeable and none replaces the others:
///
/// <list type="table">
/// <listheader><term>model</term><description>marker · who writes it</description></listheader>
/// <item><term><see cref="DocxEditor"/></term><description>
/// <c>{{placeholder}}</c> — plain text, typed by anyone. A convention this library invented, and
/// breakable by editing inside the token.</description></item>
/// <item><term><see cref="DocxMailMerge"/></term><description>
/// <c>MERGEFIELD</c> — a real Word field, from <i>Insert → Merge Field</i>.</description></item>
/// <item><term>this class</term><description>
/// a content control — a named region Word itself protects, which is what makes it sturdier than a
/// placeholder.</description></item>
/// </list>
///
/// <b><see cref="Validate(byte[], IReadOnlyDictionary{string, DocxFormValue}, DocxFormKey)"/> checks
/// keys and, for a typed control, values.</b> A drop-down value outside its list, a non-date for a
/// date picker and a non-boolean for a check box are each reported under their own kind. A plain
/// text control validates anything, because there is no constraint to check it against.
///
/// <b><see cref="Fill(byte[], IReadOnlyDictionary{string, DocxFormValue}, DocxFormKey)"/> is
/// lenient</b>, where <see cref="DocxMailMerge.Merge"/> refuses. Two measured differences justify
/// that: a control given no value keeps its <i>own</i> existing text rather than showing an injected
/// marker, and unlike mail merge this class ships a <c>Validate</c> a caller can run first.
///
/// <b>Images are supplied as bytes and never as a path</b> — see <see cref="DocxFormValue"/>.
/// </remarks>
public static class DocxForm
{
    /// <summary>Reads every content control in <paramref name="docx"/>, and what it holds.</summary>
    /// <remarks>
    /// A document with no content controls reports none rather than failing — which is how a caller
    /// catches having passed the wrong document, since filling one succeeds and changes nothing.
    /// </remarks>
    /// <param name="docx">The document to read.</param>
    /// <param name="key">Which name identifies a control.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The document could not be opened or read.</exception>
    public static DocxFormReport Inspect(byte[] docx, DocxFormKey key = DocxFormKey.TagThenAlias)
    {
        RequireContent(docx);

        using var source = new MemoryStream(docx, writable: false);
        return InspectCore(source, key);
    }

    /// <summary>
    /// Reads every content control in the document in <paramref name="source"/>.
    /// <paramref name="source"/> is <b>read</b> to its end and is neither disposed, closed nor
    /// sought.
    /// </summary>
    /// <inheritdoc cref="Inspect(byte[], DocxFormKey)" path="/remarks"/>
    /// <param name="source">The document to read.</param>
    /// <param name="key">Which name identifies a control.</param>
    /// <param name="ct">Cancels before the document is read.</param>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be opened or read.</exception>
    public static async Task<DocxFormReport> InspectAsync(
        Stream source, DocxFormKey key = DocxFormKey.TagThenAlias, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, EmptySource, nameof(source), InspectFailure, ct)
            .ConfigureAwait(false);

        return InspectCore(docx, key);
    }

    /// <summary>
    /// Checks <paramref name="values"/> against the controls in <paramref name="docx"/>, without
    /// writing anything.
    /// </summary>
    /// <remarks>
    /// Run this before
    /// <see cref="Fill(byte[], IReadOnlyDictionary{string, DocxFormValue}, DocxFormKey)"/>, which is
    /// lenient by design and will not tell you what it skipped.
    ///
    /// <b>Every issue is reported</b> and <see cref="DocxFormValidation.IsValid"/> means "none of any
    /// kind". Filter <see cref="DocxFormValidation.Issues"/> by <see cref="DocxFormIssue.Kind"/> if
    /// you do not care about one of them.
    /// </remarks>
    /// <param name="docx">The document to check against.</param>
    /// <param name="values">The values to check.</param>
    /// <param name="key">Which name identifies a control.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> or <paramref name="values"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a value is null.</exception>
    /// <exception cref="DocumentConversionException">The document could not be opened or read.</exception>
    public static DocxFormValidation Validate(
        byte[] docx, IReadOnlyDictionary<string, DocxFormValue> values,
        DocxFormKey key = DocxFormKey.TagThenAlias)
    {
        RequireContent(docx);
        RequireValues(values);

        using var source = new MemoryStream(docx, writable: false);
        return ValidateCore(source, values, key);
    }

    /// <summary>
    /// Checks <paramref name="values"/> against the document in <paramref name="source"/>.
    /// <paramref name="source"/> is <b>read</b> to its end and is neither disposed, closed nor
    /// sought.
    /// </summary>
    /// <inheritdoc cref="Validate(byte[], IReadOnlyDictionary{string, DocxFormValue}, DocxFormKey)" path="/remarks"/>
    /// <param name="source">The document to check against.</param>
    /// <param name="values">The values to check.</param>
    /// <param name="key">Which name identifies a control.</param>
    /// <param name="ct">Cancels before the document is read.</param>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> or <paramref name="values"/> is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, <paramref name="source"/> held no bytes, or a value is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be opened or read.</exception>
    public static async Task<DocxFormValidation> ValidateAsync(
        Stream source, IReadOnlyDictionary<string, DocxFormValue> values,
        DocxFormKey key = DocxFormKey.TagThenAlias, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        RequireValues(values);
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, EmptySource, nameof(source), ValidateFailure, ct)
            .ConfigureAwait(false);

        return ValidateCore(docx, values, key);
    }

    /// <summary>
    /// A copy of <paramref name="docx"/> with each named control set to its value.
    /// </summary>
    /// <remarks>
    /// <b>Lenient.</b> A control with no entry in <paramref name="values"/> is left exactly as it
    /// was — measured: it keeps its own text rather than showing a marker, so a partial fill is a
    /// supported workflow. Call
    /// <see cref="Validate(byte[], IReadOnlyDictionary{string, DocxFormValue}, DocxFormKey)"/> first
    /// if you need to know what will be skipped.
    ///
    /// This returns bytes rather than a count of what it wrote. The count answers <i>"did my data
    /// match this template?"</i>, which is <c>Validate</c>'s question and is better asked before a
    /// document exists than inferred from one.
    /// </remarks>
    /// <param name="docx">The document to fill.</param>
    /// <param name="values">The value for each control.</param>
    /// <param name="key">Which name identifies a control.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> or <paramref name="values"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a value is null.</exception>
    /// <exception cref="DocumentConversionException">The document could not be read or written.</exception>
    public static byte[] Fill(
        byte[] docx, IReadOnlyDictionary<string, DocxFormValue> values,
        DocxFormKey key = DocxFormKey.TagThenAlias)
    {
        RequireContent(docx);
        RequireValues(values);

        using var source = new MemoryStream(docx, writable: false);
        using var result = FillCore(source, values, key);
        return result.ToArray();
    }

    /// <summary>
    /// Writes a copy of the document in <paramref name="source"/> to <paramref name="destination"/>
    /// with each named control set to its value. <paramref name="source"/> is <b>read</b> to its end
    /// and <paramref name="destination"/> is <b>written</b>; neither is disposed, closed nor sought.
    /// </summary>
    /// <inheritdoc cref="Fill(byte[], IReadOnlyDictionary{string, DocxFormValue}, DocxFormKey)" path="/remarks"/>
    /// <param name="source">The document to fill.</param>
    /// <param name="destination">Receives the filled document.</param>
    /// <param name="values">The value for each control.</param>
    /// <param name="key">Which name identifies a control.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    /// <exception cref="ArgumentNullException">A stream or <paramref name="values"/> is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, <paramref name="source"/> held no bytes, or a value is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be read or written.</exception>
    public static async Task FillAsync(
        Stream source, Stream destination, IReadOnlyDictionary<string, DocxFormValue> values,
        DocxFormKey key = DocxFormKey.TagThenAlias, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        RequireValues(values);
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, EmptySource, nameof(source), FillFailure, ct)
            .ConfigureAwait(false);

        using var result = FillCore(docx, values, key);
        await StreamPipeline.EmitAsync(result, destination, FillFailure, ct).ConfigureAwait(false);
    }

    private const string EmptySource = "DOCX content was empty.";
    private const string InspectFailure = "Failed to read the document's content controls. See the inner exception for details.";
    private const string ValidateFailure = "Failed to check the document's content controls. See the inner exception for details.";
    private const string FillFailure = "Failed to fill the document's content controls. See the inner exception for details.";

    private static void RequireContent(byte[] docx)
    {
        ArgumentNullException.ThrowIfNull(docx);
        if (docx.Length == 0)
            throw new ArgumentException(EmptySource, nameof(docx));
    }

    /// <summary>
    /// Refuses a null VALUE, on the rule <see cref="DocxMailMerge"/> adopted after measuring that a
    /// null is written as an empty string and reported as a success — the one silent half-fill
    /// nothing downstream can catch.
    /// </summary>
    private static void RequireValues(IReadOnlyDictionary<string, DocxFormValue> values)
    {
        ArgumentNullException.ThrowIfNull(values);

        foreach (KeyValuePair<string, DocxFormValue> pair in values)
        {
            if (pair.Value is null)
            {
                throw new ArgumentException(
                    $"The value for '{pair.Key}' is null. Pass a DocxFormValue - "
                    + "DocxFormValue.FromText(string.Empty) to mean \"leave it blank\".",
                    nameof(values));
            }
        }
    }

    private static DocxFormReport InspectCore(Stream source, DocxFormKey key)
    {
        try
        {
            using var document = OfficeIMOWordDocument.Load(source);

            DocxFormField[] fields = [.. document
                .ExtractContentControlValues(Upstream(key))
                .Select(pair => new DocxFormField(pair.Key, DocxFormValue.FromUpstream(pair.Value)))];

            return new DocxFormReport(fields);
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException(InspectFailure, ex);
        }
    }

    private static DocxFormValidation ValidateCore(
        Stream source, IReadOnlyDictionary<string, DocxFormValue> values, DocxFormKey key)
    {
        try
        {
            using var document = OfficeIMOWordDocument.Load(source);

            // Both flags at their strictest, deliberately. They suppress whole issue kinds, so
            // exposing them would make IsValid mean different things on different calls; reporting
            // everything and letting the caller filter is strictly more information.
            var result = document.ValidateContentControlValues(
                Upstream(values), Upstream(key), requireAllControls: true, allowUnusedValues: false);

            return new DocxFormValidation(
                result.IsValid,
                [.. result.ExpectedKeys],
                [.. result.SuppliedKeys],
                [.. result.Issues.Select(Issue)]);
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException(ValidateFailure, ex);
        }
    }

    /// <summary>
    /// The one implementation behind both filling overloads, so a <c>byte[]</c> call and a
    /// <c>Stream</c> call can never drift apart.
    /// </summary>
    /// <remarks>
    /// <b>The catch is unfiltered, and it disposes.</b> This method hands its buffer to its caller,
    /// so it owns that buffer until it returns — and a <i>filtered</i> catch that does not match
    /// never runs its body, which is exactly how a <see cref="MemoryStream"/> once escaped this
    /// repository with an exception.
    /// </remarks>
    private static MemoryStream FillCore(
        Stream source, IReadOnlyDictionary<string, DocxFormValue> values, DocxFormKey key)
    {
        var result = new MemoryStream();
        try
        {
            using var document = OfficeIMOWordDocument.Load(source);
            document.FillContentControlValues(Upstream(values), Upstream(key));
            document.Save(result);
            result.Position = 0;
            return result;
        }
        catch (Exception ex)
        {
            result.Dispose();
            throw new DocumentConversionException(FillFailure, ex);
        }
    }

    private static Dictionary<string, object?> Upstream(
        IReadOnlyDictionary<string, DocxFormValue> values)
    {
        // EVERY key is carried through, including one whose value maps to null. Dropping those
        // silently made Validate report MissingValue for a key the caller HAD supplied, and made
        // SuppliedKeys disagree with what was passed - a value read back from Inspect for an unset
        // date picker or an unselected drop-down maps to null, so the advertised round trip hit it.
        var copy = new Dictionary<string, object?>(values.Count);
        foreach (KeyValuePair<string, DocxFormValue> pair in values)
            copy[pair.Key] = pair.Value.ToUpstream();
        return copy;
    }

    private static OfficeIMOFormKey Upstream(DocxFormKey key) => key switch
    {
        DocxFormKey.Tag => OfficeIMOFormKey.Tag,
        DocxFormKey.Alias => OfficeIMOFormKey.Alias,
        DocxFormKey.AliasThenTag => OfficeIMOFormKey.AliasThenTag,
        _ => OfficeIMOFormKey.TagThenAlias,
    };

    private static DocxFormIssue Issue(OfficeIMOIssue issue)
        => new(issue.Key ?? string.Empty, Kind(issue.Kind), issue.Message ?? string.Empty);

    /// <summary>
    /// All nine kinds, one for one.
    /// </summary>
    /// <remarks>
    /// An earlier version mapped six of them to <see cref="DocxFormIssueKind.Other"/> on a
    /// measurement that turned out to be an artefact of its own fixtures — see
    /// <see cref="DocxFormIssueKind"/>. The <c>_</c> arm remains for a kind added upstream later,
    /// which would otherwise fail to compile a consumer's exhaustive switch.
    /// </remarks>
    private static DocxFormIssueKind Kind(OfficeIMOIssueKind kind) => kind switch
    {
        OfficeIMOIssueKind.MissingValue => DocxFormIssueKind.MissingValue,
        OfficeIMOIssueKind.UnusedValue => DocxFormIssueKind.UnusedValue,
        OfficeIMOIssueKind.DuplicateKey => DocxFormIssueKind.DuplicateKey,
        OfficeIMOIssueKind.UnmappedControl => DocxFormIssueKind.UnmappedControl,
        OfficeIMOIssueKind.InvalidBoolean => DocxFormIssueKind.InvalidBoolean,
        OfficeIMOIssueKind.InvalidDate => DocxFormIssueKind.InvalidDate,
        OfficeIMOIssueKind.InvalidChoice => DocxFormIssueKind.InvalidChoice,
        OfficeIMOIssueKind.InvalidImage => DocxFormIssueKind.InvalidImage,
        OfficeIMOIssueKind.InvalidRepeatingSection => DocxFormIssueKind.InvalidRepeatingSection,
        _ => DocxFormIssueKind.Other,
    };
}
