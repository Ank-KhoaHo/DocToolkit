using System.Runtime.CompilerServices;
using OfficeIMOMailMerge = OfficeIMO.Word.WordMailMerge;
using OfficeIMOFieldResult = OfficeIMO.Word.WordMailMergeFieldResult;
using OfficeIMOFieldStatus = OfficeIMO.Word.WordMailMergeFieldStatus;
using OfficeIMOIssue = OfficeIMO.Word.WordMailMergeTemplateIssue;
using OfficeIMOIssueKind = OfficeIMO.Word.WordMailMergeTemplateIssueKind;
using OfficeIMOWordDocument = OfficeIMO.Word.WordDocument;

namespace DocToolkit;

/// <summary>
/// Fills a Word mail-merge template — a document carrying <c>MERGEFIELD</c> instructions — from a
/// set of named values.
/// </summary>
/// <remarks>
/// <b>This is not <see cref="DocxEditor.FillRows(byte[], string, IEnumerable{IReadOnlyDictionary{string, string}})"/>
/// under another name, and the difference is who authored the template.</b>
///
/// <list type="table">
/// <listheader><term/><description>marker · authored by</description></listheader>
/// <item><term><c>DocxEditor</c></term><description>
/// <c>{{placeholder}}</c> — plain text, typed by anyone in any editor. A convention this library
/// invented.</description></item>
/// <item><term>this class</term><description>
/// a real Word field, produced by <i>Insert → Merge Field</i>, showing as <c>«FirstName»</c> with
/// field shading. This library reads what Word already writes.</description></item>
/// </list>
///
/// Neither substitutes for the other: a caller holding an existing Word mail-merge template has not
/// one <c>{{</c> in it, and a caller who does not own Word cannot author merge fields. Behind one
/// name, the same call would do nothing at all depending on how the template happened to be
/// authored.
///
/// <b>Both on-disk field encodings are handled.</b> Word writes the <i>complex</i> form —
/// <c>fldChar</c> begin, <c>instrText</c>, separate, result, end — while most generators and
/// hand-built documents emit the <i>simple</i> <c>w:fldSimple</c>. Measured: both merge.
///
/// <b>Field names match case-insensitively</b>, so a template field <c>FirstName</c> is filled by a
/// key spelled <c>firstname</c>. Measured, and it is the engine's own matching rather than a
/// property of the dictionary handed in.
///
/// <b>A null value is refused rather than merged.</b> Measured: the engine treats <c>null</c> as an
/// empty string, writes nothing, and reports the field <i>merged</i> and the document
/// <i>complete</i> — so a database NULL becomes a letter reading "Your balance is " that nothing
/// flags. A caller who means "leave it blank" writes <c>string.Empty</c> and says so; the one who
/// did not decide gets told. An empty string is accepted and merges.
///
/// <b>Produced documents are flattened.</b> The merged fields become ordinary text rather than live
/// fields, so re-opening the result in Word cannot re-merge it and shows no field shading. Measured:
/// the text is identical either way, so this is invisible to anything that reads a document back —
/// which is why it is asserted structurally instead.
/// </remarks>
public static class DocxMailMerge
{
    /// <summary>
    /// Reads what <paramref name="docx"/> asks for, without merging anything.
    /// </summary>
    /// <remarks>
    /// Use this to learn a template's field names, and to tell a sound template apart from one
    /// whose fields are malformed. A document carrying no merge fields reports none and is valid —
    /// which is how a caller catches having passed the wrong document, since merging one succeeds,
    /// changes nothing, and reports itself complete.
    /// </remarks>
    /// <param name="docx">The template to read.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The document could not be opened or read.</exception>
    public static DocxMailMergeTemplate InspectTemplate(byte[] docx)
    {
        RequireContent(docx);

        using var source = new MemoryStream(docx, writable: false);
        return InspectCore(source);
    }

    /// <summary>
    /// Reads what the template in <paramref name="source"/> asks for, without merging anything.
    /// <paramref name="source"/> is <b>read</b> to its end and is neither disposed, closed nor
    /// sought.
    /// </summary>
    /// <inheritdoc cref="InspectTemplate(byte[])" path="/remarks"/>
    /// <param name="source">The template to read.</param>
    /// <param name="ct">Cancels before the document is read.</param>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be opened or read.</exception>
    public static async Task<DocxMailMergeTemplate> InspectTemplateAsync(
        Stream source, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, EmptySource, nameof(source), InspectFailure, ct)
            .ConfigureAwait(false);

        return InspectCore(docx);
    }

    /// <summary>
    /// A copy of <paramref name="docx"/> with every merge field filled from
    /// <paramref name="values"/>.
    /// </summary>
    /// <remarks>
    /// <b>This refuses to produce a document that still has an unfilled field.</b> That is the
    /// whole difference between this and
    /// <see cref="MergeWithReport(byte[], IReadOnlyDictionary{string, string})"/>, and it is
    /// deliberate: measured, an unfilled field survives as a live field and the document reads
    /// <c>Your balance is «Balance»</c> — valid, opening cleanly, and looking finished. Nothing
    /// about it says otherwise except a report, and a report only helps a caller who reads one.
    ///
    /// Values naming fields the template does not have are ignored. A <b>mistyped key</b> is still
    /// caught, because the field it should have filled then goes unfilled and this throws.
    /// </remarks>
    /// <param name="docx">The template to fill.</param>
    /// <param name="values">The value for each merge field, matched case-insensitively.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> or <paramref name="values"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a value is null.</exception>
    /// <exception cref="DocumentConversionException">
    /// A merge field received no value, or the document could not be read or written.
    /// </exception>
    public static byte[] Merge(byte[] docx, IReadOnlyDictionary<string, string> values)
    {
        RequireContent(docx);
        RequireValues(values);

        using var source = new MemoryStream(docx, writable: false);
        using var result = MergeCore(source, values, strict: true, out _);
        return result.ToArray();
    }

    /// <summary>
    /// Writes a copy of the template in <paramref name="source"/> to
    /// <paramref name="destination"/> with every merge field filled from <paramref name="values"/>.
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed nor sought.
    /// </summary>
    /// <inheritdoc cref="Merge(byte[], IReadOnlyDictionary{string, string})" path="/remarks"/>
    /// <param name="source">The template to fill.</param>
    /// <param name="destination">Receives the filled document.</param>
    /// <param name="values">The value for each merge field, matched case-insensitively.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    /// <exception cref="ArgumentNullException">A stream or <paramref name="values"/> is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, <paramref name="source"/> held no bytes, or a value is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// A merge field received no value, or the document could not be read or written.
    /// </exception>
    public static async Task MergeAsync(
        Stream source, Stream destination, IReadOnlyDictionary<string, string> values,
        CancellationToken ct = default)
    {
        await MergeToStreamAsync(source, destination, values, strict: true, ct)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// A copy of <paramref name="docx"/> with every merge field filled from
    /// <paramref name="values"/>, <b>together with what happened to each one</b>.
    /// </summary>
    /// <remarks>
    /// The lenient half of the pair. This always produces a document, complete or not, and the
    /// report is how a caller learns which is which — see
    /// <see cref="DocxMailMergeResult.Document"/>, which is worth reading before shipping one.
    /// </remarks>
    /// <param name="docx">The template to fill.</param>
    /// <param name="values">The value for each merge field, matched case-insensitively.</param>
    /// <exception cref="ArgumentNullException"><paramref name="docx"/> or <paramref name="values"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a value is null.</exception>
    /// <exception cref="DocumentConversionException">The document could not be read or written.</exception>
    public static DocxMailMergeResult MergeWithReport(
        byte[] docx, IReadOnlyDictionary<string, string> values)
    {
        RequireContent(docx);
        RequireValues(values);

        using var source = new MemoryStream(docx, writable: false);
        using var result = MergeCore(source, values, strict: false, out DocxMailMergeReport report);
        return new DocxMailMergeResult(result.ToArray(), report);
    }

    /// <summary>
    /// Writes a copy of the template in <paramref name="source"/> to
    /// <paramref name="destination"/> with every merge field filled, and returns what happened to
    /// each one. <paramref name="source"/> is <b>read</b> to its end and
    /// <paramref name="destination"/> is <b>written</b>; neither is disposed, closed nor sought.
    /// </summary>
    /// <remarks>
    /// <inheritdoc cref="MergeWithReport(byte[], IReadOnlyDictionary{string, string})" path="/remarks"/>
    ///
    /// This returns a <see cref="DocxMailMergeReport"/> rather than a
    /// <see cref="DocxMailMergeResult"/> because the document went to
    /// <paramref name="destination"/>. Handing it back a second time would buffer a whole document
    /// nobody asked for, which is the opposite of what a <see cref="Stream"/> overload is for.
    /// </remarks>
    /// <param name="source">The template to fill.</param>
    /// <param name="destination">Receives the filled document.</param>
    /// <param name="values">The value for each merge field, matched case-insensitively.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    /// <exception cref="ArgumentNullException">A stream or <paramref name="values"/> is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, <paramref name="source"/> held no bytes, or a value is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The document could not be read or written.</exception>
    public static async Task<DocxMailMergeReport> MergeWithReportAsync(
        Stream source, Stream destination, IReadOnlyDictionary<string, string> values,
        CancellationToken ct = default)
    {
        return await MergeToStreamAsync(source, destination, values, strict: false, ct)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Fills <paramref name="docx"/> once per entry in <paramref name="records"/>, yielding each
    /// filled document in order.
    /// </summary>
    /// <remarks>
    /// <b>Strict, the same way <see cref="Merge(byte[], IReadOnlyDictionary{string, string})"/> is
    /// strict — this refuses the moment a record is incomplete, mid-sequence.</b> Everything already
    /// yielded before that point is unaffected; nothing after it runs. A lenient form that never
    /// throws for an incomplete record — <c>MergeBatchWithReport</c> — follows in a later change.
    ///
    /// <b>This is lazy.</b> Memory stays proportional to one document in flight, not the whole
    /// batch — <paramref name="records"/> is walked one entry at a time as the caller enumerates the
    /// result, and each document's bytes are only held until the caller moves on to the next one.
    /// </remarks>
    /// <param name="docx">The template to fill, once per record.</param>
    /// <param name="records">
    /// One dictionary of values per output document, matched case-insensitively — see
    /// <see cref="Merge(byte[], IReadOnlyDictionary{string, string})"/> for the matching rules,
    /// which apply here unchanged. An empty sequence yields no documents.
    /// </param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="docx"/> or <paramref name="records"/> is null, or an individual record in
    /// <paramref name="records"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a record's value is null.</exception>
    /// <exception cref="DocumentConversionException">
    /// The package could not be read or written, or a record is missing a value for a field the
    /// template requires — the message names the record's position (0-based) and the missing
    /// field(s).
    /// </exception>
    public static IEnumerable<byte[]> MergeBatch(
        byte[] docx, IEnumerable<IReadOnlyDictionary<string, string>> records)
    {
        RequireContent(docx);
        ArgumentNullException.ThrowIfNull(records);

        return MergeBatchCore(docx, records, strict: true).Select(item => item.Document);
    }

    /// <inheritdoc cref="MergeBatch(byte[], IEnumerable{IReadOnlyDictionary{string, string}})"/>
    /// <param name="docx">The template to fill, once per record.</param>
    /// <param name="records">
    /// One dictionary of values per output document, matched case-insensitively — see
    /// <see cref="Merge(byte[], IReadOnlyDictionary{string, string})"/> for the matching rules,
    /// which apply here unchanged. An empty sequence yields no documents.
    /// </param>
    /// <param name="ct">Cancels before the next record's merge runs.</param>
    public static async IAsyncEnumerable<byte[]> MergeBatchAsync(
        byte[] docx, IEnumerable<IReadOnlyDictionary<string, string>> records,
        [EnumeratorCancellation] CancellationToken ct = default)
    {
        RequireContent(docx);
        ArgumentNullException.ThrowIfNull(records);

        foreach (var item in MergeBatchCore(docx, records, strict: true))
        {
            ct.ThrowIfCancellationRequested();
            yield return item.Document;
        }
    }

    private const string EmptySource = "DOCX content was empty.";
    private const string InspectFailure = "Failed to read the document's mail-merge template. See the inner exception for details.";
    private const string MergeFailure = "Failed to fill the document's merge fields. See the inner exception for details.";

    private static void RequireContent(byte[] docx)
    {
        ArgumentNullException.ThrowIfNull(docx);
        if (docx.Length == 0)
            throw new ArgumentException(EmptySource, nameof(docx));
    }

    /// <summary>
    /// Refuses a null VALUE, for the reason on the class: the engine merges it as an empty string
    /// and reports the document complete, so it is the one way to ship a half-finished letter that
    /// neither the strict overload nor the report can catch.
    /// </summary>
    private static void RequireValues(IReadOnlyDictionary<string, string> values)
    {
        ArgumentNullException.ThrowIfNull(values);

        foreach (KeyValuePair<string, string> pair in values)
        {
            if (pair.Value is null)
            {
                throw new ArgumentException(
                    $"The value for '{pair.Key}' is null. A null merges as an empty string and is "
                    + "reported complete, so it cannot be told apart from a value somebody chose. "
                    + "Pass string.Empty to mean \"leave it blank\".",
                    nameof(values));
            }
        }
    }

    private static async Task<DocxMailMergeReport> MergeToStreamAsync(
        Stream source, Stream destination, IReadOnlyDictionary<string, string> values, bool strict,
        CancellationToken ct)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        RequireValues(values);
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, EmptySource, nameof(source), MergeFailure, ct)
            .ConfigureAwait(false);

        using MemoryStream merged = MergeCore(docx, values, strict, out DocxMailMergeReport report);
        await StreamPipeline.EmitAsync(merged, destination, MergeFailure, ct).ConfigureAwait(false);
        return report;
    }

    /// <summary>
    /// The one implementation behind all four merging overloads, so a <c>byte[]</c> call and a
    /// <c>Stream</c> call can never drift apart.
    /// </summary>
    /// <remarks>
    /// <b>The catch is unfiltered, and it disposes</b>, for the reason recorded on
    /// <see cref="DocxReview"/>: this method hands its buffer to its caller, and a <i>filtered</i>
    /// catch that does not match never runs its body — which is exactly how a
    /// <see cref="MemoryStream"/> once escaped this repository with an exception.
    ///
    /// <b>The strict refusal is deliberately OUTSIDE that try.</b> Inside it, the unfiltered catch
    /// would swallow this method's own <see cref="DocumentConversionException"/> and re-wrap it in
    /// the generic read-or-write message, hiding the field names that make it actionable — the same
    /// hazard <c>DocxReview.ApplyCore</c> records from the other direction.
    ///
    /// <b>It refuses on the engine's own verdict</b>, <c>IsComplete</c>, which is the value
    /// <c>EnsureComplete</c> itself tests. Reading it rather than catching it is what lets the
    /// message name <i>every</i> unfilled field instead of only the first.
    /// </remarks>
    private static MemoryStream MergeCore(
        Stream source, IReadOnlyDictionary<string, string> values, bool strict,
        out DocxMailMergeReport report)
    {
        var result = new MemoryStream();
        DocxMailMergeReport merged;
        try
        {
            using var document = OfficeIMOWordDocument.Load(source);

            merged = Map(OfficeIMOMailMerge.ExecuteWithReport(
                document, Copy(values), removeFields: true));

            document.Save(result);
            result.Position = 0;
        }
        catch (Exception ex)
        {
            result.Dispose();
            throw new DocumentConversionException(MergeFailure, ex);
        }

        report = merged;
        if (strict && !merged.IsComplete)
        {
            result.Dispose();
            throw new DocumentConversionException(
                $"{merged.MissingFieldNames.Count} merge field(s) received no value: "
                + $"{string.Join(", ", merged.MissingFieldNames)}. The document would have been "
                + "produced with each of them still showing its placeholder text. Supply the "
                + "missing values, or call MergeWithReport to take the document as it is.");
        }

        return result;
    }

    /// <summary>
    /// The one loop behind every batch overload that produces documents in memory — strict and
    /// lenient alike — so they can never drift apart the same way <see cref="MergeCore"/> already
    /// keeps <see cref="Merge(byte[], IReadOnlyDictionary{string, string})"/> and
    /// <see cref="MergeWithReport(byte[], IReadOnlyDictionary{string, string})"/> from drifting.
    /// </summary>
    /// <remarks>
    /// Deliberately a `yield return` iterator, and deliberately PRIVATE: every public caller
    /// (<see cref="MergeBatch"/> today, and <c>MergeBatchWithReport</c> in a later change) validates
    /// its own arguments in an ordinary, non-iterator method body before calling this — a
    /// `yield return` in a PUBLIC method would defer that validation until the caller starts
    /// enumerating, which is not how every other guard in this class behaves.
    /// </remarks>
    private static IEnumerable<DocxMailMergeBatchItem> MergeBatchCore(
        byte[] docx, IEnumerable<IReadOnlyDictionary<string, string>> records, bool strict)
    {
        var index = 0;
        foreach (var record in records)
        {
            RequireValues(record);

            using var source = new MemoryStream(docx, writable: false);
            DocxMailMergeReport report;
            MemoryStream result;
            try
            {
                result = MergeCore(source, record, strict, out report);
            }
            catch (DocumentConversionException ex) when (strict)
            {
                throw new DocumentConversionException($"Record {index}: {ex.Message}", ex);
            }

            using (result)
            {
                yield return new DocxMailMergeBatchItem(index, result.ToArray(), report);
            }

            index++;
        }
    }

    private static DocxMailMergeTemplate InspectCore(Stream source)
    {
        try
        {
            using var document = OfficeIMOWordDocument.Load(source);

            // null, NOT an empty sequence. They mean different things here and the difference is a
            // trap: null asks the engine to DISCOVER the names, while an empty sequence asserts the
            // template has none - so a perfectly good template comes back IsValid=false with one
            // issue per field it does have. Measured. Passing null is why no caller of this API can
            // make that mistake.
            var inspection = OfficeIMOMailMerge.InspectTemplate(document, null!, null!, null!);

            return new DocxMailMergeTemplate(
                [.. inspection.MergeFieldNames],
                [.. inspection.Issues.Select(Issue)],
                inspection.IsValid);
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException(InspectFailure, ex);
        }
    }

    /// <summary>
    /// The engine takes an <c>IDictionary</c>; the public API takes an <c>IReadOnlyDictionary</c>,
    /// which is the shape a caller can hand over without their own copy being mutable from here.
    /// </summary>
    /// <remarks>
    /// The comparer is deliberately the default. Case-insensitive matching is the engine's own —
    /// measured by handing it an ordinary ordinal dictionary and watching <c>firstname</c> fill a
    /// field named <c>FirstName</c> — so imposing one here would claim a guarantee this code does
    /// not provide.
    /// </remarks>
    private static Dictionary<string, string> Copy(IReadOnlyDictionary<string, string> values)
    {
        var copy = new Dictionary<string, string>(values.Count);
        foreach (KeyValuePair<string, string> pair in values)
            copy[pair.Key] = pair.Value;
        return copy;
    }

    private static DocxMailMergeReport Map(OfficeIMO.Word.WordMailMergeExecutionReport native)
        => new(
            [.. native.Fields.Select(Field)],
            [.. native.MissingValueNames],
            native.MergedCount,
            native.IsComplete);

    private static DocxMailMergeField Field(OfficeIMOFieldResult field)
        => new(Text(field.Name), Status(field.Status), field.Value, Text(field.Message));

    private static DocxMailMergeIssue Issue(OfficeIMOIssue issue)
        => new(Text(issue.Name), Text(issue.Message), Kind(issue.Kind));

    private static DocxMailMergeFieldStatus Status(OfficeIMOFieldStatus status) => status switch
    {
        OfficeIMOFieldStatus.Merged => DocxMailMergeFieldStatus.Merged,
        OfficeIMOFieldStatus.MissingValue => DocxMailMergeFieldStatus.MissingValue,
        OfficeIMOFieldStatus.UnsupportedFormatting => DocxMailMergeFieldStatus.UnsupportedFormatting,
        _ => DocxMailMergeFieldStatus.Malformed,
    };

    /// <summary>
    /// Twelve kinds to three. See <see cref="DocxMailMergeIssueKind"/> for why the other nine
    /// arrive as <see cref="DocxMailMergeIssueKind.Other"/> rather than under their own names.
    /// </summary>
    private static DocxMailMergeIssueKind Kind(OfficeIMOIssueKind kind) => kind switch
    {
        OfficeIMOIssueKind.MalformedMergeField => DocxMailMergeIssueKind.MalformedField,
        OfficeIMOIssueKind.UnsupportedMergeFieldFormatting => DocxMailMergeIssueKind.UnsupportedFormatting,
        _ => DocxMailMergeIssueKind.Other,
    };

    private static string Text(string? value) => value ?? string.Empty;
}
