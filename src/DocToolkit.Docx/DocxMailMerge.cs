using System.Runtime.CompilerServices;
using OfficeIMOMailMerge = OfficeIMO.Word.WordMailMerge;
using OfficeIMOBlockData = OfficeIMO.Word.WordMailMergeBlockData;
using OfficeIMOFieldResult = OfficeIMO.Word.WordMailMergeFieldResult;
using OfficeIMOFieldStatus = OfficeIMO.Word.WordMailMergeFieldStatus;
using OfficeIMOIssue = OfficeIMO.Word.WordMailMergeTemplateIssue;
using OfficeIMOIssueKind = OfficeIMO.Word.WordMailMergeTemplateIssueKind;
using OfficeIMOWordDocument = OfficeIMO.Word.WordDocument;
using OfficeIMOTableRowGroup = OfficeIMO.Word.WordMailMergeTableRowGroup;

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
/// did not decide gets told. An empty string is accepted and merges. <b>This holds for every method
/// here that takes values</b>, including the per-record and per-row collections the repeating and
/// table-row methods take, where the refusal names which record carried the null.
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
        RequireValues(values, nameof(values));

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
        RequireValues(values, nameof(values));

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
    /// A copy of <paramref name="docx"/> with every conditional block (<c>{{#Name}}</c> …
    /// <c>{{/Name}}</c>) resolved — included and its markers removed when its condition is
    /// <see langword="true"/>, removed entirely (markers and content) when
    /// <see langword="false"/>.
    /// </summary>
    /// <remarks>
    /// <b>Refuses to produce a document with a condition the template asks for but
    /// <paramref name="conditions"/> did not supply</b> — measured: the underlying engine throws
    /// immediately for an unsupplied name, so this preflights via <see cref="InspectTemplate(byte[])"/>
    /// and refuses before the document is ever touched, naming every missing condition at once. Use
    /// <see cref="MergeConditionalWithReport(byte[], IReadOnlyDictionary{string, bool})"/> when you
    /// want the document anyway.
    ///
    /// <b>Run this before <see cref="Merge(byte[], IReadOnlyDictionary{string, string})"/></b> when
    /// a template mixes conditional blocks with ordinary merge fields — a field inside a block that
    /// ends up excluded is removed along with the block, so running the field-level merge first
    /// would fill a field that this call is about to delete.
    ///
    /// A marker paragraph must contain <b>only</b> the marker — trailing text on the same paragraph
    /// is not recognised as a marker at all, which is left as literal text in the output. This is
    /// inherent to the marker convention and cannot be detected here.
    /// </remarks>
    /// <param name="docx">The template to resolve.</param>
    /// <param name="conditions">Whether to include each named block.</param>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">
    /// A condition the template asks for was not supplied, the marker structure is unbalanced, or
    /// the document could not be read or written.
    /// </exception>
    public static byte[] MergeConditional(byte[] docx, IReadOnlyDictionary<string, bool> conditions)
    {
        RequireContent(docx);
        ArgumentNullException.ThrowIfNull(conditions);

        using var source = new MemoryStream(docx, writable: false);
        using var result = MergeConditionalCore(source, conditions, strict: true, out _);
        return result.ToArray();
    }

    /// <summary>
    /// Writes a copy of the template in <paramref name="source"/> to <paramref name="destination"/>
    /// with every conditional block resolved. <paramref name="source"/> is <b>read</b> to its end
    /// and <paramref name="destination"/> is <b>written</b>; neither is disposed, closed nor sought.
    /// </summary>
    /// <inheritdoc cref="MergeConditional(byte[], IReadOnlyDictionary{string, bool})" path="/remarks"/>
    /// <param name="source">The template to resolve.</param>
    /// <param name="destination">Receives the resolved document.</param>
    /// <param name="conditions">Whether to include each named block.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    /// <exception cref="ArgumentNullException">A stream or <paramref name="conditions"/> is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, or <paramref name="source"/> held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// A condition the template asks for was not supplied, the marker structure is unbalanced, or
    /// the document could not be read or written.
    /// </exception>
    public static async Task MergeConditionalAsync(
        Stream source, Stream destination, IReadOnlyDictionary<string, bool> conditions,
        CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ArgumentNullException.ThrowIfNull(conditions);
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, EmptySource, nameof(source), ConditionalFailure, ct)
            .ConfigureAwait(false);

        using MemoryStream merged = MergeConditionalCore(docx, conditions, strict: true, out _);
        await StreamPipeline.EmitAsync(merged, destination, ConditionalFailure, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// A copy of <paramref name="docx"/> with every conditional block resolved, <b>together with
    /// which condition names the template asked for that <paramref name="conditions"/> did not
    /// supply</b>. Always produces a document, except when the marker structure is genuinely
    /// unbalanced.
    /// </summary>
    /// <remarks>
    /// <inheritdoc cref="MergeConditional(byte[], IReadOnlyDictionary{string, bool})" path="/remarks"/>
    ///
    /// <b>An unsupplied condition is defaulted to <see langword="false"/></b> — the block is
    /// removed, exactly as if the caller had explicitly said not to show it — and named in
    /// <see cref="DocxMailMergeBlockReport.MissingNames"/>. That is the one difference from the
    /// strict overload: this never refuses for a missing name. <b>It still refuses for a genuinely
    /// unbalanced marker structure</b> — an unmatched or mismatched start/end pair makes the
    /// underlying engine throw regardless of what <paramref name="conditions"/> contains, which is
    /// not something any dictionary content can work around.
    /// </remarks>
    /// <param name="docx">The template to resolve.</param>
    /// <param name="conditions">Whether to include each named block.</param>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">
    /// The marker structure is unbalanced, or the document could not be read or written.
    /// </exception>
    public static DocxMailMergeBlockResult MergeConditionalWithReport(
        byte[] docx, IReadOnlyDictionary<string, bool> conditions)
    {
        RequireContent(docx);
        ArgumentNullException.ThrowIfNull(conditions);

        using var source = new MemoryStream(docx, writable: false);
        using var result = MergeConditionalCore(source, conditions, strict: false, out var report);
        return new DocxMailMergeBlockResult(result.ToArray(), report);
    }

    /// <summary>
    /// Writes a copy of the template in <paramref name="source"/> to <paramref name="destination"/>
    /// with every conditional block resolved, and returns which condition names were not supplied.
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed nor sought.
    /// </summary>
    /// <remarks>
    /// <inheritdoc cref="MergeConditionalWithReport(byte[], IReadOnlyDictionary{string, bool})" path="/remarks"/>
    ///
    /// This returns a <see cref="DocxMailMergeBlockReport"/> rather than a
    /// <see cref="DocxMailMergeBlockResult"/> because the document went to
    /// <paramref name="destination"/>.
    /// </remarks>
    /// <param name="source">The template to resolve.</param>
    /// <param name="destination">Receives the resolved document.</param>
    /// <param name="conditions">Whether to include each named block.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    /// <exception cref="ArgumentNullException">A stream or <paramref name="conditions"/> is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, or <paramref name="source"/> held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The marker structure is unbalanced, or the document could not be read or written.
    /// </exception>
    public static async Task<DocxMailMergeBlockReport> MergeConditionalWithReportAsync(
        Stream source, Stream destination, IReadOnlyDictionary<string, bool> conditions,
        CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ArgumentNullException.ThrowIfNull(conditions);
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, EmptySource, nameof(source), ConditionalFailure, ct)
            .ConfigureAwait(false);

        using MemoryStream merged = MergeConditionalCore(docx, conditions, strict: false, out var report);
        await StreamPipeline.EmitAsync(merged, destination, ConditionalFailure, ct).ConfigureAwait(false);
        return report;
    }

    /// <summary>
    /// A copy of <paramref name="docx"/> with every repeating block (<c>{{#each Name}}</c> …
    /// <c>{{/each Name}}</c>) expanded once per entry in its region, merge fields inside each
    /// expansion filled from that entry.
    /// </summary>
    /// <remarks>
    /// <b>Refuses to produce a document with a region the template asks for but
    /// <paramref name="regions"/> did not supply</b> — the same reasoning as
    /// <see cref="MergeConditional(byte[], IReadOnlyDictionary{string, bool})"/>: the underlying
    /// engine throws immediately for an unsupplied name, so this preflights and refuses before the
    /// document is touched. Use
    /// <see cref="MergeRepeatingWithReport(byte[], IReadOnlyDictionary{string, IEnumerable{IReadOnlyDictionary{string, string}}})"/>
    /// when you want the document anyway.
    ///
    /// <b>An empty sequence for a region removes the whole marked region</b> — markers and content
    /// both — measured.
    ///
    /// <b>Run this before <see cref="MergeConditional(byte[], IReadOnlyDictionary{string, bool})"/>
    /// and before <see cref="Merge(byte[], IReadOnlyDictionary{string, string})"/></b>, for the
    /// same reason: a conditional block or a merge field nested inside a repeating region only
    /// exists once this call has expanded it.
    ///
    /// <b>A missing field inside one record's expansion is not caught here</b> — it leaves that
    /// field's raw placeholder in the generated row, silently, because
    /// <c>ExecuteRepeatingBlocks</c> has no report of its own. A follow-up call to
    /// <see cref="Merge(byte[], IReadOnlyDictionary{string, string})"/> or
    /// <see cref="MergeWithReport(byte[], IReadOnlyDictionary{string, string})"/> against the
    /// result — even with an empty values dictionary — finds and reports it, because it scans the
    /// whole document for remaining <c>MERGEFIELD</c>s. Measured.
    /// </remarks>
    /// <param name="docx">The template to expand.</param>
    /// <param name="regions">One sequence of value sets per named repeating region.</param>
    /// <exception cref="ArgumentNullException">An argument is null, or an individual record in it is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a record's value is null.</exception>
    /// <exception cref="DocumentConversionException">
    /// A region the template asks for was not supplied, the marker structure is unbalanced, or the
    /// document could not be read or written.
    /// </exception>
    public static byte[] MergeRepeating(
        byte[] docx, IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> regions)
    {
        RequireContent(docx);
        ArgumentNullException.ThrowIfNull(regions);

        using var source = new MemoryStream(docx, writable: false);
        using var result = MergeRepeatingCore(source, regions, strict: true, out _);
        return result.ToArray();
    }

    /// <summary>
    /// Writes a copy of the template in <paramref name="source"/> to <paramref name="destination"/>
    /// with every repeating block expanded. <paramref name="source"/> is <b>read</b> to its end and
    /// <paramref name="destination"/> is <b>written</b>; neither is disposed, closed nor sought.
    /// </summary>
    /// <inheritdoc cref="MergeRepeating(byte[], IReadOnlyDictionary{string, IEnumerable{IReadOnlyDictionary{string, string}}})" path="/remarks"/>
    /// <param name="source">The template to expand.</param>
    /// <param name="destination">Receives the expanded document.</param>
    /// <param name="regions">One sequence of value sets per named repeating region.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    /// <exception cref="ArgumentNullException">A stream or <paramref name="regions"/> is null, or an individual record in it is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, or <paramref name="source"/> held no bytes, or a record's value is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// A region the template asks for was not supplied, the marker structure is unbalanced, or the
    /// document could not be read or written.
    /// </exception>
    public static async Task MergeRepeatingAsync(
        Stream source, Stream destination,
        IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> regions,
        CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ArgumentNullException.ThrowIfNull(regions);
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, EmptySource, nameof(source), RepeatingFailure, ct)
            .ConfigureAwait(false);

        using MemoryStream merged = MergeRepeatingCore(docx, regions, strict: true, out _);
        await StreamPipeline.EmitAsync(merged, destination, RepeatingFailure, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// A copy of <paramref name="docx"/> with every repeating block expanded, <b>together with
    /// which region names the template asked for that <paramref name="regions"/> did not supply</b>.
    /// Always produces a document, except when the marker structure is genuinely unbalanced.
    /// </summary>
    /// <remarks>
    /// <inheritdoc cref="MergeRepeating(byte[], IReadOnlyDictionary{string, IEnumerable{IReadOnlyDictionary{string, string}}})" path="/remarks"/>
    ///
    /// <b>An unsupplied region is defaulted to zero rows</b> — the whole marked region is removed,
    /// exactly as an explicitly empty sequence would be — and named in
    /// <see cref="DocxMailMergeBlockReport.MissingNames"/>. It still refuses for a genuinely
    /// unbalanced marker structure, which no dictionary content can work around.
    /// </remarks>
    /// <param name="docx">The template to expand.</param>
    /// <param name="regions">One sequence of value sets per named repeating region.</param>
    /// <exception cref="ArgumentNullException">An argument is null, or an individual record in it is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a record's value is null.</exception>
    /// <exception cref="DocumentConversionException">
    /// The marker structure is unbalanced, or the document could not be read or written.
    /// </exception>
    public static DocxMailMergeBlockResult MergeRepeatingWithReport(
        byte[] docx, IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> regions)
    {
        RequireContent(docx);
        ArgumentNullException.ThrowIfNull(regions);

        using var source = new MemoryStream(docx, writable: false);
        using var result = MergeRepeatingCore(source, regions, strict: false, out var report);
        return new DocxMailMergeBlockResult(result.ToArray(), report);
    }

    /// <summary>
    /// Writes a copy of the template in <paramref name="source"/> to <paramref name="destination"/>
    /// with every repeating block expanded, and returns which region names were not supplied.
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed nor sought.
    /// </summary>
    /// <remarks>
    /// <inheritdoc cref="MergeRepeatingWithReport(byte[], IReadOnlyDictionary{string, IEnumerable{IReadOnlyDictionary{string, string}}})" path="/remarks"/>
    ///
    /// This returns a <see cref="DocxMailMergeBlockReport"/> rather than a
    /// <see cref="DocxMailMergeBlockResult"/> because the document went to
    /// <paramref name="destination"/>.
    /// </remarks>
    /// <param name="source">The template to expand.</param>
    /// <param name="destination">Receives the expanded document.</param>
    /// <param name="regions">One sequence of value sets per named repeating region.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    /// <exception cref="ArgumentNullException">A stream or <paramref name="regions"/> is null, or an individual record in it is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, or <paramref name="source"/> held no bytes, or a record's value is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The marker structure is unbalanced, or the document could not be read or written.
    /// </exception>
    public static async Task<DocxMailMergeBlockReport> MergeRepeatingWithReportAsync(
        Stream source, Stream destination,
        IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> regions,
        CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ArgumentNullException.ThrowIfNull(regions);
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, EmptySource, nameof(source), RepeatingFailure, ct)
            .ConfigureAwait(false);

        using MemoryStream merged = MergeRepeatingCore(docx, regions, strict: false, out var report);
        await StreamPipeline.EmitAsync(merged, destination, RepeatingFailure, ct).ConfigureAwait(false);
        return report;
    }

    /// <summary>
    /// A copy of <paramref name="docx"/> with every repeating block expanded once per entry in its
    /// region, the same as
    /// <see cref="MergeRepeating(byte[], IReadOnlyDictionary{string, IEnumerable{IReadOnlyDictionary{string, string}}})"/>
    /// — except each entry may itself carry further nested regions, for a template whose repeating
    /// blocks are nested inside one another.
    /// </summary>
    /// <remarks>
    /// <inheritdoc cref="MergeRepeating(byte[], IReadOnlyDictionary{string, IEnumerable{IReadOnlyDictionary{string, string}}})" path="/remarks"/>
    ///
    /// <b>"Missing" is checked at every nesting level, not just the top one.</b>
    /// <see cref="DocxMailMergeTemplate.RepeatingBlockNames"/> is flat regardless of nesting depth
    /// — a nested marker's name appears in that list exactly like a top-level one, measured — so
    /// this walks <paramref name="regions"/> and every <see cref="DocxMailMergeBlockData.Regions"/>
    /// inside it recursively before comparing against what the template asks for.
    /// </remarks>
    /// <param name="docx">The template to expand.</param>
    /// <param name="regions">One sequence of block rows per named top-level repeating region.</param>
    /// <exception cref="ArgumentNullException">An argument is null, or an individual block row in it is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a block row's value is null.</exception>
    /// <exception cref="DocumentConversionException">
    /// A region the template asks for was not supplied at any nesting level, or was supplied by one
    /// block row and omitted by a sibling — the preflight catches the first, the underlying engine
    /// throws for the second, which is a NAME-level report's blind spot and is why
    /// <see cref="MergeRepeatingRegionsWithReport(byte[], IReadOnlyDictionary{string, IEnumerable{DocxMailMergeBlockData}})"/>
    /// reports nothing missing for it. Also when the marker structure is unbalanced, or the
    /// document could not be read or written.
    /// </exception>
    public static byte[] MergeRepeatingRegions(
        byte[] docx, IReadOnlyDictionary<string, IEnumerable<DocxMailMergeBlockData>> regions)
    {
        RequireContent(docx);
        ArgumentNullException.ThrowIfNull(regions);

        using var source = new MemoryStream(docx, writable: false);
        using var result = MergeRepeatingRegionsCore(source, regions, strict: true, out _);
        return result.ToArray();
    }

    /// <summary>
    /// Writes a copy of the template in <paramref name="source"/> to <paramref name="destination"/>
    /// with every nested repeating region expanded. <paramref name="source"/> is <b>read</b> to its
    /// end and <paramref name="destination"/> is <b>written</b>; neither is disposed, closed nor
    /// sought.
    /// </summary>
    /// <inheritdoc cref="MergeRepeatingRegions(byte[], IReadOnlyDictionary{string, IEnumerable{DocxMailMergeBlockData}})" path="/remarks"/>
    /// <param name="source">The template to expand.</param>
    /// <param name="destination">Receives the expanded document.</param>
    /// <param name="regions">One sequence of block rows per named top-level repeating region.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    /// <exception cref="ArgumentNullException">A stream or <paramref name="regions"/> is null, or an individual block row in it is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, or <paramref name="source"/> held no bytes, or a block row's value is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// A region the template asks for was not supplied at any nesting level, or was supplied by one
    /// block row and omitted by a sibling, the marker structure is unbalanced, or the document could
    /// not be read or written.
    /// </exception>
    public static async Task MergeRepeatingRegionsAsync(
        Stream source, Stream destination,
        IReadOnlyDictionary<string, IEnumerable<DocxMailMergeBlockData>> regions,
        CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ArgumentNullException.ThrowIfNull(regions);
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, EmptySource, nameof(source), RepeatingRegionsFailure, ct)
            .ConfigureAwait(false);

        using MemoryStream merged = MergeRepeatingRegionsCore(docx, regions, strict: true, out _);
        await StreamPipeline.EmitAsync(merged, destination, RepeatingRegionsFailure, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// A copy of <paramref name="docx"/> with every nested repeating region expanded, <b>together
    /// with which region names the template asked for — at any nesting level — that
    /// <paramref name="regions"/> did not supply</b>. Always produces a document, except when the
    /// marker structure is genuinely unbalanced.
    /// </summary>
    /// <remarks>
    /// <inheritdoc cref="MergeRepeatingRegions(byte[], IReadOnlyDictionary{string, IEnumerable{DocxMailMergeBlockData}})" path="/remarks"/>
    ///
    /// <b>An unsupplied region, at any nesting level, is defaulted to zero rows</b> — the whole
    /// marked region is removed — and, when the name is unsupplied <i>everywhere</i>, it is also
    /// named in <see cref="DocxMailMergeBlockReport.MissingNames"/>. It still refuses for a
    /// genuinely unbalanced marker structure.
    ///
    /// <b><see cref="DocxMailMergeBlockReport.MissingNames"/> answers about NAMES, not about
    /// individual block rows</b>, and the difference shows up only under nesting. A name supplied
    /// by <i>any</i> row reads as supplied overall, so a template with <c>Orders</c> containing
    /// <c>Lines</c>, called with one order that carries <c>Lines</c> and a second that does not,
    /// reports nothing missing — the second order's <c>Lines</c> region is still defaulted to zero
    /// rows and removed, silently. Measured: without that default the underlying engine throws for
    /// that second row, exactly as it does for a name missing everywhere.
    ///
    /// <b>Unlike a missing individual MERGEFIELD, which a follow-up <see cref="Merge(byte[],
    /// IReadOnlyDictionary{string, string})"/> or <c>MergeWithReport</c> pass can still detect as a
    /// raw leftover placeholder, a per-row missing nested region leaves no detectable artifact in
    /// the output</b> — the result is byte-indistinguishable from a row that genuinely had no nested
    /// rows, so nothing downstream can recover this information.
    /// </remarks>
    /// <param name="docx">The template to expand.</param>
    /// <param name="regions">One sequence of block rows per named top-level repeating region.</param>
    /// <exception cref="ArgumentNullException">An argument is null, or an individual block row in it is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a block row's value is null.</exception>
    /// <exception cref="DocumentConversionException">
    /// The marker structure is unbalanced, or the document could not be read or written.
    /// </exception>
    public static DocxMailMergeBlockResult MergeRepeatingRegionsWithReport(
        byte[] docx, IReadOnlyDictionary<string, IEnumerable<DocxMailMergeBlockData>> regions)
    {
        RequireContent(docx);
        ArgumentNullException.ThrowIfNull(regions);

        using var source = new MemoryStream(docx, writable: false);
        using var result = MergeRepeatingRegionsCore(source, regions, strict: false, out var report);
        return new DocxMailMergeBlockResult(result.ToArray(), report);
    }

    /// <summary>
    /// Writes a copy of the template in <paramref name="source"/> to <paramref name="destination"/>
    /// with every nested repeating region expanded, and returns which region names were not
    /// supplied at any nesting level. <paramref name="source"/> is <b>read</b> to its end and
    /// <paramref name="destination"/> is <b>written</b>; neither is disposed, closed nor sought.
    /// </summary>
    /// <remarks>
    /// <inheritdoc cref="MergeRepeatingRegionsWithReport(byte[], IReadOnlyDictionary{string, IEnumerable{DocxMailMergeBlockData}})" path="/remarks"/>
    ///
    /// This returns a <see cref="DocxMailMergeBlockReport"/> rather than a
    /// <see cref="DocxMailMergeBlockResult"/> because the document went to
    /// <paramref name="destination"/>.
    /// </remarks>
    /// <param name="source">The template to expand.</param>
    /// <param name="destination">Receives the expanded document.</param>
    /// <param name="regions">One sequence of block rows per named top-level repeating region.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    /// <exception cref="ArgumentNullException">A stream or <paramref name="regions"/> is null, or an individual block row in it is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, or <paramref name="source"/> held no bytes, or a block row's value is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The marker structure is unbalanced, or the document could not be read or written.
    /// </exception>
    public static async Task<DocxMailMergeBlockReport> MergeRepeatingRegionsWithReportAsync(
        Stream source, Stream destination,
        IReadOnlyDictionary<string, IEnumerable<DocxMailMergeBlockData>> regions,
        CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ArgumentNullException.ThrowIfNull(regions);
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, EmptySource, nameof(source), RepeatingRegionsFailure, ct)
            .ConfigureAwait(false);

        using MemoryStream merged = MergeRepeatingRegionsCore(docx, regions, strict: false, out var report);
        await StreamPipeline.EmitAsync(merged, destination, RepeatingRegionsFailure, ct).ConfigureAwait(false);
        return report;
    }

    /// <summary>
    /// A copy of <paramref name="docx"/> with the row at <paramref name="templateRowIndex"/> in the
    /// table at <paramref name="tableIndex"/> repeated once per entry in <paramref name="rows"/>,
    /// merge fields inside each generated row filled from that entry.
    /// </summary>
    /// <remarks>
    /// <b>Index-based, not marker-based</b> — unlike
    /// <see cref="MergeRepeating(byte[], IReadOnlyDictionary{string, IEnumerable{IReadOnlyDictionary{string, string}}})"/>,
    /// there is no <c>{{...}}</c> convention for a table row; the underlying engine selects both the
    /// table and the row by position.
    ///
    /// <b><paramref name="tableIndex"/> counts the tables the underlying engine sees, which is NOT
    /// the same set <see cref="DocxEditor.ReadTable(byte[], int)"/> indexes.</b> Both skip a table
    /// nested inside another table's cell. Only <c>ReadTable</c> descends into a
    /// content control (<c>w:sdt</c>), so a document holding a control-wrapped table followed by an
    /// ordinary one gives <c>ReadTable</c> two tables and this method one — index 0 is the wrapped
    /// table to <c>ReadTable</c> and the ordinary table here, and index 1 is out of range here while
    /// <c>ReadTable</c> answers it. Measured; do not read one index off the other. Content controls
    /// are the only known divergence, so a template with none can use either count.
    ///
    /// <b>No strict/lenient split.</b> A caller supplies <paramref name="rows"/> directly rather
    /// than the template asking for a name that might go unsupplied, so there is nothing to
    /// preflight — matching
    /// <see cref="DocxEditor.FillRows(byte[], string, IEnumerable{IReadOnlyDictionary{string, string}})"/>'s
    /// own single-form shape exactly.
    ///
    /// <b>An empty <paramref name="rows"/> removes the template row.</b> A record missing a field
    /// the row asks for leaves that field's raw placeholder in the generated row, silently — the
    /// same as <see cref="MergeRepeating(byte[], IReadOnlyDictionary{string, IEnumerable{IReadOnlyDictionary{string, string}}})"/>'s
    /// per-record behavior, and caught the same way: a follow-up call to
    /// <see cref="Merge(byte[], IReadOnlyDictionary{string, string})"/> or
    /// <see cref="MergeWithReport(byte[], IReadOnlyDictionary{string, string})"/> finds it.
    ///
    /// <b>Run this before <see cref="MergeConditional(byte[], IReadOnlyDictionary{string, bool})"/>
    /// and before <see cref="Merge(byte[], IReadOnlyDictionary{string, string})"/></b>, and here
    /// the order is sharper than it is for the marker-based methods: a conditional pass removes
    /// content, so running one first can change which table <paramref name="tableIndex"/> lands on.
    /// Position is not a name.
    /// </remarks>
    /// <param name="docx">The template to expand.</param>
    /// <param name="tableIndex">Zero-based index of the table, in document order.</param>
    /// <param name="templateRowIndex">Zero-based row index within that table to clone and bind.</param>
    /// <param name="rows">One value set per generated row.</param>
    /// <exception cref="ArgumentNullException">An argument is null, or an individual row in it is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a row's value is null.</exception>
    /// <exception cref="DocumentConversionException">
    /// <paramref name="tableIndex"/> or <paramref name="templateRowIndex"/> is out of range, or the
    /// document could not be read or written.
    /// </exception>
    public static byte[] MergeTableRows(
        byte[] docx, int tableIndex, int templateRowIndex,
        IEnumerable<IReadOnlyDictionary<string, string>> rows)
    {
        RequireContent(docx);
        ArgumentNullException.ThrowIfNull(rows);

        using var source = new MemoryStream(docx, writable: false);
        using var result = MergeTableRowsCore(source, tableIndex, templateRowIndex, rows);
        return result.ToArray();
    }

    /// <summary>
    /// Writes a copy of the template in <paramref name="source"/> to <paramref name="destination"/>
    /// with the table row expanded. <paramref name="source"/> is <b>read</b> to its end and
    /// <paramref name="destination"/> is <b>written</b>; neither is disposed, closed nor sought.
    /// </summary>
    /// <inheritdoc cref="MergeTableRows(byte[], int, int, IEnumerable{IReadOnlyDictionary{string, string}})" path="/remarks"/>
    /// <param name="source">The template to expand.</param>
    /// <param name="destination">Receives the expanded document.</param>
    /// <param name="tableIndex">Zero-based index of the table, in document order.</param>
    /// <param name="templateRowIndex">Zero-based row index within that table to clone and bind.</param>
    /// <param name="rows">One value set per generated row.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    /// <exception cref="ArgumentNullException">A stream or <paramref name="rows"/> is null, or an individual row in it is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, or <paramref name="source"/> held no bytes, or a row's value is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// <paramref name="tableIndex"/> or <paramref name="templateRowIndex"/> is out of range, or the
    /// document could not be read or written.
    /// </exception>
    public static async Task MergeTableRowsAsync(
        Stream source, Stream destination, int tableIndex, int templateRowIndex,
        IEnumerable<IReadOnlyDictionary<string, string>> rows, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ArgumentNullException.ThrowIfNull(rows);
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, EmptySource, nameof(source), TableRowsFailure, ct)
            .ConfigureAwait(false);

        using MemoryStream merged = MergeTableRowsCore(docx, tableIndex, templateRowIndex, rows);
        await StreamPipeline.EmitAsync(merged, destination, TableRowsFailure, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// A copy of <paramref name="docx"/> with a group/header row and its detail row template,
    /// both in the table at <paramref name="tableIndex"/>, repeated once per group in
    /// <paramref name="groups"/>.
    /// </summary>
    /// <remarks>
    /// <b>Index-based, not marker-based</b>, exactly as
    /// <see cref="MergeTableRows(byte[], int, int, IEnumerable{IReadOnlyDictionary{string, string}})"/>
    /// is — and <paramref name="tableIndex"/> counts the same set of tables, which is <b>not</b> the
    /// set <see cref="DocxEditor.ReadTable(byte[], int)"/> indexes. See that method's remarks for
    /// the measured difference: a table wrapped in a content control is one of <c>ReadTable</c>'s
    /// and none of this one's.
    ///
    /// <b>No strict/lenient split</b>, for the same reason: a caller supplies
    /// <paramref name="groups"/> directly rather than the template asking for a name that might go
    /// unsupplied, so there is nothing to preflight.
    ///
    /// <b>An empty <paramref name="groups"/> removes both template rows.</b> A group or detail
    /// record missing a field the corresponding row asks for leaves that field's raw placeholder in
    /// the generated row, silently, and is caught the same way — a follow-up call to
    /// <see cref="Merge(byte[], IReadOnlyDictionary{string, string})"/> or
    /// <see cref="MergeWithReport(byte[], IReadOnlyDictionary{string, string})"/> finds it.
    ///
    /// <b>Run this before <see cref="MergeConditional(byte[], IReadOnlyDictionary{string, bool})"/>
    /// and before <see cref="Merge(byte[], IReadOnlyDictionary{string, string})"/></b> — a
    /// conditional pass removes content, so running one first can change which table
    /// <paramref name="tableIndex"/> lands on. Position is not a name.
    /// </remarks>
    /// <param name="docx">The template to expand.</param>
    /// <param name="tableIndex">Zero-based index of the table, in document order.</param>
    /// <param name="groupTemplateRowIndex">Zero-based row index of the group/header row template.</param>
    /// <param name="detailTemplateRowIndex">Zero-based row index of the detail row template.</param>
    /// <param name="groups">One group/header value set, with its detail rows, per generated group.</param>
    /// <exception cref="ArgumentNullException">An argument is null, or an individual group or detail row in it is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a group or detail row's value is null.</exception>
    /// <exception cref="DocumentConversionException">
    /// <paramref name="tableIndex"/>, <paramref name="groupTemplateRowIndex"/> or
    /// <paramref name="detailTemplateRowIndex"/> is out of range, or the document could not be read
    /// or written.
    /// </exception>
    public static byte[] MergeTableRowGroups(
        byte[] docx, int tableIndex, int groupTemplateRowIndex, int detailTemplateRowIndex,
        IEnumerable<DocxMailMergeTableRowGroup> groups)
    {
        RequireContent(docx);
        ArgumentNullException.ThrowIfNull(groups);

        using var source = new MemoryStream(docx, writable: false);
        using var result = MergeTableRowGroupsCore(
            source, tableIndex, groupTemplateRowIndex, detailTemplateRowIndex, groups);
        return result.ToArray();
    }

    /// <summary>
    /// Writes a copy of the template in <paramref name="source"/> to <paramref name="destination"/>
    /// with the group and detail rows expanded. <paramref name="source"/> is <b>read</b> to its end
    /// and <paramref name="destination"/> is <b>written</b>; neither is disposed, closed nor sought.
    /// </summary>
    /// <inheritdoc cref="MergeTableRowGroups(byte[], int, int, int, IEnumerable{DocxMailMergeTableRowGroup})" path="/remarks"/>
    /// <param name="source">The template to expand.</param>
    /// <param name="destination">Receives the expanded document.</param>
    /// <param name="tableIndex">Zero-based index of the table, in document order.</param>
    /// <param name="groupTemplateRowIndex">Zero-based row index of the group/header row template.</param>
    /// <param name="detailTemplateRowIndex">Zero-based row index of the detail row template.</param>
    /// <param name="groups">One group/header value set, with its detail rows, per generated group.</param>
    /// <param name="ct">Cancels before the document is read, and while it is written.</param>
    /// <exception cref="ArgumentNullException">A stream or <paramref name="groups"/> is null, or an individual group or detail row in it is null.</exception>
    /// <exception cref="ArgumentException">A stream is unusable, or <paramref name="source"/> held no bytes, or a group or detail row's value is null.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// <paramref name="tableIndex"/>, <paramref name="groupTemplateRowIndex"/> or
    /// <paramref name="detailTemplateRowIndex"/> is out of range, or the document could not be read
    /// or written.
    /// </exception>
    public static async Task MergeTableRowGroupsAsync(
        Stream source, Stream destination, int tableIndex, int groupTemplateRowIndex,
        int detailTemplateRowIndex, IEnumerable<DocxMailMergeTableRowGroup> groups,
        CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ArgumentNullException.ThrowIfNull(groups);
        ct.ThrowIfCancellationRequested();

        using var docx = await StreamPipeline
            .DrainAsync(source, EmptySource, nameof(source), TableRowGroupsFailure, ct)
            .ConfigureAwait(false);

        using MemoryStream merged = MergeTableRowGroupsCore(
            docx, tableIndex, groupTemplateRowIndex, detailTemplateRowIndex, groups);
        await StreamPipeline.EmitAsync(merged, destination, TableRowGroupsFailure, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Fills <paramref name="docx"/> once per entry in <paramref name="records"/>, yielding each
    /// filled document in order.
    /// </summary>
    /// <remarks>
    /// <b>Strict, the same way <see cref="Merge(byte[], IReadOnlyDictionary{string, string})"/> is
    /// strict — this refuses the moment a record is incomplete, mid-sequence.</b> Everything already
    /// yielded before that point is unaffected; nothing after it runs. See
    /// <see cref="MergeBatchWithReport(byte[], IEnumerable{IReadOnlyDictionary{string, string}})"/>
    /// for the lenient form, which never throws for an incomplete record.
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

    /// <summary>
    /// The async form of <see cref="MergeBatch(byte[], IEnumerable{IReadOnlyDictionary{string, string}})"/>
    /// — see its documentation for exactly what is matched and how strictness works.
    /// </summary>
    /// <inheritdoc cref="MergeBatch(byte[], IEnumerable{IReadOnlyDictionary{string, string}})" path="/remarks"/>
    /// <param name="docx">The template to fill, once per record.</param>
    /// <param name="records">
    /// One dictionary of values per output document, matched case-insensitively — see
    /// <see cref="Merge(byte[], IReadOnlyDictionary{string, string})"/> for the matching rules,
    /// which apply here unchanged. An empty sequence yields no documents.
    /// </param>
    /// <param name="ct">Cancels before the next record's merge runs.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="docx"/> or <paramref name="records"/> is null, or an individual record in
    /// <paramref name="records"/> is null. Unlike the synchronous
    /// <see cref="MergeBatch(byte[], IEnumerable{IReadOnlyDictionary{string, string}})"/>, this is
    /// not thrown until the caller starts enumerating the result — inherent to how an
    /// <see cref="IAsyncEnumerable{T}"/> iterator method defers its whole body, argument validation
    /// included, not a gap specific to this method.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="docx"/> is empty, or a record's value is null. Unlike the synchronous
    /// <see cref="MergeBatch(byte[], IEnumerable{IReadOnlyDictionary{string, string}})"/>, this is
    /// not thrown until the caller starts enumerating the result — the same
    /// <see cref="IAsyncEnumerable{T}"/> deferral as the <see cref="ArgumentNullException"/> case
    /// above, not a gap specific to this method.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The package could not be read or written, or a record is missing a value for a field the
    /// template requires — the message names the record's position (0-based) and the missing
    /// field(s).
    /// </exception>
    public static async IAsyncEnumerable<byte[]> MergeBatchAsync(
        byte[] docx, IEnumerable<IReadOnlyDictionary<string, string>> records,
        [EnumeratorCancellation] CancellationToken ct = default)
    {
        RequireContent(docx);
        ArgumentNullException.ThrowIfNull(records);

        using var e = MergeBatchCore(docx, records, strict: true).GetEnumerator();
        while (true)
        {
            ct.ThrowIfCancellationRequested();
            if (!e.MoveNext()) break;
            yield return e.Current.Document;
        }
    }

    /// <summary>
    /// Fills <paramref name="docx"/> once per entry in <paramref name="records"/>, yielding each
    /// record's document <b>together with what happened to every field in it</b>.
    /// </summary>
    /// <remarks>
    /// The lenient half of the pair. This always produces a document for every record, complete or
    /// not, and never throws for an incomplete one — see
    /// <see cref="MergeBatch(byte[], IEnumerable{IReadOnlyDictionary{string, string}})"/> for the
    /// strict form.
    ///
    /// <b>This is lazy.</b> Memory stays proportional to one item in flight, not the whole batch —
    /// <paramref name="records"/> is walked one entry at a time as the caller enumerates the
    /// result, and each item is only held until the caller moves on to the next one.
    /// </remarks>
    /// <param name="docx">The template to fill, once per record.</param>
    /// <param name="records">
    /// One dictionary of values per output document, matched case-insensitively. An empty sequence
    /// yields no items.
    /// </param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="docx"/> or <paramref name="records"/> is null, or an individual record in
    /// <paramref name="records"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> is empty, or a record's value is null.</exception>
    /// <exception cref="DocumentConversionException">The package could not be read or written.</exception>
    public static IEnumerable<DocxMailMergeBatchItem> MergeBatchWithReport(
        byte[] docx, IEnumerable<IReadOnlyDictionary<string, string>> records)
    {
        RequireContent(docx);
        ArgumentNullException.ThrowIfNull(records);

        return MergeBatchCore(docx, records, strict: false);
    }

    /// <summary>
    /// The async form of <see cref="MergeBatchWithReport(byte[], IEnumerable{IReadOnlyDictionary{string, string}})"/>
    /// — see its documentation for exactly what is matched and how lenience works.
    /// </summary>
    /// <inheritdoc cref="MergeBatchWithReport(byte[], IEnumerable{IReadOnlyDictionary{string, string}})" path="/remarks"/>
    /// <param name="docx">The template to fill, once per record.</param>
    /// <param name="records">
    /// One dictionary of values per output document, matched case-insensitively. An empty sequence
    /// yields no items.
    /// </param>
    /// <param name="ct">Cancels before the next record's merge runs.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="docx"/> or <paramref name="records"/> is null, or an individual record in
    /// <paramref name="records"/> is null. Unlike the synchronous
    /// <see cref="MergeBatchWithReport(byte[], IEnumerable{IReadOnlyDictionary{string, string}})"/>, this is
    /// not thrown until the caller starts enumerating the result — inherent to how an
    /// <see cref="IAsyncEnumerable{T}"/> iterator method defers its whole body, argument validation
    /// included, not a gap specific to this method.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="docx"/> is empty, or a record's value is null. Unlike the synchronous
    /// <see cref="MergeBatchWithReport(byte[], IEnumerable{IReadOnlyDictionary{string, string}})"/>,
    /// this is not thrown until the caller starts enumerating the result — the same
    /// <see cref="IAsyncEnumerable{T}"/> deferral as the <see cref="ArgumentNullException"/> case
    /// above, not a gap specific to this method.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be read or written.</exception>
    public static async IAsyncEnumerable<DocxMailMergeBatchItem> MergeBatchWithReportAsync(
        byte[] docx, IEnumerable<IReadOnlyDictionary<string, string>> records,
        [EnumeratorCancellation] CancellationToken ct = default)
    {
        RequireContent(docx);
        ArgumentNullException.ThrowIfNull(records);

        using var e = MergeBatchCore(docx, records, strict: false).GetEnumerator();
        while (true)
        {
            ct.ThrowIfCancellationRequested();
            if (!e.MoveNext()) break;
            yield return e.Current;
        }
    }

    /// <summary>
    /// Reads a template from <paramref name="templatePath"/>, fills it once per entry in
    /// <paramref name="records"/>, and writes each result to the path
    /// <paramref name="outputPathFactory"/> returns for it.
    /// </summary>
    /// <remarks>
    /// <b>Strict, the same way <see cref="MergeBatch(byte[], IEnumerable{IReadOnlyDictionary{string, string}})"/>
    /// is strict</b> — refuses the moment a record is incomplete, and nothing after that record is
    /// written. See <see cref="MergeBatchToFilesWithReport(string, IEnumerable{IReadOnlyDictionary{string, string}}, Func{int, IReadOnlyDictionary{string, string}, string})"/>
    /// for the lenient form.
    ///
    /// <b>Every output path is checked for a collision against every other, before anything is
    /// written.</b> Two records producing the same path is refused outright, naming both record
    /// indices — measured against the underlying engine's own batch writer, a collision silently
    /// overwrites one record's document with another's, with no exception and no warning. This
    /// refuses rather than risk it. Paths are compared as exact strings, not resolved or
    /// normalized — two different spellings of the same file (a relative path and its absolute
    /// equivalent, or two different cases on a case-insensitive filesystem) are not detected as a
    /// collision.
    ///
    /// <b><paramref name="templatePath"/> itself is not one of the paths this check compares
    /// against.</b> <paramref name="outputPathFactory"/> returning the template's own path is not
    /// treated as a collision — the template is already fully read into memory before any record is
    /// merged, so nothing about the write itself fails, but the result is that the template file on
    /// disk is silently overwritten with a merged record's output, with no exception and no warning.
    /// </remarks>
    /// <param name="templatePath">The template to fill, once per record.</param>
    /// <param name="records">
    /// One dictionary of values per output document, matched case-insensitively — see
    /// <see cref="Merge(byte[], IReadOnlyDictionary{string, string})"/> for the matching rules. An
    /// empty sequence writes nothing.
    /// </param>
    /// <param name="outputPathFactory">
    /// Given a record's 0-based index and its own values, returns the path its document is written
    /// to. Called once per record before any document is merged.
    /// </param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="templatePath"/>, <paramref name="records"/> or
    /// <paramref name="outputPathFactory"/> is null, or an individual record is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="templatePath"/> is blank, a record's value is null, or
    /// <paramref name="outputPathFactory"/> produced a null/blank path, or the same path for two
    /// different records.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="templatePath"/> does not exist.</exception>
    /// <exception cref="DocumentConversionException">
    /// The package could not be read or written, or a record is missing a value for a field the
    /// template requires — the message names the record's position (0-based) and the missing
    /// field(s).
    /// </exception>
    public static IReadOnlyList<string> MergeBatchToFiles(
        string templatePath, IEnumerable<IReadOnlyDictionary<string, string>> records,
        Func<int, IReadOnlyDictionary<string, string>, string> outputPathFactory)
    {
        var docx = FilePipeline.Read(templatePath, nameof(templatePath));
        ArgumentNullException.ThrowIfNull(records);
        ArgumentNullException.ThrowIfNull(outputPathFactory);

        var items = MergeBatchToFilesCore(docx, [.. records], outputPathFactory, strict: true, CancellationToken.None);
        return [.. items.Select(item => item.OutputPath)];
    }

    /// <summary>
    /// The async form of <see cref="MergeBatchToFiles(string, IEnumerable{IReadOnlyDictionary{string, string}}, Func{int, IReadOnlyDictionary{string, string}, string})"/>
    /// — see its documentation for exactly what is matched, how strictness works, and how the
    /// path-collision guard works.
    /// </summary>
    /// <inheritdoc cref="MergeBatchToFiles(string, IEnumerable{IReadOnlyDictionary{string, string}}, Func{int, IReadOnlyDictionary{string, string}, string})" path="/remarks"/>
    /// <param name="templatePath">The template to fill, once per record.</param>
    /// <param name="records">
    /// One dictionary of values per output document, matched case-insensitively — see
    /// <see cref="Merge(byte[], IReadOnlyDictionary{string, string})"/> for the matching rules. An
    /// empty sequence writes nothing.
    /// </param>
    /// <param name="outputPathFactory">
    /// Given a record's 0-based index and its own values, returns the path its document is written
    /// to. Called once per record before any document is merged.
    /// </param>
    /// <param name="ct">Cancels before the template is read, and again before each record's merge.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="templatePath"/>, <paramref name="records"/> or
    /// <paramref name="outputPathFactory"/> is null, or an individual record is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="templatePath"/> is blank, a record's value is null, or
    /// <paramref name="outputPathFactory"/> produced a null/blank path, or the same path for two
    /// different records.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="templatePath"/> does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled before the template finished reading.</exception>
    /// <exception cref="DocumentConversionException">
    /// The package could not be read or written, or a record is missing a value for a field the
    /// template requires — the message names the record's position (0-based) and the missing
    /// field(s).
    /// </exception>
    public static async Task<IReadOnlyList<string>> MergeBatchToFilesAsync(
        string templatePath, IEnumerable<IReadOnlyDictionary<string, string>> records,
        Func<int, IReadOnlyDictionary<string, string>, string> outputPathFactory,
        CancellationToken ct = default)
    {
        var docx = await FilePipeline.ReadAsync(templatePath, nameof(templatePath), ct)
            .ConfigureAwait(false);
        ArgumentNullException.ThrowIfNull(records);
        ArgumentNullException.ThrowIfNull(outputPathFactory);

        var items = MergeBatchToFilesCore(docx, [.. records], outputPathFactory, strict: true, ct);
        return [.. items.Select(item => item.OutputPath)];
    }

    /// <summary>
    /// Reads a template from <paramref name="templatePath"/>, fills it once per entry in
    /// <paramref name="records"/>, and writes each result to the path
    /// <paramref name="outputPathFactory"/> returns for it — <b>together with what happened to
    /// every field in it.</b>
    /// </summary>
    /// <remarks>
    /// The lenient half of the pair. This always writes a file for every record, complete or not,
    /// and never throws for an incomplete one — see
    /// <see cref="MergeBatchToFiles(string, IEnumerable{IReadOnlyDictionary{string, string}}, Func{int, IReadOnlyDictionary{string, string}, string})"/>
    /// for the strict form. The path-collision refusal is unconditional and applies here too — see
    /// that method's remarks.
    /// </remarks>
    /// <param name="templatePath">The template to fill, once per record.</param>
    /// <param name="records">
    /// One dictionary of values per output document, matched case-insensitively. An empty sequence
    /// writes nothing.
    /// </param>
    /// <param name="outputPathFactory">
    /// Given a record's 0-based index and its own values, returns the path its document is written
    /// to. Called once per record before any document is merged.
    /// </param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="templatePath"/>, <paramref name="records"/> or
    /// <paramref name="outputPathFactory"/> is null, or an individual record is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="templatePath"/> is blank, a record's value is null,
    /// <paramref name="outputPathFactory"/> produced a null/blank path, or
    /// <paramref name="outputPathFactory"/> produced the same path for two different records.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="templatePath"/> does not exist.</exception>
    /// <exception cref="DocumentConversionException">The package could not be read or written.</exception>
    public static IReadOnlyList<DocxMailMergeFileBatchItem> MergeBatchToFilesWithReport(
        string templatePath, IEnumerable<IReadOnlyDictionary<string, string>> records,
        Func<int, IReadOnlyDictionary<string, string>, string> outputPathFactory)
    {
        var docx = FilePipeline.Read(templatePath, nameof(templatePath));
        ArgumentNullException.ThrowIfNull(records);
        ArgumentNullException.ThrowIfNull(outputPathFactory);

        return MergeBatchToFilesCore(docx, [.. records], outputPathFactory, strict: false, CancellationToken.None);
    }

    /// <summary>
    /// The async form of <see cref="MergeBatchToFilesWithReport(string, IEnumerable{IReadOnlyDictionary{string, string}}, Func{int, IReadOnlyDictionary{string, string}, string})"/>
    /// — see its documentation for exactly what is matched, how strictness works, and how the
    /// path-collision guard works.
    /// </summary>
    /// <inheritdoc cref="MergeBatchToFilesWithReport(string, IEnumerable{IReadOnlyDictionary{string, string}}, Func{int, IReadOnlyDictionary{string, string}, string})" path="/remarks"/>
    /// <param name="templatePath">The template to fill, once per record.</param>
    /// <param name="records">
    /// One dictionary of values per output document, matched case-insensitively. An empty sequence
    /// writes nothing.
    /// </param>
    /// <param name="outputPathFactory">
    /// Given a record's 0-based index and its own values, returns the path its document is written
    /// to. Called once per record before any document is merged.
    /// </param>
    /// <param name="ct">Cancels before the template is read, and again before each record's merge.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="templatePath"/>, <paramref name="records"/> or
    /// <paramref name="outputPathFactory"/> is null, or an individual record is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="templatePath"/> is blank, a record's value is null, or
    /// <paramref name="outputPathFactory"/> produced a null/blank path, or the same path for two
    /// different records.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="templatePath"/> does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled before the template finished reading.</exception>
    /// <exception cref="DocumentConversionException">The package could not be read or written.</exception>
    public static async Task<IReadOnlyList<DocxMailMergeFileBatchItem>> MergeBatchToFilesWithReportAsync(
        string templatePath, IEnumerable<IReadOnlyDictionary<string, string>> records,
        Func<int, IReadOnlyDictionary<string, string>, string> outputPathFactory,
        CancellationToken ct = default)
    {
        var docx = await FilePipeline.ReadAsync(templatePath, nameof(templatePath), ct)
            .ConfigureAwait(false);
        ArgumentNullException.ThrowIfNull(records);
        ArgumentNullException.ThrowIfNull(outputPathFactory);

        return MergeBatchToFilesCore(docx, [.. records], outputPathFactory, strict: false, ct);
    }

    private const string EmptySource = "DOCX content was empty.";
    private const string InspectFailure = "Failed to read the document's mail-merge template. See the inner exception for details.";
    private const string MergeFailure = "Failed to fill the document's merge fields. See the inner exception for details.";
    private const string ConditionalFailure = "Failed to resolve the document's conditional blocks. See the inner exception for details.";
    private const string RepeatingFailure = "Failed to expand the document's repeating blocks. See the inner exception for details.";
    private const string RepeatingRegionsFailure = "Failed to expand the document's nested repeating regions. See the inner exception for details.";
    private const string TableRowsFailure = "Failed to expand the table's rows. See the inner exception for details.";
    private const string TableRowGroupsFailure = "Failed to expand the table's row groups. See the inner exception for details.";

    /// <summary>
    /// Stands in for a block row's absent <see cref="DocxMailMergeBlockData.Regions"/> so
    /// <see cref="ToEngineRegions"/> has something to pad into. Never written to.
    /// </summary>
    private static readonly IReadOnlyDictionary<string, IEnumerable<DocxMailMergeBlockData>> NoNestedRegions
        = new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>();

    /// <summary>
    /// The issue kinds a strict <c>MergeConditional*</c> call refuses on — the ones belonging to the
    /// construct it guards, and nothing else.
    /// </summary>
    /// <remarks>
    /// <b>Deliberately not <c>InspectTemplate</c>'s own <c>IsValid</c>, which is what this used to
    /// test and was wrong.</b> That flag is false whenever the inspection found <i>anything</i>, and
    /// measured against OfficeIMO.Word 3.2.6, plenty of what it finds has no bearing on conditional
    /// blocks at all: one malformed <c>MERGEFIELD</c> anywhere in the body, a <c>\b</c>, <c>\f</c> or
    /// <c>\v</c> switch on an unrelated field, or a Word-native <c>NEXT</c>/<c>MERGEREC</c> control
    /// field — the last two being things a real Word mail-merge template carries routinely.
    /// Refusing on those closed this method to the whole document under a message naming the wrong
    /// problem, and took the documented composition pipeline down with it: the field-level
    /// <see cref="Merge(byte[], IReadOnlyDictionary{string, string})"/> that runs last refuses only
    /// for an unfilled field, so it would happily have produced the document these three would not
    /// let it reach. (Measured: <c>\#</c>, <c>\@</c> and <c>\*</c> switches report nothing, so an
    /// ordinary currency- or date-formatted field was never part of the blast radius.)
    ///
    /// <b>Only the refusal narrows.</b> <see cref="DocxMailMergeBlockReport.Issues"/> still carries
    /// everything the inspection found, relevant or not — which is what keeps an unrelated problem
    /// visible to a caller who asked for a report rather than a refusal.
    /// </remarks>
    private static readonly DocxMailMergeIssueKind[] ConditionalIssueKinds =
    [
        DocxMailMergeIssueKind.MissingConditionalValue,
        DocxMailMergeIssueKind.UnmatchedConditionalStart,
        DocxMailMergeIssueKind.UnmatchedConditionalEnd,
        DocxMailMergeIssueKind.MismatchedConditionalEnd,
    ];

    /// <summary>
    /// The repeating-block twin of <see cref="ConditionalIssueKinds"/>, shared by
    /// <see cref="MergeRepeatingCore"/> and <see cref="MergeRepeatingRegionsCore"/> — see that
    /// field's remarks for why a strict refusal is decided by issue kind rather than by
    /// <c>IsValid</c>.
    /// </summary>
    private static readonly DocxMailMergeIssueKind[] RepeatingIssueKinds =
    [
        DocxMailMergeIssueKind.MissingRepeatingBlockData,
        DocxMailMergeIssueKind.UnmatchedRepeatingBlockStart,
        DocxMailMergeIssueKind.UnmatchedRepeatingBlockEnd,
        DocxMailMergeIssueKind.MismatchedRepeatingBlockEnd,
    ];

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
    /// <remarks>
    /// Takes the parameter name explicitly so a caller whose own parameter is not literally called
    /// <c>values</c> — <see cref="MergeBatchCore"/>'s <c>records</c>, for one — reports against the
    /// name it actually declared, rather than one it never did.
    ///
    /// <paramref name="position"/> says WHICH value set the bad value was in, for the callers whose
    /// parameter carries many — a template with fifty rows in it names one, not "somewhere in
    /// <c>rows</c>". It follows the <c>Record {index}:</c> shape <see cref="MergeBatchCore"/>
    /// already uses when it re-throws a per-record failure, so a caller reads one convention rather
    /// than two. Null for the single-value callers, whose parameter identifies the set on its own.
    /// </remarks>
    private static void RequireValues(
        IReadOnlyDictionary<string, string> values, string paramName, string? position = null)
    {
        ArgumentNullException.ThrowIfNull(values, paramName);

        foreach (KeyValuePair<string, string> pair in values)
        {
            if (pair.Value is null)
            {
                throw new ArgumentException(
                    (position is null ? string.Empty : position + ": ")
                    + $"The value for '{pair.Key}' is null. A null merges as an empty string and is "
                    + "reported complete, so it cannot be told apart from a value somebody chose. "
                    + "Pass string.Empty to mean \"leave it blank\".",
                    paramName);
            }
        }
    }

    /// <summary>
    /// A list-backed copy of <paramref name="regions"/>, with every record checked by
    /// <see cref="RequireValues"/> on the way through.
    /// </summary>
    /// <remarks>
    /// <b>The copy and the check are one pass on purpose.</b> The class guarantees a null value is
    /// refused rather than merged, and honouring that here means reading every record — which, over
    /// a caller's own <see cref="IEnumerable{T}"/>, is a walk. Checking in a pass of its own would
    /// leave the engine's later walk as a SECOND one, and a genuinely single-pass source (an
    /// iterator over a <c>DbDataReader</c>, say) comes back empty on that second walk: a document
    /// that looks complete and has quietly lost every row. Materialising once removes the hazard
    /// rather than documenting it.
    ///
    /// The records themselves are not copied — <see cref="ToEngineRecords"/> already does that
    /// before the engine sees them, and copying twice would say nothing extra.
    /// </remarks>
    private static Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> MaterialiseRecords(
        IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> regions,
        string paramName)
    {
        var copy = new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>(regions.Count);
        foreach (KeyValuePair<string, IEnumerable<IReadOnlyDictionary<string, string>>> pair in regions)
        {
            ArgumentNullException.ThrowIfNull(pair.Value, paramName);

            var records = new List<IReadOnlyDictionary<string, string>>();
            foreach (IReadOnlyDictionary<string, string> record in pair.Value)
            {
                ArgumentNullException.ThrowIfNull(record, paramName);
                RequireValues(record, paramName, $"Region '{pair.Key}' record {records.Count}");
                records.Add(record);
            }

            copy[pair.Key] = records;
        }

        return copy;
    }

    /// <summary>
    /// The nested twin of <see cref="MaterialiseRecords"/> — a list-backed copy of a whole
    /// <see cref="DocxMailMergeBlockData"/> tree, every level of it, with every row's values
    /// checked on the way through.
    /// </summary>
    /// <remarks>
    /// <b>This one closes a bug rather than merely preventing one.</b>
    /// <see cref="MergeRepeatingRegionsCore"/> walks the caller's tree twice —
    /// <see cref="CollectNamesRecursively"/> to learn which names are supplied, then
    /// <see cref="ToEngineRegions"/> to build what the engine takes — and the two walked the
    /// caller's own sequences independently. Measured against a source that yields on its first
    /// walk and nothing after: <c>MergeRepeatingRegions</c> produced a valid, empty document, no
    /// exception and no report entry. Taking the copy here, before either walker runs, is what
    /// makes both of them read the same rows.
    ///
    /// <paramref name="path"/> accumulates the enclosing region and record on the way down, so a
    /// null value three levels in names all three rather than only the innermost.
    /// </remarks>
    private static Dictionary<string, IEnumerable<DocxMailMergeBlockData>> MaterialiseBlockData(
        IReadOnlyDictionary<string, IEnumerable<DocxMailMergeBlockData>> regions,
        string paramName, string? path)
    {
        var copy = new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>(regions.Count);
        foreach (KeyValuePair<string, IEnumerable<DocxMailMergeBlockData>> pair in regions)
        {
            ArgumentNullException.ThrowIfNull(pair.Value, paramName);

            var rows = new List<DocxMailMergeBlockData>();
            foreach (DocxMailMergeBlockData row in pair.Value)
            {
                ArgumentNullException.ThrowIfNull(row, paramName);

                string here = path is null
                    ? $"Region '{pair.Key}' record {rows.Count}"
                    : $"{path} -> region '{pair.Key}' record {rows.Count}";
                RequireValues(row.Values, paramName, here);

                rows.Add(row.Regions is null
                    ? row
                    : new DocxMailMergeBlockData(
                        row.Values, MaterialiseBlockData(row.Regions, paramName, here)));
            }

            copy[pair.Key] = rows;
        }

        return copy;
    }

    /// <summary>
    /// The rows for <see cref="OfficeIMOMailMerge.ExecuteTableRows"/>, checked and materialised in
    /// one pass — see <see cref="MaterialiseRecords"/> for why those are the same pass.
    /// </summary>
    private static List<IDictionary<string, string>> ToEngineRows(
        IEnumerable<IReadOnlyDictionary<string, string>> rows, string paramName)
    {
        var result = new List<IDictionary<string, string>>();
        foreach (IReadOnlyDictionary<string, string> row in rows)
        {
            ArgumentNullException.ThrowIfNull(row, paramName);
            RequireValues(row, paramName, $"Row {result.Count}");
            result.Add(Copy(row));
        }

        return result;
    }

    /// <summary>
    /// The groups for <see cref="OfficeIMOMailMerge.ExecuteTableRowGroups"/>, checked and
    /// materialised in one pass — the grouped twin of <see cref="ToEngineRows"/>. A group's own
    /// values and each of its detail rows are checked separately, and named separately, because a
    /// null in a header cell and a null in a detail cell are different mistakes.
    /// </summary>
    private static List<OfficeIMOTableRowGroup> ToEngineGroups(
        IEnumerable<DocxMailMergeTableRowGroup> groups, string paramName)
    {
        var result = new List<OfficeIMOTableRowGroup>();
        foreach (DocxMailMergeTableRowGroup group in groups)
        {
            ArgumentNullException.ThrowIfNull(group, paramName);
            RequireValues(group.Values, paramName, $"Group {result.Count}");

            var rows = new List<IDictionary<string, string>>();
            foreach (IReadOnlyDictionary<string, string> row in group.Rows)
            {
                ArgumentNullException.ThrowIfNull(row, paramName);
                RequireValues(row, paramName, $"Group {result.Count} row {rows.Count}");
                rows.Add(Copy(row));
            }

            result.Add(new OfficeIMOTableRowGroup(Copy(group.Values), rows));
        }

        return result;
    }

    private static async Task<DocxMailMergeReport> MergeToStreamAsync(
        Stream source, Stream destination, IReadOnlyDictionary<string, string> values, bool strict,
        CancellationToken ct)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        RequireValues(values, nameof(values));
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
    /// The one implementation behind all four <c>MergeConditional*</c> overloads.
    /// </summary>
    /// <remarks>
    /// <b>Preflights via <see cref="OfficeIMOMailMerge.InspectTemplate"/> before ever calling
    /// <see cref="OfficeIMOMailMerge.ExecuteConditionalBlocks"/></b> — measured, the engine throws
    /// immediately if its dictionary is missing an entry for a marker name the document contains,
    /// with no partial-tolerance mode. The strict path refuses here, before <c>Execute</c> is
    /// reached, whenever a CONDITIONAL-BLOCK issue is reported — see
    /// <see cref="ConditionalIssueKinds"/> for why the refusal is decided by issue kind rather than
    /// by the inspection's own <c>IsValid</c> — which is also what keeps a genuinely unbalanced
    /// marker's raw <see cref="InvalidOperationException"/> from ever escaping the strict overloads.
    /// The lenient path instead pads a COPY of <paramref name="conditions"/> with
    /// <see langword="false"/> for every missing name before calling <c>Execute</c>, so
    /// <c>Execute</c> never sees an unsupplied key either way.
    ///
    /// <b>The strict refusal happens INSIDE the try, using a captured flag rather than an early
    /// throw</b> — matching <see cref="MergeCore"/>'s own reasoning: throwing the specific
    /// <see cref="DocumentConversionException"/> from inside an unfiltered
    /// <c>catch (Exception ex)</c> below it would re-wrap it in the generic failure message.
    /// </remarks>
    private static MemoryStream MergeConditionalCore(
        Stream source, IReadOnlyDictionary<string, bool> conditions, bool strict,
        out DocxMailMergeBlockReport report)
    {
        var result = new MemoryStream();
        List<string> missing;
        List<DocxMailMergeIssue> allIssues;
        bool refuse;
        string? refusalMessage = null;

        try
        {
            using var document = OfficeIMOWordDocument.Load(source);

            var inspection = OfficeIMOMailMerge.InspectTemplate(document, null!, conditions.Keys, null!);
            allIssues = [.. inspection.Issues.Select(Issue)];
            missing = [.. allIssues
                .Where(i => i.Kind == DocxMailMergeIssueKind.MissingConditionalValue)
                .Select(i => i.Name)];

            List<DocxMailMergeIssue> blocking =
                [.. allIssues.Where(i => ConditionalIssueKinds.Contains(i.Kind))];

            refuse = strict && blocking.Count > 0;
            if (refuse)
            {
                refusalMessage = $"{blocking.Count} conditional block issue(s), refusing before "
                    + $"merging any: {string.Join("; ", blocking.Select(i => i.Message))}. Call "
                    + "MergeConditionalWithReport to execute anyway.";
            }
            else
            {
                var execute = CopyWithDefault(conditions, missing, defaultValue: false);
                OfficeIMOMailMerge.ExecuteConditionalBlocks(document, execute, removeMarkers: true);
                document.Save(result);
                result.Position = 0;
            }
        }
        catch (Exception ex)
        {
            result.Dispose();
            throw new DocumentConversionException(ConditionalFailure, ex);
        }

        if (refuse)
        {
            result.Dispose();
            throw new DocumentConversionException(refusalMessage!);
        }

        report = new DocxMailMergeBlockReport(missing, allIssues);
        return result;
    }

    /// <summary>
    /// A mutable copy of <paramref name="values"/> with <paramref name="defaultValue"/> filled in
    /// for every name in <paramref name="missingNames"/> — what lets the lenient overloads call the
    /// underlying engine, which has no tolerance of its own for a dictionary missing an entry.
    /// </summary>
    private static Dictionary<string, bool> CopyWithDefault(
        IReadOnlyDictionary<string, bool> values, IReadOnlyList<string> missingNames, bool defaultValue)
    {
        var copy = new Dictionary<string, bool>(values.Count + missingNames.Count);
        foreach (KeyValuePair<string, bool> pair in values)
            copy[pair.Key] = pair.Value;
        foreach (string name in missingNames)
            copy[name] = defaultValue;
        return copy;
    }

    /// <summary>
    /// The one implementation behind all four <c>MergeRepeating*</c> overloads. Same shape as
    /// <see cref="MergeConditionalCore"/> — see its remarks for why the preflight and the
    /// try/catch are structured the way they are.
    /// </summary>
    private static MemoryStream MergeRepeatingCore(
        Stream source, IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> regions,
        bool strict, out DocxMailMergeBlockReport report)
    {
        // Outside the try, so a null value reaches the caller as the ArgumentException the class
        // guarantees rather than being re-wrapped as a read-or-write failure.
        Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> records =
            MaterialiseRecords(regions, nameof(regions));

        var result = new MemoryStream();
        List<string> missing;
        List<DocxMailMergeIssue> allIssues;
        bool refuse;
        string? refusalMessage = null;

        try
        {
            using var document = OfficeIMOWordDocument.Load(source);

            var inspection = OfficeIMOMailMerge.InspectTemplate(document, null!, null!, records.Keys);
            allIssues = [.. inspection.Issues.Select(Issue)];
            missing = [.. allIssues
                .Where(i => i.Kind == DocxMailMergeIssueKind.MissingRepeatingBlockData)
                .Select(i => i.Name)];

            List<DocxMailMergeIssue> blocking =
                [.. allIssues.Where(i => RepeatingIssueKinds.Contains(i.Kind))];

            refuse = strict && blocking.Count > 0;
            if (refuse)
            {
                refusalMessage = $"{blocking.Count} repeating block issue(s), refusing before "
                    + $"merging any: {string.Join("; ", blocking.Select(i => i.Message))}. Call "
                    + "MergeRepeatingWithReport to execute anyway.";
            }
            else
            {
                var padded = CopyRegionsWithDefault(records, missing);
                var forEngine = ToEngineRecords(padded);
                OfficeIMOMailMerge.ExecuteRepeatingBlocks(document, forEngine, removeFields: true);
                document.Save(result);
                result.Position = 0;
            }
        }
        catch (Exception ex)
        {
            result.Dispose();
            throw new DocumentConversionException(RepeatingFailure, ex);
        }

        if (refuse)
        {
            result.Dispose();
            throw new DocumentConversionException(refusalMessage!);
        }

        report = new DocxMailMergeBlockReport(missing, allIssues);
        return result;
    }

    /// <summary>
    /// A mutable copy of <paramref name="regions"/> with an empty sequence filled in for every name
    /// in <paramref name="missingNames"/> — the repeating-region twin of <see cref="CopyWithDefault"/>.
    /// </summary>
    private static Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> CopyRegionsWithDefault(
        IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> regions,
        IReadOnlyList<string> missingNames)
    {
        var copy = new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>(
            regions.Count + missingNames.Count);
        foreach (KeyValuePair<string, IEnumerable<IReadOnlyDictionary<string, string>>> pair in regions)
            copy[pair.Key] = pair.Value;
        foreach (string name in missingNames)
            copy[name] = Array.Empty<IReadOnlyDictionary<string, string>>();
        return copy;
    }

    /// <summary>
    /// Converts the padded, read-only regions dictionary to the mutable shape
    /// <see cref="OfficeIMOMailMerge.ExecuteRepeatingBlocks"/> actually takes —
    /// <c>IDictionary&lt;string, IEnumerable&lt;IDictionary&lt;string, string&gt;&gt;&gt;</c>, one
    /// level less read-only than this method's own public parameter. Reuses the existing
    /// <see cref="Copy(IReadOnlyDictionary{string, string})"/> helper per record, matching how the
    /// field-level <see cref="MergeCore"/> already converts a read-only dictionary to what the
    /// engine wants.
    /// </summary>
    private static Dictionary<string, IEnumerable<IDictionary<string, string>>> ToEngineRecords(
        IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> regions)
    {
        var result = new Dictionary<string, IEnumerable<IDictionary<string, string>>>(regions.Count);
        foreach (KeyValuePair<string, IEnumerable<IReadOnlyDictionary<string, string>>> pair in regions)
            result[pair.Key] = pair.Value.Select(Copy).Cast<IDictionary<string, string>>().ToList();
        return result;
    }

    /// <summary>
    /// The one implementation behind all four <c>MergeRepeatingRegions*</c> overloads. Same shape
    /// as <see cref="MergeRepeatingCore"/> — see <see cref="MergeConditionalCore"/>'s remarks for
    /// why the preflight and the try/catch are structured the way they are — except for the two
    /// things nesting genuinely changes, both measured against OfficeIMO.Word 3.2.6 before this
    /// method was written.
    /// </summary>
    /// <remarks>
    /// <b>The supplied-name comparison walks every nesting level</b>, via
    /// <see cref="CollectNamesRecursively"/>, rather than reading only the top-level keys.
    /// <c>InspectTemplate</c>'s <c>RepeatingBlockNames</c> is flat regardless of depth — a template
    /// with <c>{{#each Orders}}</c> around <c>{{#each Lines}}</c> reports <c>[Lines, Orders]</c>
    /// with no indication either is inside the other — so comparing it against top-level keys alone
    /// would report <c>Lines</c> missing on a perfectly correct nested call.
    ///
    /// <b>The lenient path's padding is RECURSIVE, and this is where this method deliberately
    /// diverges from <see cref="MergeRepeatingCore"/>.</b> The flat case pads a missing name into
    /// the one top-level dictionary and that is enough. Here it is not: measured, the engine
    /// resolves a nested marker from its ENCLOSING block row's own regions, so a top-level pad
    /// leaves <c>InvalidOperationException: Repeating block 'Lines' was not supplied.</c> thrown
    /// exactly as if nothing had been padded at all. <see cref="ToEngineRegions"/> therefore fills
    /// every name the template asks for into every level of the tree it builds, and it pads with
    /// <c>RepeatingBlockNames</c> rather than with <c>missing</c> for the second measured reason:
    /// a name supplied by one block row and not another does not appear in <c>missing</c> at all
    /// (the recursive collection above flattens per-row supply into one set) yet throws that same
    /// exception for the row that omitted it. Padding by name-the-template-asks-for covers both.
    /// An extra region key the markers in scope do not use is measured to be inert, which is what
    /// makes the broader pad safe.
    ///
    /// <b>The strict path pads nothing</b>, which needs no special case: it only reaches
    /// <c>Execute</c> when no REPEATING-BLOCK issue is reported — see
    /// <see cref="RepeatingIssueKinds"/> for why the refusal is decided by issue kind rather than
    /// by the inspection's own <c>IsValid</c> — and it hands the caller's own tree straight through
    /// so a per-row omission surfaces rather than being silently defaulted.
    /// </remarks>
    private static MemoryStream MergeRepeatingRegionsCore(
        Stream source, IReadOnlyDictionary<string, IEnumerable<DocxMailMergeBlockData>> regions,
        bool strict, out DocxMailMergeBlockReport report)
    {
        // Outside the try, and BEFORE either walker below runs — see MaterialiseBlockData for the
        // double-enumeration bug that placement closes, and for why the null-value check shares
        // this pass.
        Dictionary<string, IEnumerable<DocxMailMergeBlockData>> rows =
            MaterialiseBlockData(regions, nameof(regions), path: null);

        var result = new MemoryStream();
        List<string> missing;
        List<DocxMailMergeIssue> allIssues;
        bool refuse;
        string? refusalMessage = null;

        try
        {
            using var document = OfficeIMOWordDocument.Load(source);

            var suppliedNames = new HashSet<string>();
            CollectNamesRecursively(rows, suppliedNames);

            var inspection = OfficeIMOMailMerge.InspectTemplate(document, null!, null!, suppliedNames);
            allIssues = [.. inspection.Issues.Select(Issue)];
            missing = [.. allIssues
                .Where(i => i.Kind == DocxMailMergeIssueKind.MissingRepeatingBlockData)
                .Select(i => i.Name)];

            List<DocxMailMergeIssue> blocking =
                [.. allIssues.Where(i => RepeatingIssueKinds.Contains(i.Kind))];

            refuse = strict && blocking.Count > 0;
            if (refuse)
            {
                refusalMessage = $"{blocking.Count} repeating region issue(s), refusing before "
                    + $"merging any: {string.Join("; ", blocking.Select(i => i.Message))}. Call "
                    + "MergeRepeatingRegionsWithReport to execute anyway.";
            }
            else
            {
                List<string> pad = strict ? [] : [.. inspection.RepeatingBlockNames];
                var forEngine = ToEngineRegions(rows, pad);
                OfficeIMOMailMerge.ExecuteRepeatingBlockRegions(document, forEngine, removeFields: true);
                document.Save(result);
                result.Position = 0;
            }
        }
        catch (Exception ex)
        {
            result.Dispose();
            throw new DocumentConversionException(RepeatingRegionsFailure, ex);
        }

        if (refuse)
        {
            result.Dispose();
            throw new DocumentConversionException(refusalMessage!);
        }

        report = new DocxMailMergeBlockReport(missing, allIssues);
        return result;
    }

    /// <summary>
    /// Walks <paramref name="regions"/> and every nested <see cref="DocxMailMergeBlockData.Regions"/>
    /// inside it, adding every region name found at any depth to <paramref name="names"/>.
    /// </summary>
    private static void CollectNamesRecursively(
        IReadOnlyDictionary<string, IEnumerable<DocxMailMergeBlockData>> regions, HashSet<string> names)
    {
        foreach (KeyValuePair<string, IEnumerable<DocxMailMergeBlockData>> pair in regions)
        {
            names.Add(pair.Key);
            foreach (DocxMailMergeBlockData row in pair.Value)
            {
                if (row.Regions is not null)
                    CollectNamesRecursively(row.Regions, names);
            }
        }
    }

    /// <summary>
    /// Converts DocToolkit's own <see cref="DocxMailMergeBlockData"/> tree to OfficeIMO's
    /// <c>WordMailMergeBlockData</c> tree the engine actually takes — recursively, since a block
    /// row can carry further nested regions — filling in an empty sequence for every name in
    /// <paramref name="padNames"/> that a level does not already supply.
    /// </summary>
    /// <remarks>
    /// Padding happens HERE rather than in a <see cref="CopyRegionsWithDefault"/>-style pass over
    /// the caller's own dictionary, for two reasons. It has to reach every nesting level — see
    /// <see cref="MergeRepeatingRegionsCore"/> for the measurement that forced that — and this is
    /// already the pass that rebuilds the whole tree, so padding on the way through cannot mutate
    /// what the caller handed over. <see cref="DocxMailMergeBlockData"/> is immutable anyway; a
    /// separate padding pass would have had to clone every row to say the same thing.
    /// </remarks>
    private static Dictionary<string, IEnumerable<OfficeIMOBlockData>> ToEngineRegions(
        IReadOnlyDictionary<string, IEnumerable<DocxMailMergeBlockData>> regions,
        IReadOnlyList<string> padNames)
    {
        var result = new Dictionary<string, IEnumerable<OfficeIMOBlockData>>(
            regions.Count + padNames.Count);
        foreach (KeyValuePair<string, IEnumerable<DocxMailMergeBlockData>> pair in regions)
            result[pair.Key] = pair.Value.Select(row => ToEngineBlockData(row, padNames)).ToList();

        foreach (string name in padNames)
        {
            if (!result.ContainsKey(name))
                result[name] = Array.Empty<OfficeIMOBlockData>();
        }

        return result;
    }

    private static OfficeIMOBlockData ToEngineBlockData(
        DocxMailMergeBlockData data, IReadOnlyList<string> padNames)
    {
        // Nothing to nest and nothing to pad: use the values-only form, so a row the caller built
        // without regions reaches the engine as one.
        if (data.Regions is null && padNames.Count == 0)
            return new OfficeIMOBlockData(Copy(data.Values));

        return new OfficeIMOBlockData(
            Copy(data.Values), ToEngineRegions(data.Regions ?? NoNestedRegions, padNames));
    }

    private static MemoryStream MergeTableRowsCore(
        Stream source, int tableIndex, int templateRowIndex,
        IEnumerable<IReadOnlyDictionary<string, string>> rows)
    {
        // Outside the try, so a null value reaches the caller as the ArgumentException the class
        // guarantees rather than being re-wrapped as a read-or-write failure.
        List<IDictionary<string, string>> forEngine = ToEngineRows(rows, nameof(rows));

        var result = new MemoryStream();
        try
        {
            using var document = OfficeIMOWordDocument.Load(source);

            var table = document.Tables[tableIndex];
            OfficeIMOMailMerge.ExecuteTableRows(table, templateRowIndex, forEngine, removeFields: true);

            document.Save(result);
            result.Position = 0;
        }
        catch (Exception ex)
        {
            result.Dispose();
            throw new DocumentConversionException(TableRowsFailure, ex);
        }

        return result;
    }

    private static MemoryStream MergeTableRowGroupsCore(
        Stream source, int tableIndex, int groupTemplateRowIndex, int detailTemplateRowIndex,
        IEnumerable<DocxMailMergeTableRowGroup> groups)
    {
        // Outside the try, so a null value reaches the caller as the ArgumentException the class
        // guarantees rather than being re-wrapped as a read-or-write failure.
        List<OfficeIMOTableRowGroup> forEngine = ToEngineGroups(groups, nameof(groups));

        var result = new MemoryStream();
        try
        {
            using var document = OfficeIMOWordDocument.Load(source);

            var table = document.Tables[tableIndex];
            OfficeIMOMailMerge.ExecuteTableRowGroups(
                table, groupTemplateRowIndex, detailTemplateRowIndex, forEngine, removeFields: true);

            document.Save(result);
            result.Position = 0;
        }
        catch (Exception ex)
        {
            result.Dispose();
            throw new DocumentConversionException(TableRowGroupsFailure, ex);
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
    /// Deliberately a `yield return` iterator, and deliberately PRIVATE: every public caller —
    /// <see cref="MergeBatch"/>, <see cref="MergeBatchAsync"/>,
    /// <see cref="MergeBatchWithReport(byte[], IEnumerable{IReadOnlyDictionary{string, string}})"/>
    /// and <see cref="MergeBatchWithReportAsync"/> — is expected to validate its own arguments
    /// before any document is produced. <see cref="MergeBatch"/> and
    /// <see cref="MergeBatchWithReport(byte[], IEnumerable{IReadOnlyDictionary{string, string}})"/>
    /// do this eagerly, each in an ordinary non-iterator method body, before calling this.
    /// <see cref="MergeBatchAsync"/> and <see cref="MergeBatchWithReportAsync"/> cannot do the
    /// same — both are `async IAsyncEnumerable` methods themselves, so their whole body (including
    /// their own argument validation) is deferred until the caller starts enumerating, which is
    /// inherent to async iterators in C#, not a gap. A `yield return` in this PRIVATE method would
    /// defer validation the same way, which is why it stays out of the public surface instead.
    /// </remarks>
    private static IEnumerable<DocxMailMergeBatchItem> MergeBatchCore(
        byte[] docx, IEnumerable<IReadOnlyDictionary<string, string>> records, bool strict)
    {
        var index = 0;
        foreach (var record in records)
        {
            RequireValues(record, nameof(records));

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

            var document = result.ToArray();
            result.Dispose();
            yield return new DocxMailMergeBatchItem(index, document, report);

            index++;
        }
    }

    /// <summary>
    /// The one loop behind every batch overload that writes to disk — strict and lenient alike —
    /// so they can never drift apart. Reuses <see cref="MergeCore"/> exactly like
    /// <see cref="MergeBatchCore"/> does; the only difference is where a result ends up.
    /// </summary>
    /// <remarks>
    /// <paramref name="records"/> is an <see cref="IReadOnlyList{T}"/> here, not the
    /// <see cref="IEnumerable{T}"/> the public methods take — every output path has to be computed
    /// and checked for collisions before any record is merged, which needs indexed, repeatable
    /// access. The public callers materialise the caller's sequence once, up front, before calling
    /// this.
    ///
    /// <b>Writes are deliberately synchronous</b> — <c>File.WriteAllBytes</c>, not
    /// <c>File.WriteAllBytesAsync</c> — because the per-record merge this loop performs is the
    /// dominant cost and is itself synchronous; making only the write awaitable would move a
    /// negligible fraction of the work off the calling thread. Cancellation is still observed
    /// before each record's merge, independent of the write's synchrony.
    /// </remarks>
    private static List<DocxMailMergeFileBatchItem> MergeBatchToFilesCore(
        byte[] docx, IReadOnlyList<IReadOnlyDictionary<string, string>> records,
        Func<int, IReadOnlyDictionary<string, string>, string> outputPathFactory, bool strict,
        CancellationToken ct)
    {
        var paths = new string[records.Count];
        for (var i = 0; i < records.Count; i++)
        {
            RequireValues(records[i], nameof(records));

            var path = outputPathFactory(i, records[i]);
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException(
                    $"Record {i}: outputPathFactory returned a null or blank path.",
                    nameof(outputPathFactory));
            }

            paths[i] = path;
        }

        CheckNoPathCollisions(paths);

        var items = new List<DocxMailMergeFileBatchItem>(records.Count);
        for (var i = 0; i < records.Count; i++)
        {
            ct.ThrowIfCancellationRequested();

            using var source = new MemoryStream(docx, writable: false);
            DocxMailMergeReport report;
            MemoryStream result;
            try
            {
                result = MergeCore(source, records[i], strict, out report);
            }
            catch (DocumentConversionException ex) when (strict)
            {
                throw new DocumentConversionException($"Record {i}: {ex.Message}", ex);
            }

            using (result)
            {
                File.WriteAllBytes(paths[i], result.ToArray());
            }

            items.Add(new DocxMailMergeFileBatchItem(i, paths[i], report));
        }

        return items;
    }

    /// <summary>
    /// Refuses if any two entries in <paramref name="paths"/> are identical, naming both positions
    /// and the path they share.
    /// </summary>
    /// <remarks>
    /// Ordinal comparison, deliberately — no <see cref="Path.GetFullPath(string)"/> normalisation,
    /// no case-insensitivity. Normalising would mean guessing whether two differently-spelled paths
    /// the caller's own factory produced were meant to collide; comparing exactly what the factory
    /// returned is the one behaviour that cannot be wrong about the caller's intent.
    ///
    /// <b>A path matching <c>templatePath</c> itself is not treated as a collision</b> —
    /// <c>outputPathFactory</c> returning the template's own path will overwrite it with a merged
    /// record's output, with no warning. This only compares output paths against each other.
    /// </remarks>
    private static void CheckNoPathCollisions(IReadOnlyList<string> paths)
    {
        var seen = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var i = 0; i < paths.Count; i++)
        {
            if (seen.TryGetValue(paths[i], out var firstIndex))
            {
                throw new ArgumentException(
                    $"Records {firstIndex} and {i} both produced the output path '{paths[i]}'. "
                    + "This refuses rather than silently overwriting one record's document with "
                    + "another's — give outputPathFactory a way to tell every record apart.",
                    "outputPathFactory");
            }

            seen[paths[i]] = i;
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
                [.. inspection.ConditionalBlockNames],
                [.. inspection.RepeatingBlockNames],
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
    /// Eleven of twelve kinds mapped by name; only <c>MissingMergeFieldValue</c> still collapses to
    /// <see cref="DocxMailMergeIssueKind.Other"/> — see <see cref="DocxMailMergeIssueKind"/> for why.
    /// </summary>
    private static DocxMailMergeIssueKind Kind(OfficeIMOIssueKind kind) => kind switch
    {
        OfficeIMOIssueKind.MalformedMergeField => DocxMailMergeIssueKind.MalformedField,
        OfficeIMOIssueKind.UnsupportedMergeFieldFormatting => DocxMailMergeIssueKind.UnsupportedFormatting,
        OfficeIMOIssueKind.MissingConditionalValue => DocxMailMergeIssueKind.MissingConditionalValue,
        OfficeIMOIssueKind.UnmatchedConditionalStart => DocxMailMergeIssueKind.UnmatchedConditionalStart,
        OfficeIMOIssueKind.UnmatchedConditionalEnd => DocxMailMergeIssueKind.UnmatchedConditionalEnd,
        OfficeIMOIssueKind.MismatchedConditionalEnd => DocxMailMergeIssueKind.MismatchedConditionalEnd,
        OfficeIMOIssueKind.MissingRepeatingBlockData => DocxMailMergeIssueKind.MissingRepeatingBlockData,
        OfficeIMOIssueKind.UnmatchedRepeatingBlockStart => DocxMailMergeIssueKind.UnmatchedRepeatingBlockStart,
        OfficeIMOIssueKind.UnmatchedRepeatingBlockEnd => DocxMailMergeIssueKind.UnmatchedRepeatingBlockEnd,
        OfficeIMOIssueKind.MismatchedRepeatingBlockEnd => DocxMailMergeIssueKind.MismatchedRepeatingBlockEnd,
        OfficeIMOIssueKind.UnsupportedMailMergeControlField => DocxMailMergeIssueKind.UnsupportedMailMergeControlField,
        // MissingMergeFieldValue, and any future OfficeIMO addition, stays Other.
        _ => DocxMailMergeIssueKind.Other,
    };

    private static string Text(string? value) => value ?? string.Empty;
}
