using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
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
/// lenient about a MISSING value</b>, where <see cref="DocxMailMerge.Merge"/> refuses. Two measured
/// differences justify that: a control given no value keeps its <i>own</i> existing text rather than
/// showing an injected marker, and unlike mail merge this class ships a <c>Validate</c> a caller can
/// run first.
///
/// <b>It is NOT lenient about a value that does not fit a typed control, and the three typed kinds
/// do not agree with each other.</b> Measured: a drop-down value outside its list <i>throws</i>,
/// while a non-date for a date picker and a non-boolean for a check box are silently skipped and the
/// control keeps its old content. That asymmetry is the library beneath, not a choice made here —
/// which is the strongest reason to run <c>Validate</c> first, since it reports all three the same
/// way, before anything is written.
///
/// <b>Only the document BODY is read or written.</b> A content control in a header or a footer is
/// invisible to every method here: it is not in <see cref="Inspect(byte[], DocxFormKey)"/>'s report,
/// a value aimed at it is reported as <see cref="DocxFormIssueKind.UnusedValue"/> — which reads as
/// though the caller invented the name — and <c>Fill</c> leaves it untouched. Measured.
/// <see cref="DocxMailMerge"/> <i>does</i> reach headers, so the two template APIs genuinely differ
/// here; if your form lives in a header, that is the one to use.
///
/// <b>Images are supplied as bytes and never as a path</b> — see <see cref="DocxFormValue"/>.
/// </remarks>
public static class DocxForm
{
    /// <summary>
    /// Reads the content controls in <paramref name="docx"/>'s <b>body</b>, and what they hold.
    /// </summary>
    /// <remarks>
    /// <b>Not every control in the document</b> — see <see cref="DocxFormReport.Fields"/> for the
    /// three reasons one can be absent. A document with no content controls reports none rather than
    /// failing, which is how a caller catches having passed the wrong document, since filling one
    /// succeeds and changes nothing.
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
    /// <see cref="Fill(byte[], IReadOnlyDictionary{string, DocxFormValue}, DocxFormKey)"/>, which
    /// will not tell you what it skipped.
    ///
    /// <b>It is not a complete guarantee that <c>Fill</c> will succeed.</b> Measured: image bytes
    /// that are not a readable image validate clean and then throw from <c>Fill</c>, because nothing
    /// here decodes them. A clean result means no <i>key or typed-value</i> problem was found.
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
    /// <exception cref="DocumentConversionException">
    /// The document could not be read or written, <b>or a value did not fit a typed control in a way
    /// the library beneath refuses</b> — measured: a drop-down value outside its list throws here,
    /// while a bad date or boolean is skipped silently. Run
    /// <see cref="Validate(byte[], IReadOnlyDictionary{string, DocxFormValue}, DocxFormKey)"/> first
    /// to see all three before writing.
    /// </exception>
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
    /// <exception cref="InvalidOperationException">
    /// The fill would change a content control the document locks against editing (A119). Nothing
    /// is written, and the document passed in is untouched. Before 0.54.0 this succeeded and left
    /// the lock in place, producing a file that declared the control protected while its content
    /// had been replaced.
    /// </exception>
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

    /// <summary>
    /// Each control's alias, reachable by whichever name <c>ExtractContentControlValues</c> used
    /// as its key.
    /// </summary>
    /// <remarks>
    /// <b>A control is matched when EITHER its tag or its alias equals the key</b>, rather than by
    /// replaying the <see cref="DocxFormKey"/> preference order here. Replaying it would be a
    /// second implementation of the decision OfficeIMO already made, and two implementations of
    /// one rule is the drift <c>SetCellValue</c>, <c>SectionPropertiesFactory</c>,
    /// <c>ValidateSheetName</c> and <c>ContentControls</c> each exist to prevent. Asking "which
    /// control does this name reach?" needs no preference order at all.
    ///
    /// <b>An ambiguous name yields no alias rather than a guess.</b> Word does not require either
    /// name to be unique, and <see cref="DocxFormValidation"/> already reports that as
    /// <see cref="DocxFormIssueKind.DuplicateKey"/> - so the honest answer here is to say nothing
    /// rather than pick one of two controls and be silently wrong half the time.
    /// </remarks>
    private static Dictionary<string, string> AliasesByName(OfficeIMOWordDocument document)
    {
        var byName = new Dictionary<string, List<string>>(StringComparer.Ordinal);

        foreach (var control in document.StructuredDocumentTags)
        {
            var alias = control.Alias ?? string.Empty;
            foreach (var name in new[] { control.Tag, control.Alias })
            {
                if (string.IsNullOrEmpty(name)) continue;
                if (!byName.TryGetValue(name, out var found)) byName[name] = found = [];
                found.Add(alias);
            }
        }

        return byName
            .Where(pair => pair.Value.Distinct(StringComparer.Ordinal).Count() == 1)
            .ToDictionary(pair => pair.Key, pair => pair.Value[0], StringComparer.Ordinal);
    }

    private static DocxFormReport InspectCore(Stream source, DocxFormKey key)
    {
        try
        {
            using var document = OfficeIMOWordDocument.Load(source);

            // The alias comes from OfficeIMO's own control collection rather than from a second
            // walk of the OOXML (A120). Measured 2026-09-03: WordStructuredDocumentTag exposes
            // Tag and Alias together and nothing else useful - no lock, and no route to the
            // element underneath. So the alias is reachable without inventing a reader, and
            // IsLocked is not reachable at all, which is why only half of A120 shipped.
            var aliases = AliasesByName(document);

            DocxFormField[] fields = [.. document
                .ExtractContentControlValues(Upstream(key))
                .Select(pair => new DocxFormField(
                    pair.Key,
                    DocxFormValue.FromUpstream(pair.Value),
                    aliases.TryGetValue(pair.Key, out var alias) ? alias : string.Empty))];

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

            // Derived rather than copied from result.IsValid, which is a second source of truth
            // for the same thing: the doc comment promises "true when Issues is empty, of any
            // kind", and an upstream release that excluded a kind from its own flag would make that
            // promise silently false while nothing failed.
            return new DocxFormValidation(
                result.Issues.Count == 0,
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
            // The locked controls are read BEFORE the fill and compared after (A119). Deciding up
            // front would mean resolving each key to a control here, which is the DocxFormKey
            // decision OfficeIMO already makes - and a second reader keyed on the same names is
            // the drift ContentControls, SetCellValue and ValidateSheetName each exist to prevent.
            // Comparing afterwards asks only "did a locked control change?", which the XML answers
            // on its own.
            var before = LockedContent(source);

            source.Position = 0;
            using (var document = OfficeIMOWordDocument.Load(source))
            {
                document.FillContentControlValues(Upstream(values), Upstream(key));
                document.Save(result);
            }

            result.Position = 0;
            RefuseIfALockedControlChanged(before, LockedContent(result));

            result.Position = 0;
            return result;
        }
        catch (Exception ex) when (ex is not DocumentConversionException and not InvalidOperationException)
        {
            result.Dispose();
            throw new DocumentConversionException(FillFailure, ex);
        }
        catch
        {
            result.Dispose();
            throw;
        }
    }

    /// <summary>
    /// Each locked control's content, keyed by its position among the locked controls.
    /// </summary>
    /// <remarks>
    /// <b>Position rather than tag, deliberately.</b> A tag is optional and need not be unique, so
    /// keying on one would silently compare the wrong pair - and this only has to answer whether
    /// anything changed, which position answers without inventing an identity scheme.
    ///
    /// The stream's position is restored, because the caller reads it next.
    /// </remarks>
    private static List<string> LockedContent(Stream docx)
    {
        var at = docx.Position;
        docx.Position = 0;
        try
        {
            using var document = WordprocessingDocument.Open(docx, false);
            var body = document.MainDocumentPart?.Document?.Body;
            if (body is null) return [];

            return [.. body.Descendants<SdtElement>()
                .Where(c => c.Descendants<DocumentFormat.OpenXml.Wordprocessing.Lock>().Any())
                .Select(c => string.Concat(c.Descendants<Text>().Select(t => t.Text)))];
        }
        finally
        {
            docx.Position = at;
        }
    }

    /// <summary>
    /// Refuses the fill when it changed a control the document declares locked (A119).
    /// </summary>
    /// <remarks>
    /// <b>Measured 2026-09-02: the old behaviour wrote through the lock AND left the lock in
    /// place</b>, so the document came back declaring a control protected while its content had
    /// been replaced. Nothing in the file recorded that, and no caller could detect it.
    ///
    /// A count change is treated as a change too: a fill that added or removed a locked control is
    /// not something this method should be doing either.
    /// </remarks>
    private static void RefuseIfALockedControlChanged(List<string> before, List<string> after)
    {
        if (before.Count == after.Count && before.SequenceEqual(after, StringComparer.Ordinal)) return;

        throw new InvalidOperationException(
            "The fill would change a content control the document locks against editing, so nothing "
            + "was written. A w:lock is the author's instruction about their own document, and "
            + "writing through it produces a file that still declares the control locked while its "
            + "content has changed. Remove the lock in Word, or leave that control out of the values.");
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
