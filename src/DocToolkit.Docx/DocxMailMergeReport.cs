namespace DocToolkit;

/// <summary>What a mail-merge template asks for, read without merging anything.</summary>
/// <remarks>
/// The question this answers is <i>"what does this template want, and is it sound?"</i> — which is
/// the first thing anyone integrating a template somebody else authored needs, and the only way to
/// tell a template with no merge fields apart from one whose fields are all about to go unfilled.
/// A document with no merge fields is <b>not</b> an error: it reports no field names and
/// <see cref="IsValid"/> true, because there is nothing wrong with it.
/// </remarks>
public sealed class DocxMailMergeTemplate
{
    internal DocxMailMergeTemplate(
        IReadOnlyList<string> fieldNames, IReadOnlyList<string> conditionalBlockNames,
        IReadOnlyList<string> repeatingBlockNames, IReadOnlyList<DocxMailMergeIssue> issues,
        bool isValid)
    {
        FieldNames = fieldNames;
        ConditionalBlockNames = conditionalBlockNames;
        RepeatingBlockNames = repeatingBlockNames;
        Issues = issues;
        IsValid = isValid;
    }

    /// <summary>
    /// The distinct merge-field names this template asks for.
    /// </summary>
    /// <remarks>
    /// Fields that are not <c>MERGEFIELD</c>s are not listed — measured with a <c>DATE</c> field,
    /// which is correctly ignored. A field carrying a formatting switch, such as
    /// <c>\# "#,##0.00"</c>, <b>is</b> listed under its own name.
    /// </remarks>
    public IReadOnlyList<string> FieldNames { get; }

    /// <summary>
    /// The distinct conditional-block names this template asks for — the <c>Name</c> in every
    /// <c>{{#Name}}</c> marker pair. See MergeConditional.
    /// </summary>
    public IReadOnlyList<string> ConditionalBlockNames { get; }

    /// <summary>
    /// The distinct repeating-block names this template asks for — the <c>Name</c> in every
    /// <c>{{#each Name}}</c> marker pair. See MergeRepeating.
    /// </summary>
    /// <remarks>
    /// <b>Flat across nesting depth.</b> A marker nested inside another appears in this list
    /// exactly like a top-level one, with nothing indicating which marker it is nested inside —
    /// measured. A caller reasoning about MergeRepeatingRegions's nested shape from this list
    /// alone cannot tell nesting apart from two independent regions.
    /// </remarks>
    public IReadOnlyList<string> RepeatingBlockNames { get; }

    /// <summary>Everything wrong with the template, in document order. Empty when there is nothing.</summary>
    public IReadOnlyList<DocxMailMergeIssue> Issues { get; }

    /// <summary>
    /// Whether the template is sound. False means <see cref="Issues"/> is non-empty.
    /// </summary>
    /// <remarks>
    /// <b>This says nothing about whether a merge will be complete</b> — that depends on the values
    /// supplied, and is what <see cref="DocxMailMergeReport.IsComplete"/> answers.
    /// </remarks>
    public bool IsValid { get; }
}

/// <summary>Something wrong with a mail-merge template.</summary>
public sealed class DocxMailMergeIssue
{
    internal DocxMailMergeIssue(string name, string message, DocxMailMergeIssueKind kind)
    {
        Name = name;
        Message = message;
        Kind = kind;
    }

    /// <summary>The field the issue concerns. Empty when the issue names none.</summary>
    public string Name { get; }

    /// <summary>What is wrong, in the underlying library's words.</summary>
    public string Message { get; }

    /// <summary>Which kind of problem this is.</summary>
    public DocxMailMergeIssueKind Kind { get; }
}

/// <summary>The kinds of template problem this API distinguishes.</summary>
/// <remarks>
/// <b>The underlying library reports twelve kinds; this names eleven of them and collapses one.</b>
/// <see cref="Other"/> now covers only <c>MissingMergeFieldValue</c> — a field-level gap unrelated
/// to conditional blocks or repeating regions, and out of scope for the ticket that named the other
/// nine. Before conditional-block and repeating-block execution shipped, all nine of those arrived
/// as <see cref="Other"/> too, because surfacing them by name would have advertised capabilities a
/// caller could not reach — see MergeConditional and MergeRepeating, which is why that is no
/// longer true.
/// </remarks>
public enum DocxMailMergeIssueKind
{
    /// <summary>
    /// A problem this API does not distinguish by its own kind — currently only a MERGEFIELD
    /// missing its value in a supplied-names inspection. Read <see cref="DocxMailMergeIssue.Message"/>.
    /// </summary>
    Other = 0,

    /// <summary>
    /// The field is not a usable <c>MERGEFIELD</c> — measured with an instruction carrying no field
    /// name, which reports <i>"A MERGEFIELD instruction must contain exactly one field name."</i>
    /// </summary>
    MalformedField = 1,

    /// <summary>The field asks for formatting this engine does not apply.</summary>
    UnsupportedFormatting = 2,

    /// <summary>A conditional block (<c>{{#Name}}</c>) was found without a supplied condition.</summary>
    MissingConditionalValue = 3,

    /// <summary>A conditional block's start marker (<c>{{#Name}}</c>) has no matching end marker.</summary>
    UnmatchedConditionalStart = 4,

    /// <summary>A conditional block's end marker (<c>{{/Name}}</c>) has no matching start marker.</summary>
    UnmatchedConditionalEnd = 5,

    /// <summary>A conditional block's end marker closed a different block name than the current start marker.</summary>
    MismatchedConditionalEnd = 6,

    /// <summary>A repeating block (<c>{{#each Name}}</c>) was found without supplied rows.</summary>
    MissingRepeatingBlockData = 7,

    /// <summary>A repeating block's start marker (<c>{{#each Name}}</c>) has no matching end marker.</summary>
    UnmatchedRepeatingBlockStart = 8,

    /// <summary>A repeating block's end marker (<c>{{/each Name}}</c>) has no matching start marker.</summary>
    UnmatchedRepeatingBlockEnd = 9,

    /// <summary>A repeating block's end marker closed a different block name than the current start marker.</summary>
    MismatchedRepeatingBlockEnd = 10,

    /// <summary>A Word-native mail-merge control field was found that this engine does not execute.</summary>
    UnsupportedMailMergeControlField = 11,
}

/// <summary>A merged document, together with what happened to every field in it.</summary>
/// <remarks>
/// Returned by the lenient overload only. <see cref="DocxMailMerge.Merge(byte[], IReadOnlyDictionary{string, string})"/>
/// returns bytes and refuses to produce a document at all when a field went unfilled, so it has
/// nothing to report.
/// </remarks>
public sealed class DocxMailMergeResult
{
    internal DocxMailMergeResult(byte[] document, DocxMailMergeReport report)
    {
        Document = document;
        Report = report;
    }

    /// <summary>
    /// The merged document — <b>complete or not</b>. When
    /// <see cref="DocxMailMergeReport.IsComplete"/> is false this still opens cleanly and looks
    /// finished; the unfilled fields show their placeholder text.
    /// </summary>
    public byte[] Document { get; }

    /// <summary>What happened to every field.</summary>
    public DocxMailMergeReport Report { get; }
}

/// <summary>One record's document, from a batch call, together with what happened to every field
/// in it.</summary>
/// <remarks>
/// Returned by the lenient batch overloads only — <see cref="DocxMailMerge.MergeBatch(byte[], System.Collections.Generic.IEnumerable{System.Collections.Generic.IReadOnlyDictionary{string, string}})"/>
/// refuses to produce a document for an incomplete record at all, so there is nothing to report for
/// that record.
/// </remarks>
public sealed class DocxMailMergeBatchItem
{
    internal DocxMailMergeBatchItem(int recordIndex, byte[] document, DocxMailMergeReport report)
    {
        RecordIndex = recordIndex;
        Document = document;
        Report = report;
    }

    /// <summary>The record's position in the sequence passed in, starting at 0.</summary>
    public int RecordIndex { get; }

    /// <summary>
    /// The merged document — <b>complete or not</b>. When <see cref="Report"/>'s
    /// <see cref="DocxMailMergeReport.IsComplete"/> is false this still opens cleanly and looks
    /// finished; the unfilled fields show their placeholder text.
    /// </summary>
    public byte[] Document { get; }

    /// <summary>What happened to every field in this record's document.</summary>
    public DocxMailMergeReport Report { get; }
}

/// <summary>One record's output file, from a file-path batch call, together with what happened to
/// every field in it.</summary>
/// <remarks>
/// Returned by the lenient file-path batch forms only —
/// <see cref="DocxMailMerge.MergeBatchToFilesWithReport(string, System.Collections.Generic.IEnumerable{System.Collections.Generic.IReadOnlyDictionary{string, string}}, System.Func{int, System.Collections.Generic.IReadOnlyDictionary{string, string}, string})"/>
/// and its async twin — the strict forms refuse to write a file for an incomplete record at all,
/// so there is nothing to report for that record.
/// </remarks>
public sealed class DocxMailMergeFileBatchItem
{
    internal DocxMailMergeFileBatchItem(int recordIndex, string outputPath, DocxMailMergeReport report)
    {
        RecordIndex = recordIndex;
        OutputPath = outputPath;
        Report = report;
    }

    /// <summary>The record's position in the sequence passed in, starting at 0.</summary>
    public int RecordIndex { get; }

    /// <summary>Where this record's document was written — exactly what <c>outputPathFactory</c>
    /// returned for this record.</summary>
    public string OutputPath { get; }

    /// <summary>What happened to every field in this record's document.</summary>
    public DocxMailMergeReport Report { get; }
}

/// <summary>What happened to every merge field in a document.</summary>
public sealed class DocxMailMergeReport
{
    internal DocxMailMergeReport(
        IReadOnlyList<DocxMailMergeField> fields, IReadOnlyList<string> missingFieldNames,
        int mergedCount, bool isComplete)
    {
        Fields = fields;
        MissingFieldNames = missingFieldNames;
        MergedCount = mergedCount;
        IsComplete = isComplete;
    }

    /// <summary>
    /// Every merge field the template contained, in document order — <b>one entry per occurrence,
    /// not per name</b>. A name used twice in a letter, in the salutation and again in the body,
    /// appears twice. Measured.
    /// </summary>
    public IReadOnlyList<DocxMailMergeField> Fields { get; }

    /// <summary>
    /// The distinct names no value was supplied for. Empty when <see cref="IsComplete"/> is true.
    /// </summary>
    public IReadOnlyList<string> MissingFieldNames { get; }

    /// <summary>How many field occurrences received a value.</summary>
    public int MergedCount { get; }

    /// <summary>
    /// Whether every field in the template received a value.
    /// </summary>
    /// <remarks>
    /// <b>True does not mean the document reads correctly.</b> A supplied value that is an empty
    /// string merges and counts as complete, which is the honest report — the engine was given a
    /// value and used it. See <see cref="DocxMailMerge"/> for why a <c>null</c> value is refused
    /// rather than treated as one of these.
    /// </remarks>
    public bool IsComplete { get; }
}

/// <summary>What happened to one merge field.</summary>
public sealed class DocxMailMergeField
{
    internal DocxMailMergeField(
        string name, DocxMailMergeFieldStatus status, string? value, string message)
    {
        Name = name;
        Status = status;
        Value = value;
        Message = message;
    }

    /// <summary>The field's name, as the template spells it.</summary>
    public string Name { get; }

    /// <summary>What happened to it.</summary>
    public DocxMailMergeFieldStatus Status { get; }

    /// <summary>
    /// The value written into the document, or <see langword="null"/> when nothing was written.
    /// </summary>
    public string? Value { get; }

    /// <summary>What happened, in the underlying library's words.</summary>
    public string Message { get; }
}

/// <summary>What happened to a single merge field.</summary>
public enum DocxMailMergeFieldStatus
{
    /// <summary>A value was supplied and written into the document.</summary>
    Merged = 0,

    /// <summary>
    /// No value was supplied. The field keeps its placeholder text, so the document still reads
    /// <c>«FieldName»</c> where the value should be — measured.
    /// </summary>
    MissingValue = 1,

    /// <summary>The field asked for formatting this engine does not apply.</summary>
    UnsupportedFormatting = 2,

    /// <summary>The field's instruction could not be read as a <c>MERGEFIELD</c>.</summary>
    Malformed = 3,
}

/// <summary>One repeated block row, for the nested form of repeating-region merging.</summary>
/// <remarks>
/// Mirrors <c>OfficeIMO.Word.WordMailMergeBlockData</c> without exposing it — no
/// <c>OfficeIMO.*</c> type is ever public in this library. See MergeRepeatingRegions.
/// </remarks>
public sealed class DocxMailMergeBlockData
{
    /// <summary>Creates a repeated block row with no nested regions.</summary>
    /// <param name="values">Values applied to merge fields inside this block row.</param>
    public DocxMailMergeBlockData(IReadOnlyDictionary<string, string> values)
        : this(values, regions: null)
    {
    }

    /// <summary>Creates a repeated block row with nested repeated regions.</summary>
    /// <param name="values">Values applied to merge fields inside this block row.</param>
    /// <param name="regions">
    /// Nested repeated regions available inside this block row, keyed by marker name. Null when
    /// this row has none.
    /// </param>
    public DocxMailMergeBlockData(
        IReadOnlyDictionary<string, string> values,
        IReadOnlyDictionary<string, IEnumerable<DocxMailMergeBlockData>>? regions)
    {
        ArgumentNullException.ThrowIfNull(values);
        Values = values;
        Regions = regions;
    }

    /// <summary>Values applied to merge fields inside this block row.</summary>
    public IReadOnlyDictionary<string, string> Values { get; }

    /// <summary>Nested repeated regions available inside this block row, keyed by marker name.</summary>
    public IReadOnlyDictionary<string, IEnumerable<DocxMailMergeBlockData>>? Regions { get; }
}

/// <summary>One grouped table-row mail-merge data set — a group/header row plus its detail rows.</summary>
/// <remarks>
/// Mirrors <c>OfficeIMO.Word.WordMailMergeTableRowGroup</c> without exposing it. See
/// MergeTableRowGroups.
/// </remarks>
public sealed class DocxMailMergeTableRowGroup
{
    /// <summary>Creates a grouped table-row data set.</summary>
    /// <param name="values">Values applied to the group template row.</param>
    /// <param name="rows">Values applied to repeated detail rows inside the group.</param>
    public DocxMailMergeTableRowGroup(
        IReadOnlyDictionary<string, string> values,
        IEnumerable<IReadOnlyDictionary<string, string>> rows)
    {
        ArgumentNullException.ThrowIfNull(values);
        ArgumentNullException.ThrowIfNull(rows);
        Values = values;
        Rows = rows;
    }

    /// <summary>Values applied to the group template row.</summary>
    public IReadOnlyDictionary<string, string> Values { get; }

    /// <summary>Values applied to repeated detail rows inside the group.</summary>
    public IEnumerable<IReadOnlyDictionary<string, string>> Rows { get; }
}

/// <summary>
/// What happened when a template's conditional blocks or repeating regions were resolved —
/// which names the template asked for that the caller did not supply, and any structural problem.
/// </summary>
/// <remarks>
/// Not <see cref="DocxMailMergeReport"/> reused: that type's <see cref="DocxMailMergeReport.Fields"/>
/// and <see cref="DocxMailMergeReport.MissingFieldNames"/> describe individual
/// <c>MERGEFIELD</c>s, the wrong shape for "which condition or region names were missing."
/// </remarks>
public sealed class DocxMailMergeBlockReport
{
    internal DocxMailMergeBlockReport(
        IReadOnlyList<string> missingNames, IReadOnlyList<DocxMailMergeIssue> issues)
    {
        MissingNames = missingNames;
        Issues = issues;
    }

    /// <summary>
    /// The distinct condition or region names the template asked for that the caller did not
    /// supply. Each one received a neutral default before merging — see
    /// MergeConditionalWithReport for what "neutral" means for a condition versus a region.
    /// </summary>
    public IReadOnlyList<string> MissingNames { get; }

    /// <summary>Structural problems found — an unmatched or mismatched marker, for example.</summary>
    public IReadOnlyList<DocxMailMergeIssue> Issues { get; }

    /// <summary>Whether every name the template asked for was supplied. <c>MissingNames.Count == 0</c>.</summary>
    public bool IsComplete => MissingNames.Count == 0;
}

/// <summary>A merged document, together with what happened to its conditional blocks or repeating regions.</summary>
/// <remarks>
/// Returned by the lenient overloads only — MergeConditional and its repeating-block siblings
/// refuse to produce a document when a name goes unsupplied, so they have nothing to report.
/// </remarks>
public sealed class DocxMailMergeBlockResult
{
    internal DocxMailMergeBlockResult(byte[] document, DocxMailMergeBlockReport report)
    {
        Document = document;
        Report = report;
    }

    /// <summary>
    /// The merged document — <b>complete or not</b>. When <see cref="DocxMailMergeBlockReport.IsComplete"/>
    /// is false, every name the template asked for and the caller did not supply received a
    /// neutral default rather than being left unresolved.
    /// </summary>
    public byte[] Document { get; }

    /// <summary>What happened to every condition or region name.</summary>
    public DocxMailMergeBlockReport Report { get; }
}
