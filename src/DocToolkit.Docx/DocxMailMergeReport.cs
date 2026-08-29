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
        IReadOnlyList<string> fieldNames, IReadOnlyList<DocxMailMergeIssue> issues, bool isValid)
    {
        FieldNames = fieldNames;
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
/// <b>The underlying library reports twelve kinds and this collapses them to three</b>, the same
/// way <see cref="DocxRevisionKind"/> collapses eleven revision types to three. Nine of the twelve
/// describe conditional blocks and repeating-block regions — template features this API does not
/// execute — so surfacing them by name would advertise capabilities a caller cannot reach. They
/// arrive as <see cref="Other"/>, message intact.
/// </remarks>
public enum DocxMailMergeIssueKind
{
    /// <summary>
    /// A problem this API does not distinguish, including every conditional-block and
    /// repeating-block problem. Read <see cref="DocxMailMergeIssue.Message"/>.
    /// </summary>
    Other = 0,

    /// <summary>
    /// The field is not a usable <c>MERGEFIELD</c> — measured with an instruction carrying no field
    /// name, which reports <i>"A MERGEFIELD instruction must contain exactly one field name."</i>
    /// </summary>
    MalformedField = 1,

    /// <summary>The field asks for formatting this engine does not apply.</summary>
    UnsupportedFormatting = 2,
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
/// Returned by the lenient file-path batch overload only — the strict one refuses to write a file
/// for an incomplete record at all, so there is nothing to report for that record.
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
