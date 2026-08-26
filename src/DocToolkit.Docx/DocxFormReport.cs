namespace DocToolkit;

/// <summary>Which name identifies a content control.</summary>
/// <remarks>
/// A control carries both a tag (a machine name) and an alias (what Word shows in its UI). Measured
/// 2026-08-26, the two return genuinely different keys for the same document — <c>FullName</c>
/// against <c>Full name</c> — which is why this is a parameter rather than a fixed choice.
/// </remarks>
public enum DocxFormKey
{
    /// <summary>The tag, falling back to the alias. The default, because it works either way.</summary>
    TagThenAlias = 0,

    /// <summary>The tag only.</summary>
    Tag = 1,

    /// <summary>The alias only.</summary>
    Alias = 2,

    /// <summary>The alias, falling back to the tag.</summary>
    AliasThenTag = 3,
}

/// <summary>The content controls a document carries, and what is currently in them.</summary>
public sealed class DocxFormReport
{
    internal DocxFormReport(IReadOnlyList<DocxFormField> fields) => Fields = fields;

    /// <summary>Every content control, keyed as the <see cref="DocxFormKey"/> asked for.</summary>
    public IReadOnlyList<DocxFormField> Fields { get; }
}

/// <summary>One content control, and its current content.</summary>
public sealed class DocxFormField
{
    internal DocxFormField(string key, DocxFormValue value)
    {
        Key = key;
        Value = value;
    }

    /// <summary>The name this control answers to.</summary>
    public string Key { get; }

    /// <summary>
    /// What the control holds now — the same type <c>DocxForm.Fill</c> takes, so a round trip
    /// needs no translation.
    /// </summary>
    public DocxFormValue Value { get; }
}

/// <summary>Whether a set of values fits a document's content controls.</summary>
/// <remarks>
/// <b>This checks keys, not values.</b> Measured 2026-08-26: a drop-down value outside its list and
/// a non-date for a date control both came back valid. What it does answer — which controls got no
/// value, which values matched nothing, and which names are ambiguous — is what a caller needs
/// before filling a template somebody else authored.
/// </remarks>
public sealed class DocxFormValidation
{
    internal DocxFormValidation(
        bool isValid, IReadOnlyList<string> expectedKeys, IReadOnlyList<string> suppliedKeys,
        IReadOnlyList<DocxFormIssue> issues)
    {
        IsValid = isValid;
        ExpectedKeys = expectedKeys;
        SuppliedKeys = suppliedKeys;
        Issues = issues;
    }

    /// <summary>True when <see cref="Issues"/> is empty — of <b>any</b> kind.</summary>
    /// <remarks>
    /// Upstream has two flags that suppress whole issue kinds, and so change what "valid" means.
    /// Neither is exposed: a property whose meaning depends on an argument is the drift the
    /// <c>*Core</c> convention exists to prevent. Filter <see cref="Issues"/> by
    /// <see cref="DocxFormIssue.Kind"/> instead — strictly more information than a flag.
    /// </remarks>
    public bool IsValid { get; }

    /// <summary>The keys the document asks for.</summary>
    public IReadOnlyList<string> ExpectedKeys { get; }

    /// <summary>The keys that were supplied.</summary>
    public IReadOnlyList<string> SuppliedKeys { get; }

    /// <summary>Everything wrong, in the order the underlying check reported it.</summary>
    public IReadOnlyList<DocxFormIssue> Issues { get; }
}

/// <summary>One problem with a set of values.</summary>
public sealed class DocxFormIssue
{
    internal DocxFormIssue(string key, DocxFormIssueKind kind, string message)
    {
        Key = key;
        Kind = kind;
        Message = message;
    }

    /// <summary>The control, or the supplied name, the issue concerns.</summary>
    public string Key { get; }

    /// <summary>Which kind of problem this is.</summary>
    public DocxFormIssueKind Kind { get; }

    /// <summary>What is wrong, in the underlying library's words.</summary>
    public string Message { get; }
}

/// <summary>The kinds of problem this API distinguishes.</summary>
/// <remarks>
/// <b>Upstream reports nine kinds; three of them can actually happen.</b> Measured 2026-08-26 by
/// trying to provoke each: a drop-down value outside its list, a string where a date belongs, and a
/// bool where a date belongs all reported valid. So <c>InvalidChoice</c>, <c>InvalidDate</c>,
/// <c>InvalidBoolean</c>, <c>InvalidImage</c>, <c>InvalidRepeatingSection</c> and
/// <c>UnmappedControl</c> arrive as <see cref="Other"/> with their message intact — if one ever does
/// fire the caller still sees it, and nothing here advertises a check that does not run.
///
/// The fixtures behind that measurement were hand-built, so it remains possible those kinds need
/// markup Word writes and a probe did not. That is the reason they map to <see cref="Other"/>
/// rather than being dropped.
/// </remarks>
public enum DocxFormIssueKind
{
    /// <summary>
    /// A problem this API does not distinguish. Read <see cref="DocxFormIssue.Message"/>.
    /// </summary>
    Other = 0,

    /// <summary>A control received no value.</summary>
    MissingValue = 1,

    /// <summary>A value was supplied for a name no control answers to.</summary>
    UnusedValue = 2,

    /// <summary>Two controls answer to the same name, so a value cannot be aimed at one of them.</summary>
    DuplicateKey = 3,
}
