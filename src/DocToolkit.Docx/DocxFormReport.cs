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

    /// <summary>
    /// The content controls this document exposes under the <see cref="DocxFormKey"/> asked for.
    /// </summary>
    /// <remarks>
    /// <b>Not necessarily every control in the document</b>, for three measured reasons. A control
    /// is absent when it has no name under that key mode; only the first of several controls sharing
    /// a name appears; and <b>a control in a header or a footer is never included</b>, because only
    /// the body is read. Measured: two controls sharing a tag yield one field, a tag-only template
    /// read with <see cref="DocxFormKey.Alias"/> yields <b>none at all</b> — which looks exactly like
    /// a document that has no form — and a header control is missing while
    /// <see cref="DocxMailMerge"/> would have found it.
    ///
    /// <b>The order is not document order.</b> Measured: a form authored FullName, Plan, Start,
    /// Signed comes back Signed, Start, Plan, FullName. Nothing here promises any order, so sort by
    /// <see cref="DocxFormField.Key"/> if you are rendering a form and the sequence matters.
    ///
    /// <see cref="DocxForm.Validate(byte[], IReadOnlyDictionary{string, DocxFormValue}, DocxFormKey)"/>
    /// reports both cases, as <see cref="DocxFormIssueKind.DuplicateKey"/> and
    /// <see cref="DocxFormIssueKind.UnmappedControl"/>.
    /// </remarks>
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
/// <b>This checks keys AND values.</b> It reports which controls got no value, which values matched
/// nothing and which names are ambiguous — and, for a typed control, whether the value fits: a
/// drop-down value outside its list, a non-date for a date picker and a non-boolean for a check box
/// are each reported under their own <see cref="DocxFormIssueKind"/>.
///
/// <b>A plain text control validates anything</b>, because the underlying library has no constraint
/// to check it against. So a clean result means "nothing detectably wrong", not "every value is
/// right".
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

/// <summary>The kinds of problem a set of values can have.</summary>
/// <remarks>
/// <b>All nine of the underlying library's kinds are surfaced, and an earlier draft of this type
/// surfaced three.</b> That draft was written on a measurement which said the other six could not be
/// provoked — and the measurement was wrong: its fixtures were hand-built <c>SdtBlock</c> markup,
/// which is not a typed control, so a drop-down value outside its list really did validate clean
/// because there was no drop-down. Against a form authored the way Word authors one, three of the
/// six fire on the first attempt.
///
/// The lesson is recorded on <c>DocxFormFixtures</c>: a fixture must be authored the way the library
/// under test authors one, or what gets measured is the fixture.
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

    /// <summary>A control carries no name under the <see cref="DocxFormKey"/> in use.</summary>
    /// <remarks>
    /// Reachable two ways: a control with neither a tag nor an alias, and — more easily — reading a
    /// tag-only template with <see cref="DocxFormKey.Alias"/>.
    /// </remarks>
    UnmappedControl = 4,

    /// <summary>The value is not usable as the check box's true/false.</summary>
    InvalidBoolean = 5,

    /// <summary>The value is not usable as the date control's date.</summary>
    InvalidDate = 6,

    /// <summary>The value is not one of the drop-down's options.</summary>
    InvalidChoice = 7,

    /// <summary>The value is not usable as the picture control's image.</summary>
    /// <remarks>
    /// Reachable through this API with a file name that has no extension — the underlying library
    /// requires one.
    /// </remarks>
    InvalidImage = 8,

    /// <summary>The value does not fit the control's repeating section.</summary>
    InvalidRepeatingSection = 9,
}
