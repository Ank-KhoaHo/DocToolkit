using OfficeIMOPictureValue = OfficeIMO.Word.WordContentControlPictureValue;

namespace DocToolkit;

/// <summary>
/// A value read from, or written into, a Word content control.
/// </summary>
/// <remarks>
/// <b>This type exists to make one guarantee structural.</b> The library beneath offers
/// <c>WordContentControlPictureValue.FromFile(path)</c>, which reads the filesystem. This package
/// deliberately refuses to read local files for images — see <c>AllowLocalImages</c>, which is false
/// by default — so that route must not be reachable here either. There is therefore <b>no picture
/// factory taking a path or a URI</b>. A caller who wants a file on disk reads it themselves, in one
/// line, and the decision is visibly theirs.
///
/// Upstream models these values as <c>object</c>. A typed value is what turns that rule from
/// something a guard has to remember into something this API cannot express.
/// </remarks>
public sealed class DocxFormValue
{
    private readonly byte[]? _bytes;

    private DocxFormValue(
        DocxFormValueKind kind, string? text, bool? isChecked, DateTime? date,
        byte[]? bytes, string? fileName)
    {
        Kind = kind;
        Text = text;
        Checked = isChecked;
        Date = date;
        _bytes = bytes is null ? null : (byte[])bytes.Clone();
        FileName = fileName;
    }

    /// <summary>Which kind of value this is; which other properties are set follows from it.</summary>
    public DocxFormValueKind Kind { get; }

    /// <summary>
    /// The text — for <see cref="DocxFormValueKind.Text"/>, <see cref="DocxFormValueKind.Choice"/>
    /// and <see cref="DocxFormValueKind.Other"/>. Null otherwise.
    /// </summary>
    public string? Text { get; }

    /// <summary>Whether the box is ticked, for <see cref="DocxFormValueKind.Checked"/>.</summary>
    public bool? Checked { get; }

    /// <summary>The date, for <see cref="DocxFormValueKind.Date"/>.</summary>
    public DateTime? Date { get; }

    /// <summary>The image content, for <see cref="DocxFormValueKind.Picture"/>.</summary>
    /// <remarks>
    /// A <b>copy</b>, on the way in and on the way out. Without both, this sealed type with
    /// getter-only properties would not actually be immutable: a caller who kept their array could
    /// change what a value means after <c>Validate</c> had approved it and before <c>Fill</c> wrote
    /// it, and a caller could mutate what a report reports by writing through this property.
    /// </remarks>
    public byte[]? Bytes => _bytes is null ? null : (byte[])_bytes.Clone();

    /// <summary>The image's file name, for <see cref="DocxFormValueKind.Picture"/>.</summary>
    public string? FileName { get; }

    /// <summary>A plain-text value.</summary>
    /// <param name="value">The text to write.</param>
    /// <exception cref="ArgumentNullException"><paramref name="value"/> is null.</exception>
    public static DocxFormValue FromText(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        return new DocxFormValue(DocxFormValueKind.Text, value, null, null, null, null);
    }

    /// <summary>A check-box value.</summary>
    /// <param name="value">Whether the box is ticked.</param>
    public static DocxFormValue FromChecked(bool value)
        => new(DocxFormValueKind.Checked, null, value, null, null, null);

    /// <summary>A date value.</summary>
    /// <param name="value">The date to write.</param>
    public static DocxFormValue FromDate(DateTime value)
        => new(DocxFormValueKind.Date, null, null, value, null, null);

    /// <summary>A drop-down or combo-box selection.</summary>
    /// <param name="value">The option to select.</param>
    /// <exception cref="ArgumentNullException"><paramref name="value"/> is null.</exception>
    public static DocxFormValue FromChoice(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        return new DocxFormValue(DocxFormValueKind.Choice, value, null, null, null, null);
    }

    /// <summary>
    /// An image, <b>from bytes</b>. There is deliberately no overload taking a path — see the
    /// remarks on <see cref="DocxFormValue"/>.
    /// </summary>
    /// <param name="bytes">The image content.</param>
    /// <param name="fileName">A name for the image part, such as <c>logo.png</c>.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="bytes"/> or <paramref name="fileName"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="bytes"/> is empty, or <paramref name="fileName"/> is blank.
    /// </exception>
    public static DocxFormValue FromPicture(byte[] bytes, string fileName)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        if (bytes.Length == 0)
            throw new ArgumentException("Image content was empty.", nameof(bytes));
        ArgumentException.ThrowIfNullOrWhiteSpace(fileName);

        return new DocxFormValue(DocxFormValueKind.Picture, null, null, null, bytes, fileName);
    }

    /// <summary>
    /// Maps a value the underlying library handed back.
    /// </summary>
    /// <remarks>
    /// Measured against a form authored the way Word authors one: a check box returns
    /// <see cref="bool"/>, a date picker returns <see cref="DateTime"/>, a picture control returns a
    /// picture value, and a plain text control returns a <see cref="string"/>. An <b>unset</b> date
    /// picker or an unselected drop-down returns <see langword="null"/>.
    ///
    /// An earlier version of this remark said everything came back as a string, on a measurement
    /// taken against hand-built markup that contained no typed controls at all. The
    /// <see cref="DocxFormValueKind.Other"/> arm still exists for a type this API does not model, so
    /// such a value is carried rather than dropped or thrown on.
    /// </remarks>
    internal static DocxFormValue FromUpstream(object? value) => value switch
    {
        null => new DocxFormValue(DocxFormValueKind.Other, null, null, null, null, null),
        bool isChecked => FromChecked(isChecked),
        DateTime date => FromDate(date),
        string text => FromText(text),
        OfficeIMOPictureValue picture when picture.Bytes is { Length: > 0 } bytes
            => FromPicture(bytes, string.IsNullOrWhiteSpace(picture.FileName) ? "image.png" : picture.FileName),

        // A LINKED picture carries no bytes - only a URI. ToString() would store the CLR type name
        // here, which is worse than useless: DocxFormValueKind.Other promises Text carries the
        // content, and "OfficeIMO.Word.WordContentControlPictureValue" is not content. Name the URI
        // if there is one, and say plainly that there is nothing to carry if there is not.
        OfficeIMOPictureValue linked
            => new DocxFormValue(DocxFormValueKind.Other,
                linked.ExternalUri?.ToString() ?? string.Empty, null, null, null, linked.FileName),

        _ => new DocxFormValue(DocxFormValueKind.Other, value.ToString(), null, null, null, null),
    };

    /// <summary>
    /// The shape the underlying library takes. <see cref="DocxFormValueKind.Picture"/> goes through
    /// <c>FromBytes</c> — the only picture factory this API can reach.
    /// </summary>
    internal object? ToUpstream() => Kind switch
    {
        DocxFormValueKind.Checked => Checked,
        DocxFormValueKind.Date => Date,
        // The stored array, not the copying property: this is the one place inside the type
        // where copying again would be waste.
        DocxFormValueKind.Picture => OfficeIMOPictureValue.FromBytes(_bytes!, FileName!),
        _ => Text,
    };
}

/// <summary>The kinds of value a content control can hold.</summary>
public enum DocxFormValueKind
{
    /// <summary>Plain text.</summary>
    Text = 0,

    /// <summary>A check box.</summary>
    Checked = 1,

    /// <summary>A date.</summary>
    Date = 2,

    /// <summary>
    /// A selection from a drop-down or combo box.
    /// </summary>
    /// <remarks>
    /// <b>Only ever seen on the way in.</b> A drop-down's current selection is handed back as plain
    /// text, so <c>DocxForm.Inspect</c> reports it as <see cref="Text"/> — measured. A caller cannot
    /// tell a drop-down from a text control by kind alone.
    /// </remarks>
    Choice = 3,

    /// <summary>An image.</summary>
    Picture = 4,

    /// <summary>
    /// Something this API does not model. <see cref="DocxFormValue.Text"/> carries its text, so a
    /// caller sees the content rather than losing it.
    /// </summary>
    Other = 5,
}
