using System.Globalization;

namespace DocToolkit;

/// <summary>
/// One block of content in a document built by <see cref="DocxEditor.Create(System.Collections.Generic.IEnumerable{DocxBlock})"/>.
///
/// The hierarchy is CLOSED: the constructor is <c>private protected</c>, the concrete types are
/// <c>internal sealed</c>, and a block can only be obtained from one of the factory methods below.
/// A consumer therefore cannot define a block the writer has never heard of — an unrenderable block
/// is unrepresentable rather than a runtime failure.
///
/// Each factory validates its arguments immediately, so a bad value throws at the line that
/// produced it rather than later inside a <see cref="DocxEditor.Create(System.Collections.Generic.IEnumerable{DocxBlock})"/> call assembling many
/// blocks at once.
/// </summary>
public abstract class DocxBlock
{
    // A CLASS, not a record. C# requires a non-sealed record's compiler-generated copy constructor
    // to be public or protected (CS8878), and protected reaches derived types in ANY assembly - so
    // an external caller can derive through it:
    //     public sealed record Evil : DocxBlock { public Evil(DocxBlock seed) : base(seed) { } }
    // A class has no synthesized copy constructor, so that route does not exist. Verified by
    // compilation. Nothing here needs value equality, and record equality was actively wrong for
    // TableBlock, whose IReadOnlyList members have no structural equality.
    //
    // private protected, NOT private: a private constructor is inaccessible to derived types unless
    // they are nested inside this one, so the internal sealed types below would not compile.
    // private protected is the closed-hierarchy idiom - derivable inside this assembly, not outside.
    private protected DocxBlock() { }

    /// <summary>
    /// A heading at <paramref name="level"/>, rendered with a real Word heading style so it appears
    /// in the navigation pane and can drive a table of contents.
    /// </summary>
    /// <param name="text">The heading's text.</param>
    /// <param name="level">
    /// 1 to 6, matching HTML <c>h1</c>–<c>h6</c>. Word itself defines nine; six is the deliberate
    /// stopping point. Out of range throws rather than clamping, because a silently demoted heading
    /// is only ever noticed in the finished document.
    /// </param>
    /// <exception cref="ArgumentNullException"><paramref name="text"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="level"/> is not 1–6.</exception>
    public static DocxBlock Heading(string text, int level)
    {
        ArgumentNullException.ThrowIfNull(text);
        if (level is < 1 or > 6)
            throw new ArgumentOutOfRangeException(nameof(level), level, "Heading level must be 1 to 6.");

        return new HeadingBlock(text, level);
    }

    /// <summary>A body paragraph. Empty text is allowed and produces a blank line.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="text"/> is null.</exception>
    public static DocxBlock Paragraph(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        return new ParagraphBlock(text);
    }

    /// <summary>
    /// A table with a bold header row.
    ///
    /// Cell values are typed the same WAY as
    /// <see cref="WorkbookEditor.Create(string, System.Collections.Generic.IEnumerable{System.Collections.Generic.IEnumerable{object}})"/> —
    /// <see cref="bool"/>, <see cref="DateTime"/>, <see cref="DateOnly"/>, <see cref="TimeOnly"/>
    /// and <see cref="TimeSpan"/> are handled by name, and everything else, numbers included, is
    /// formatted through <see cref="CultureInfo.InvariantCulture"/> — but they do <b>not</b> always
    /// render to the same TEXT, and that is deliberate. A spreadsheet cell stores a typed value
    /// that Excel formats on display; a Word table cell stores only text, so the value is rendered
    /// here and the rendering is this library's choice. For example:
    /// <list type="bullet">
    /// <item>Dates use ISO-8601 (<c>2026-08-06</c>). The workbook path renders the same date in a
    /// slashed form whose field order follows the reader's culture — <c>08/06/2026</c> or
    /// <c>06/08/2026</c> for this date, depending on the machine. That ambiguity between day-first
    /// and month-first readings is precisely what ISO-8601 removes.</item>
    /// <item>A <see cref="TimeSpan"/> keeps its days (<c>1.02:03:04</c>). Excel flattens them into
    /// hours (<c>26:03:04</c>), which is a different unit convention, not just a different
    /// pattern.</item>
    /// <item>A very large <see cref="decimal"/>, <see cref="long"/> or <see cref="ulong"/> keeps
    /// every digit here. The workbook path converts through <see cref="double"/> first, so it can
    /// show a rounded value — <c>decimal.MaxValue</c> becomes <c>7.92281625142643E+28</c> there
    /// and stays <c>79228162514264337593543950335</c> here. That is a different number, not a
    /// different pattern.</item>
    /// </list>
    /// Those examples are illustrative, not an exhaustive list of the differences. What both
    /// guarantee identically is that the same input produces the same output on every machine,
    /// because neither consults the current culture.
    ///
    /// Rows are materialised immediately; mutating the caller's sequence afterwards does not change
    /// the block.
    ///
    /// A row SHORTER than the header is padded with empty cells: ragged data is normal, and padding
    /// discards nothing. A row LONGER than the header throws, and the asymmetry is the point — the
    /// surplus values could only be dropped, which is data loss with no signal anywhere. Word would
    /// render the table as though it were complete. Same reasoning as <see cref="Heading"/> refusing
    /// to clamp an out-of-range level: a loss that only shows up in the finished document is worth
    /// more than the convenience of accepting the call.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="headers"/> or <paramref name="rows"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="headers"/> is empty, a row is null, or a
    /// row has more cells than there are headers.</exception>
    public static DocxBlock Table(
        IEnumerable<string> headers, IEnumerable<IEnumerable<object?>> rows)
    {
        ArgumentNullException.ThrowIfNull(headers);
        ArgumentNullException.ThrowIfNull(rows);

        var materialisedHeaders = headers.ToList();
        if (materialisedHeaders.Count == 0)
            throw new ArgumentException("A table needs at least one header.", nameof(headers));

        var materialisedRows = rows
            .Select((row, index) =>
            {
                if (row is null)
                    throw new ArgumentException($"Row {index + 1} was null.", nameof(rows));

                var cells = (IReadOnlyList<object?>)row.ToList();
                if (cells.Count > materialisedHeaders.Count)
                    throw new ArgumentException(
                        $"Row {index + 1} has {cells.Count} cells but the table has " +
                        $"{materialisedHeaders.Count} " +
                        $"{(materialisedHeaders.Count == 1 ? "column" : "columns")}.",
                        nameof(rows));

                return cells;
            })
            .ToList();

        return new TableBlock(materialisedHeaders, materialisedRows);
    }

    /// <summary>
    /// An inline image. PNG and JPEG only, decided by magic bytes rather than by anything the caller
    /// says — a part declaring <c>image/png</c> while holding JPEG bytes renders as a blank frame,
    /// silently.
    ///
    /// Size is in points, matching <see cref="DocxEditor.ReplaceImage"/>. Omit both and the image's
    /// intrinsic size at 96 DPI is used; give one and the other scales to preserve the aspect ratio;
    /// give both and the image is stretched, distortion accepted as the caller's choice.
    ///
    /// <paramref name="altText"/> becomes the drawing's <c>descr</c>, which is what a screen reader
    /// announces. Omit it and the attribute is omitted too, rather than filled with a placeholder:
    /// a generated value like "Image 1" is worse than nothing, because it is read out as though it
    /// described the picture. Supply it for any image carrying meaning; leave it off for one that is
    /// purely decorative.
    /// </summary>
    /// <param name="image">PNG or JPEG bytes.</param>
    /// <param name="widthPoints">Rendered width in points, or null to derive it.</param>
    /// <param name="heightPoints">Rendered height in points, or null to derive it.</param>
    /// <param name="altText">
    /// What a screen reader announces in place of the image. Null omits it.
    /// </param>
    /// <exception cref="ArgumentNullException"><paramref name="image"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="image"/> is empty, or is neither PNG nor JPEG.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// A supplied size is zero or negative, or the resulting size is larger than a drawing extent can
    /// hold — see <see cref="DocxEditor.ReplaceImage"/> for the same bound on the editing path.
    /// </exception>
    public static DocxBlock Image(
        byte[] image, double? widthPoints = null, double? heightPoints = null, string? altText = null)
    {
        ArgumentNullException.ThrowIfNull(image);
        if (image.Length == 0) throw new ArgumentException("Image content was empty.", nameof(image));
        if (widthPoints is <= 0)
            throw new ArgumentOutOfRangeException(nameof(widthPoints), widthPoints, "Width must be positive.");
        if (heightPoints is <= 0)
            throw new ArgumentOutOfRangeException(nameof(heightPoints), heightPoints, "Height must be positive.");

        // Rejected here rather than at write time, and surfaced as ArgumentException rather than the
        // DocumentConversionException ImageInspector raises: at this point it is plainly a bad
        // argument, not a document that failed to build.
        ImageInfo info;
        try
        {
            info = ImageInspector.Inspect(image);
        }
        catch (DocumentConversionException ex)
        {
            throw new ArgumentException(ex.Message, nameof(image), ex);
        }

        // The size bound is enforced HERE as well as inside Resolve, for the same reason the checks
        // above are: at the factory it is an ArgumentOutOfRangeException naming the caller's own
        // argument. Reached from Create instead, Resolve throws inside DocxDocumentWriter.Write's
        // try, which wraps it as "DocumentConversionException: Failed to create DOCX." — so without
        // this the same bad size produced a different exception type depending on which public entry
        // point the caller used. ReplaceImage calls Resolve outside its try and was always fine.
        _ = ImageInspector.Resolve(info, widthPoints, heightPoints);

        return new ImageBlock(image, widthPoints, heightPoints, altText);
    }
}

// Primary constructors (C# 12) keep these as terse as the records they replaced. Properties are
// get-only, so a block stays immutable once a factory has validated it.
internal sealed class HeadingBlock(string text, int level) : DocxBlock
{
    public string Text { get; } = text;
    public int Level { get; } = level;
}

internal sealed class ParagraphBlock(string text) : DocxBlock
{
    public string Text { get; } = text;
}

internal sealed class TableBlock(
    IReadOnlyList<string> headers, IReadOnlyList<IReadOnlyList<object?>> rows) : DocxBlock
{
    public IReadOnlyList<string> Headers { get; } = headers;
    public IReadOnlyList<IReadOnlyList<object?>> Rows { get; } = rows;
}

// The property is Bytes, not Image. As a record this collided with the static factory
// (CS8866); as a class it would merely shadow it confusingly. Either way the public factory keeps
// the name Image - it is the approved API - so the internal member is the one that moves.
internal sealed class ImageBlock(
    byte[] bytes, double? widthPoints, double? heightPoints, string? altText) : DocxBlock
{
    public byte[] Bytes { get; } = bytes;
    public double? WidthPoints { get; } = widthPoints;
    public double? HeightPoints { get; } = heightPoints;
    public string? AltText { get; } = altText;
}
