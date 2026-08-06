namespace DocToolkit;

/// <summary>
/// One block of content in a document built by <see cref="DocxEditor.Create"/>.
///
/// The hierarchy is CLOSED: the constructor is <c>private protected</c>, the concrete types are
/// <c>internal sealed</c>, and a block can only be obtained from one of the factory methods below.
/// A consumer therefore cannot define a block the writer has never heard of — an unrenderable block
/// is unrepresentable rather than a runtime failure.
///
/// Each factory validates its arguments immediately, so a bad value throws at the line that
/// produced it rather than later inside a <see cref="DocxEditor.Create"/> call assembling many
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
    /// A table with a bold header row. Cell values follow exactly the typing and culture rules of
    /// <see cref="WorkbookEditor.Create"/> — see that method. A Word table cell has no type of its
    /// own, unlike a spreadsheet cell, so "as a number" means invariant formatting here.
    ///
    /// Rows are materialised immediately; mutating the caller's sequence afterwards does not change
    /// the block.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="headers"/> or <paramref name="rows"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="headers"/> is empty, or a row is null.</exception>
    public static DocxBlock Table(
        IEnumerable<string> headers, IEnumerable<IEnumerable<object?>> rows)
    {
        ArgumentNullException.ThrowIfNull(headers);
        ArgumentNullException.ThrowIfNull(rows);

        var materialisedHeaders = headers.ToList();
        if (materialisedHeaders.Count == 0)
            throw new ArgumentException("A table needs at least one header.", nameof(headers));

        var materialisedRows = rows
            .Select((row, index) => row is null
                ? throw new ArgumentException($"Row {index + 1} was null.", nameof(rows))
                : (IReadOnlyList<object?>)row.ToList())
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
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="image"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="image"/> is empty, or is neither PNG nor JPEG.</exception>
    /// <exception cref="ArgumentOutOfRangeException">A supplied size is zero or negative.</exception>
    public static DocxBlock Image(
        byte[] image, double? widthPoints = null, double? heightPoints = null)
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
        try
        {
            _ = ImageInspector.Inspect(image);
        }
        catch (DocumentConversionException ex)
        {
            throw new ArgumentException(ex.Message, nameof(image), ex);
        }

        return new ImageBlock(image, widthPoints, heightPoints);
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
    byte[] bytes, double? widthPoints, double? heightPoints) : DocxBlock
{
    public byte[] Bytes { get; } = bytes;
    public double? WidthPoints { get; } = widthPoints;
    public double? HeightPoints { get; } = heightPoints;
}
