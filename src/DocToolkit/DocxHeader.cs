namespace DocToolkit;

/// <summary>Where a header or footer line sits across the page.</summary>
public enum HeaderAlignment
{
    /// <summary>Against the left margin.</summary>
    Left = 0,

    /// <summary>Centred between the margins.</summary>
    Center = 1,

    /// <summary>Against the right margin.</summary>
    Right = 2,
}

/// <summary>
/// The content of one header or footer.
/// </summary>
/// <remarks>
/// Footers use this type too. Word's content model for the two is identical, and a second type
/// differing only in its name would be one more thing to learn for no gain.
///
/// Attach it with <see cref="PageSetup.WithHeader(DocxHeader)"/> or
/// <see cref="PageSetup.WithFooter(DocxHeader)"/>; every producer that takes a
/// <see cref="PageSetup"/> then honours it, with no extra overload.
/// </remarks>
public sealed class DocxHeader
{
    private DocxHeader(HeaderAlignment alignment, IReadOnlyList<DocxHeaderSegment> segments)
    {
        Alignment = alignment;
        Segments = segments;
    }

    /// <summary>Where the line sits across the page.</summary>
    public HeaderAlignment Alignment { get; }

    /// <summary>The pieces of the line, in order.</summary>
    public IReadOnlyList<DocxHeaderSegment> Segments { get; }

    /// <summary>A header of a single run of text.</summary>
    /// <param name="text">The text.</param>
    /// <param name="alignment">Where it sits. Defaults to <see cref="HeaderAlignment.Left"/>.</param>
    /// <exception cref="ArgumentNullException"><paramref name="text"/> is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="alignment"/> is not a defined value.</exception>
    public static DocxHeader Text(string text, HeaderAlignment alignment = HeaderAlignment.Left)
        => Of(alignment, DocxHeaderSegment.Text(text));

    /// <summary>A header assembled from segments, in order.</summary>
    /// <param name="alignment">Where the line sits.</param>
    /// <param name="segments">The pieces, in order. At least one is required.</param>
    /// <exception cref="ArgumentNullException"><paramref name="segments"/>, or an entry in it, is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="alignment"/> is not a defined value.</exception>
    /// <exception cref="ArgumentException"><paramref name="segments"/> is empty.</exception>
    public static DocxHeader Of(HeaderAlignment alignment, params DocxHeaderSegment[] segments)
    {
        if (alignment is not (HeaderAlignment.Left or HeaderAlignment.Center or HeaderAlignment.Right))
        {
            throw new ArgumentOutOfRangeException(
                nameof(alignment), alignment, "Alignment must be Left, Center or Right.");
        }

        ArgumentNullException.ThrowIfNull(segments);

        if (segments.Length == 0)
        {
            throw new ArgumentException(
                "A header needs at least one segment. To have no header, do not set one — an empty "
                + "header still reserves a blank line on every page.",
                nameof(segments));
        }

        // Copied, not aliased: the caller's array stays theirs to reuse or mutate, and this type
        // advertises itself as immutable.
        var copy = new DocxHeaderSegment[segments.Length];
        for (var i = 0; i < segments.Length; i++)
        {
            copy[i] = segments[i] ?? throw new ArgumentNullException(
                nameof(segments), $"Segment {i} is null.");
        }

        // Wrapped rather than handed back as the array itself: a caller downcasting the
        // IReadOnlyList<T> back to DocxHeaderSegment[] would otherwise reach the very array this
        // copy exists to keep private, and mutate it in place after construction.
        return new DocxHeader(alignment, Array.AsReadOnly(copy));
    }

    /// <summary>The segments joined, with fields shown as <c>{PAGE}</c> and <c>{NUMPAGES}</c>.</summary>
    public override string ToString() => string.Concat(Segments);
}
