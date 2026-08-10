namespace DocToolkit;

/// <summary>
/// One piece of a header or footer line: literal text, or a field the reader evaluates per page.
/// </summary>
/// <remarks>
/// A page number cannot be a string. Written as text it is fixed at the moment the document is
/// generated, so it is correct on exactly one page and wrong on all the others — which is worse
/// than having no page number at all, because it looks right.
/// </remarks>
public abstract class DocxHeaderSegment
{
    private DocxHeaderSegment()
    {
    }

    /// <summary>Literal text.</summary>
    /// <param name="text">The text. May be empty; may not be null.</param>
    /// <exception cref="ArgumentNullException"><paramref name="text"/> is null.</exception>
    public static DocxHeaderSegment Text(string text)
    {
        ArgumentNullException.ThrowIfNull(text);
        return new LiteralSegment(text);
    }

    /// <summary>The current page number, as a field the reader evaluates.</summary>
    public static DocxHeaderSegment PageNumber { get; } = new FieldSegment(" PAGE ", "{PAGE}");

    /// <summary>The total page count, as a field the reader evaluates.</summary>
    public static DocxHeaderSegment PageCount { get; } = new FieldSegment(" NUMPAGES ", "{NUMPAGES}");

    internal sealed class LiteralSegment(string text) : DocxHeaderSegment
    {
        public string Value { get; } = text;

        public override string ToString() => Value;
    }

    internal sealed class FieldSegment(string instruction, string display) : DocxHeaderSegment
    {
        /// <summary>The OOXML field instruction, including its surrounding spaces.</summary>
        public string Instruction { get; } = instruction;

        public override string ToString() => display;
    }
}
