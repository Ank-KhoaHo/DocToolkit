namespace DocToolkit;

/// <summary>One heading found in a Markdown document by <see cref="MarkdownEditor.FindHeading"/>.</summary>
public sealed class MarkdownHeading
{
    internal MarkdownHeading(int level, string text, string anchor)
    {
        Level = level;
        Text = text;
        Anchor = anchor;
    }

    /// <summary>The heading's level: 1 for <c>#</c>, 2 for <c>##</c>, and so on.</summary>
    public int Level { get; }

    /// <summary>The heading's text, with any inline formatting markers stripped.</summary>
    public string Text { get; }

    /// <summary>
    /// The slug an in-document anchor link to this heading would use — for example
    /// <c>changed</c> for a heading reading "Changed".
    /// </summary>
    public string Anchor { get; }
}
