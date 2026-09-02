namespace DocToolkit;

/// <summary>
/// One word of a PDF's text layer, and where it sits on the page.
/// </summary>
/// <remarks>
/// <see cref="PdfEditor.ExtractText(byte[])"/> answers <i>what</i> a page says;
/// this answers <i>where</i> it says it. Locating a total on an invoice, checking a stamp landed
/// inside the margin, or routing a scanned form by region all need the position and cannot be
/// built on the string alone.
///
/// <para>
/// <b>A page with no text layer produces no words.</b> A scanned document is images, so it yields
/// an empty list per page rather than a failure — the same rule
/// <see cref="PdfEditor.ExtractText(byte[])"/> already documents, and OCR remains out of scope.
/// </para>
///
/// <para>
/// What counts as a word is PdfPig's segmentation of the page's text-showing operators, not a
/// dictionary. Punctuation usually rides along with the token it touches.
/// </para>
/// </remarks>
public sealed class PdfWord
{
    internal PdfWord(string text, PdfBounds bounds)
    {
        Text = text;
        Bounds = bounds;
    }

    /// <summary>The word's text, exactly as the page's text layer holds it.</summary>
    public string Text { get; }

    /// <summary>Where the word sits on its page, in PDF user-space points.</summary>
    public PdfBounds Bounds { get; }

    /// <summary>The word and its lower-left corner, for logs and test failure messages.</summary>
    public override string ToString() =>
        $"{Text} @ ({Bounds.Left:F1},{Bounds.Bottom:F1})";
}
