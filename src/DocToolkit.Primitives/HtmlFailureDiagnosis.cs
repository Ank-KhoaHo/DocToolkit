namespace DocToolkit;

/// <summary>
/// Turns an HTML-conversion failure the package can actually identify into a message that names it.
/// </summary>
/// <remarks>
/// <b>Every message here must name a cause the exception can DISTINGUISH</b>, which is the lesson
/// this package has already had to learn twice: a timeout message that asserted "a TCP connect that
/// will never complete" as fact cost real time chasing a network regression that did not exist, and
/// <c>PdfEditor</c>'s read failure now names three candidates rather than guessing between them. So
/// the test here is a stack frame in a specific type, not the exception's type alone - an
/// <see cref="IndexOutOfRangeException"/> by itself says nothing about tables.
///
/// <b>Anything not recognised falls through to the generic wrapper</b>, deliberately. A diagnosis
/// that is right most of the time is worse than no diagnosis, because it sends the reader somewhere
/// specific and wrong.
/// </remarks>
internal static class HtmlFailureDiagnosis
{
    /// <summary>
    /// Returns a message naming the cause, or <see langword="null"/> when the failure is not one of
    /// the recognised shapes and the caller should use its generic wrapper.
    /// </summary>
    internal static string? Describe(Exception ex)
    {
        if (OverhangingRowSpan(ex)) return OverhangingRowSpanMessage;
        if (InvalidCharacter(ex, out var character)) return InvalidCharacterMessage(character);
        return null;
    }

    /// <summary>
    /// A table cell whose <c>rowspan</c> reaches past the last row of its table.
    /// </summary>
    /// <remarks>
    /// <b>Measured 2026-08-17 across 179 real .gov pages: 14 of them - 7.7% - fail this way</b>,
    /// which made it the single most frequent conversion failure in the set, and the caller was told
    /// only "see the inner exception" over a bare <see cref="IndexOutOfRangeException"/> naming no
    /// table.
    ///
    /// The parser sizes an accumulator to the rows a table actually has and then indexes it with the
    /// rowspan, unclamped, so any cell reaching past the last row writes out of bounds. The boundary
    /// is exact and was measured as a grid: one row breaks at <c>rowspan=2</c>, two rows at 3, three
    /// rows at 4, four rows survives 4. <c>colspan</c> is unaffected at every value tried, and
    /// <c>rowspan="0"</c> is fine.
    ///
    /// <b>The markup is not malformed</b> - browsers clamp such a rowspan to the rows that exist, so
    /// it renders correctly everywhere else, which is exactly why real pages carry it. The corpus
    /// held spans of 2, 3, 14, 100 and 103 against tables of one to three rows; a <c>rowspan="100"</c>
    /// is somebody wanting a cell to reach the bottom of a table and picking a number bigger than the
    /// row count.
    ///
    /// <b>The frame is the discriminator, not the exception type.</b> Matching on
    /// <see cref="IndexOutOfRangeException"/> alone would put this message on any index error
    /// anywhere in the conversion, which is precisely the over-claiming this class exists to avoid.
    /// </remarks>
    private const string OverhangingRowSpanMessage =
        "Failed to convert HTML to DOCX: a table cell has a rowspan reaching past the last row of "
        + "its table, and the HTML parser this package uses cannot read that - it raises an index "
        + "error rather than reporting it. The markup is not invalid: browsers clamp such a rowspan "
        + "to the rows that exist, so the page renders correctly in a browser. To convert it, reduce "
        + "the rowspan to at most the number of rows below the cell, or remove the attribute. See "
        + "the inner exception for the parser's own error.";

    /// <summary>The frame that identifies it. Private to the parser, so this is a heuristic on a
    /// name - but a specific one, and it fails closed: no match means no claim.</summary>
    private const string RowSpanFrame = "TableExpression.GuessColumnsCount";

    private static bool OverhangingRowSpan(Exception ex) =>
        ex is IndexOutOfRangeException
        && (ex.StackTrace ?? string.Empty).Contains(RowSpanFrame, StringComparison.Ordinal);

    /// <summary>
    /// A character the XML writer refuses, which is what a caller gets for handing this converter
    /// something that is not text.
    /// </summary>
    /// <remarks>
    /// <b>Measured 2026-08-20 over govdocs1: 8 of 8 JPEGs and 1 of 12 .txt files fail this way</b>,
    /// and the caller was told only <i>"See the inner exception"</i> over
    /// <i>"'', hexadecimal value 0x10, is an invalid character"</i> - a message about a character
    /// nobody typed, in a document they never wrote.
    ///
    /// <b>This one matches the MESSAGE, not a stack frame, and that is the opposite of the rowspan
    /// case above for a reason.</b> There the exception is a bare
    /// <see cref="IndexOutOfRangeException"/> whose message says nothing, so only the frame can
    /// identify it. Here the message is itself the discriminator - that exact wording is produced
    /// only when an XML writer is handed a character XML cannot represent - while the frame is an
    /// <c>internal</c> writer class whose name differs between the UTF-8 and UTF-16 implementations.
    /// Matching the more stable of the two available discriminators is the same judgement, not a
    /// departure from it.
    ///
    /// <b>The message names two candidates and picks neither</b>, because the exception cannot tell
    /// them apart: binary content passed to this converter, and a stray control character inside
    /// genuine HTML, produce byte-identical failures. Verified - a JPEG and
    /// <c>&lt;p&gt;before[U+0010]after&lt;/p&gt;</c> are indistinguishable here. Asserting the
    /// first as fact would repeat the timeout message this class's remarks already record.
    ///
    /// <b>Ordinary markup is unaffected</b>, measured rather than assumed: tabs, CR/LF, character
    /// entities including <c>&amp;#10;</c>, accented Latin, CJK and astral-plane emoji all convert.
    /// Only C0 control characters are refused, which is XML's rule and not this package's.
    /// </remarks>
    private static string InvalidCharacterMessage(string character) =>
        $"Failed to convert HTML to DOCX: the content contains {character}, a control character that "
        + "is not legal in a Word document, so the XML writer refuses it. TWO THINGS COMMONLY CAUSE "
        + "THIS, and the error cannot tell them apart. Either the content is not HTML at all - "
        + "passing the bytes of an image, a PDF or an Office file to this converter produces exactly "
        + "this, and each of those has its own reader on DocToolkit - or it is genuine HTML carrying "
        + "a stray control character, in which case strip characters below U+0020 other than tab, "
        + "carriage return and line feed. Ordinary markup is unaffected: tabs, newlines, character "
        + "entities, accented text, CJK and emoji all convert. See the inner exception for the "
        + "writer's own error.";

    /// <summary>
    /// The writer's wording, which is stable across .NET versions and carries the character itself.
    /// </summary>
    /// <remarks>
    /// Both halves are required. <see cref="ArgumentException"/> alone is far too broad - a great
    /// deal of this pipeline throws it, and a real one does: U+FFFE is refused by the same writer,
    /// as the same type, with a different message. The phrase alone could in principle arrive on
    /// some other type. Matching the pair fails closed, which is this class's standing rule.
    /// </remarks>
    private static bool InvalidCharacter(Exception ex, out string character)
    {
        character = "a character";
        if (ex is not ArgumentException) return false;

        var message = ex.Message ?? string.Empty;
        if (!message.Contains("hexadecimal value 0x", StringComparison.Ordinal)
            || !message.Contains("is an invalid character", StringComparison.Ordinal)) return false;

        // Quote the character back rather than making the reader re-read the inner exception.
        const string Marker = "hexadecimal value ";
        var at = message.IndexOf(Marker, StringComparison.Ordinal) + Marker.Length;
        var end = message.IndexOf(',', at);
        if (end > at) character = message[at..end].Trim();
        return true;
    }
}
