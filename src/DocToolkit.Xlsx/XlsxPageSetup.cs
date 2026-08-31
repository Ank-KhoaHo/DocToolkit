namespace DocToolkit;

/// <summary>A sheet's print orientation.</summary>
public enum XlsxPageOrientation
{
    /// <summary>Taller than wide.</summary>
    Portrait = 0,

    /// <summary>Wider than tall.</summary>
    Landscape = 1,
}

/// <summary>
/// A sheet's own print setup: orientation, the range that prints, and the rows that repeat at the
/// top of every printed page.
/// </summary>
/// <remarks>
/// <b>A workbook this package writes carries no page setup at all until this is applied</b> — the
/// same class of gap the missing <c>sectPr</c> was for DOCX before 0.13.0. <c>XlsxToPdfConverter</c>
/// renders whatever the file happens to carry, which for a freshly created workbook is nothing, so
/// the render falls back to the reader's own default rather than reflecting anything the caller
/// asked for.
///
/// Deliberately a separate type from <see cref="DocToolkit.PageSetup"/> rather than a shared one —
/// the two formats' print concepts do not line up. DOCX's <c>PageSetup</c> is a page SIZE
/// (dimensions and margins); a worksheet's print setup is a VIEW onto data that already exists
/// (which range prints, which rows repeat), and has no page-size concept of its own to hold. Forcing
/// one type to cover both would collide them, the same reason <c>PdfMetadata</c> stays
/// separate from <see cref="DocumentMetadata"/>.
/// </remarks>
public sealed class XlsxPageSetup
{
    private XlsxPageSetup(XlsxPageOrientation orientation, string? printArea, int? repeatRowCount)
    {
        Orientation = orientation;
        PrintArea = printArea;
        RepeatRowCount = repeatRowCount;
    }

    /// <summary>The print orientation.</summary>
    public XlsxPageOrientation Orientation { get; }

    /// <summary>
    /// The range that prints, such as <c>A1:D50</c>. <see langword="null"/> prints the sheet's
    /// whole used range, Excel's own default.
    /// </summary>
    public string? PrintArea { get; }

    /// <summary>
    /// How many rows, counted from row 1, repeat at the top of every printed page.
    /// <see langword="null"/> repeats none.
    /// </summary>
    public int? RepeatRowCount { get; }

    /// <summary>Describes a sheet's print setup.</summary>
    /// <param name="orientation">The print orientation. Defaults to <see cref="XlsxPageOrientation.Portrait"/>.</param>
    /// <param name="printArea">
    /// The range that prints, such as <c>A1:D50</c>. <see langword="null"/> prints the whole used
    /// range.
    /// </param>
    /// <param name="repeatRowCount">
    /// How many rows, counted from row 1, repeat at the top of every printed page.
    /// <see langword="null"/> repeats none.
    /// </param>
    /// <exception cref="ArgumentException">
    /// <paramref name="printArea"/> is empty or whitespace, or names a sheet.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="repeatRowCount"/> is not positive.</exception>
    public static XlsxPageSetup Of(
        XlsxPageOrientation orientation = XlsxPageOrientation.Portrait,
        string? printArea = null,
        int? repeatRowCount = null)
    {
        if (printArea is not null)
        {
            if (printArea.Length == 0 || printArea.Trim().Length == 0)
                throw new ArgumentException("Print area was blank.", nameof(printArea));

            if (printArea.Contains('!'))
            {
                throw new ArgumentException(
                    $"\"{printArea}\" names a sheet, and the sheet qualifier is silently discarded "
                    + "rather than honoured. Pass the range alone; Format's own sheetName parameter "
                    + "chooses the sheet.",
                    nameof(printArea));
            }
        }

        if (repeatRowCount is <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(repeatRowCount), repeatRowCount, "Repeat row count must be positive.");
        }

        return new XlsxPageSetup(orientation, printArea, repeatRowCount);
    }
}
