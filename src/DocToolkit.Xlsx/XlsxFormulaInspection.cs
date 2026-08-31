namespace DocToolkit;

/// <summary>
/// Which formulas a workbook carries, and whether each one is understood well enough to trust its
/// value — an inventory, not a promise about what the value currently is.
/// </summary>
/// <remarks>
/// <b>The reason this exists, stated once here rather than at every call site.</b> A formula this
/// package writes carries no cached result — see <see cref="XlsxFormula"/>'s own remarks —
/// and this package's own <c>ReadCell</c>/<c>ReadSheet</c> compute a value lazily, in memory, on
/// every read. That is reliable when the reader agrees with the writer about what the formula
/// means; it says nothing about whether it would. <see cref="WorkbookEditor.InspectFormulas"/> asks
/// the underlying engine directly, rather than trusting a value it produced itself.
/// </remarks>
public sealed class XlsxFormulaInspection
{
    internal XlsxFormulaInspection(IReadOnlyList<XlsxFormulaCell> formulas)
    {
        Formulas = formulas;
        SupportedFormulas = formulas.Count(f => f.IsSupported);
        UnsupportedFormulas = formulas.Count - SupportedFormulas;
    }

    /// <summary>Every formula found, in sheet then cell order.</summary>
    public IReadOnlyList<XlsxFormulaCell> Formulas { get; }

    /// <summary>How many formulas the workbook carries.</summary>
    public int TotalFormulas => Formulas.Count;

    /// <summary>How many of <see cref="Formulas"/> the engine understands.</summary>
    public int SupportedFormulas { get; }

    /// <summary>
    /// How many of <see cref="Formulas"/> the engine does not understand — each carries its own
    /// reason in <see cref="XlsxFormulaCell.UnsupportedReason"/>.
    /// </summary>
    public int UnsupportedFormulas { get; }

    /// <summary>
    /// Whether every formula is understood. A workbook with no formulas at all is trivially
    /// <see langword="true"/>.
    /// </summary>
    public bool AllSupported => UnsupportedFormulas == 0;
}

/// <summary>One formula cell, as found by <see cref="WorkbookEditor.InspectFormulas"/>.</summary>
public sealed class XlsxFormulaCell
{
    internal XlsxFormulaCell(
        string sheetName, string cellReference, string formula, bool isSupported, string? unsupportedReason)
    {
        SheetName = sheetName;
        CellReference = cellReference;
        Formula = formula;
        IsSupported = isSupported;
        UnsupportedReason = unsupportedReason;
    }

    /// <summary>The sheet the formula lives on.</summary>
    public string SheetName { get; }

    /// <summary>The cell's A1-style reference, such as <c>C1</c>.</summary>
    public string CellReference { get; }

    /// <summary>The formula text, without a leading <c>=</c> — see <see cref="XlsxFormula.Formula"/>.</summary>
    public string Formula { get; }

    /// <summary>
    /// Whether the underlying engine understands this formula. <see langword="false"/> means any
    /// value read for this cell — through this package's own readers or any other tool — is not
    /// something anything here computed; treat it as absent rather than as a number to trust.
    /// </summary>
    public bool IsSupported { get; }

    /// <summary>
    /// Why <see cref="IsSupported"/> is <see langword="false"/>. <see langword="null"/> when the
    /// formula is supported.
    /// </summary>
    public string? UnsupportedReason { get; }
}
