namespace DocToolkit;

/// <summary>
/// A cell value meaning "this cell holds a formula". Use it anywhere a cell value is accepted —
/// inside <see cref="XlsxSheet"/> rows, inside rows passed to
/// <see cref="WorkbookEditor.AppendRows(byte[], string, System.Collections.Generic.IEnumerable{System.Collections.Generic.IEnumerable{object}})"/>,
/// or as the value argument to
/// <see cref="WorkbookEditor.SetCell(byte[], string, string, object?)"/>.
///
/// <para><b>No cached result is written.</b> The file carries the formula and nothing else.
/// Excel recalculates when it opens the file, and this package's own readers
/// (<see cref="WorkbookEditor.ReadCell(byte[], string, string)"/>,
/// <see cref="WorkbookEditor.ReadSheet(byte[], string)"/>) compute the value on read — but a
/// third-party reader that only reads cached values, such as openpyxl with
/// <c>data_only=True</c>, sees an empty cell until Excel has opened and saved the file.</para>
///
/// <para><b>This package's own <c>XlsxToPdfConverter</c> used to be exactly such a reader</b> —
/// its renderer reads the raw cell rather than going through ClosedXML's lazy evaluation, so a
/// formula cell rendered its own source text where the value belonged, found 2026-08-31. It now
/// calculates before rendering, so this is no longer a caveat for it specifically. It remains one
/// for anything else that reads this package's output: use
/// <see cref="WorkbookEditor.EvaluateFormulas(byte[])"/> to write the computed value into the file
/// itself before handing it to a reader that will not recalculate on its own, and
/// <see cref="WorkbookEditor.InspectFormulas(byte[])"/> to find out first whether the engine
/// understands a given formula well enough to trust its value.</para>
///
/// <para>A formula that cannot be evaluated reads back as its Excel error string — <c>#DIV/0!</c>,
/// <c>#NAME?</c>, <c>#REF!</c> — rather than throwing, which is what Excel itself shows.</para>
/// </summary>
public sealed class XlsxFormula
{
    private XlsxFormula(string formula) => Formula = formula;

    /// <summary>The formula, without a leading <c>=</c>, which is how the file stores it.</summary>
    public string Formula { get; }

    /// <summary>
    /// Creates a formula cell value. A leading <c>=</c> is optional and is stripped: the file
    /// format stores formulas without it, and accepting both spellings stops the value that
    /// round-trips from differing from the one that was written.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="formula"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="formula"/> is empty, whitespace, or just "=".</exception>
    public static XlsxFormula From(string formula)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(formula);

        var trimmed = formula.Trim();
        if (trimmed[0] == '=')
            trimmed = trimmed[1..].Trim();

        if (trimmed.Length == 0)
            throw new ArgumentException("Formula was empty.", nameof(formula));

        return new XlsxFormula(trimmed);
    }
}
