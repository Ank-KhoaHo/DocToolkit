namespace DocToolkit;

/// <summary>
/// What a document contains that <see cref="DocxToPdfConverter"/> may not carry into the PDF.
/// </summary>
/// <remarks>
/// <b>An inventory, not a loss report.</b> Every finding says a construct is <i>present</i>. None
/// says it was dropped, because this is produced by reading the source — the conversion has not run
/// and may not run at all.
///
/// That distinction is the whole point of the type. The renderer beneath
/// <see cref="DocxToPdfConverter"/> produces no report of its own, so a report claiming to list what
/// was lost could not be checked against anything. A report of what is <i>at risk</i> can be, and
/// stays true whatever the renderer does with it.
/// </remarks>
public sealed class DocxToPdfPreflightReport
{
    internal DocxToPdfPreflightReport(IReadOnlyList<DocxToPdfPreflightFinding> findings)
        => Findings = findings;

    /// <summary>Everything found, in a stable order. Empty means nothing on the list was present.</summary>
    public IReadOnlyList<DocxToPdfPreflightFinding> Findings { get; }

    /// <summary>
    /// Whether anything was found — the "does a human need to look at this?" question, which is
    /// what a caller batching third-party documents actually asks.
    /// </summary>
    /// <remarks>
    /// <b>Empty is not a promise the conversion is faithful.</b> It means none of the constructs
    /// this version knows about is present. Charts, SmartArt and embedded objects are not yet
    /// among them — see <see cref="DocxToPdfPreflight"/>.
    /// </remarks>
    public bool HasFindings => Findings.Count > 0;
}

/// <summary>How confident this library is that the construct is a problem.</summary>
public enum DocxToPdfRisk
{
    /// <summary>
    /// Listed on reasoning rather than measurement. No finding carries this yet, and one should not
    /// be added without saying why in the same change.
    /// </summary>
    Possible = 0,

    /// <summary>
    /// <b>Measured.</b> This project has converted a document containing it and watched the content
    /// fail to arrive in the PDF, with a control proving the rest of the document rendered.
    /// </summary>
    Known = 1,
}

/// <summary>One construct found in the source.</summary>
public sealed class DocxToPdfPreflightFinding
{
    internal DocxToPdfPreflightFinding(string code, string construct, string message, int count,
                                       DocxToPdfRisk risk)
    {
        Code = code;
        Construct = construct;
        Message = message;
        Count = count;
        Risk = risk;
    }

    /// <summary>A stable identifier, safe to branch on. Never localised, never reworded.</summary>
    public string Code { get; }

    /// <summary>What was found, in the words a reader would use.</summary>
    public string Construct { get; }

    /// <summary>What may happen to it. Deliberately hedged — see the report's remarks.</summary>
    public string Message { get; }

    /// <summary>
    /// How many were found, so a caller can triage by volume: one footnote in a memo is not the
    /// same problem as ninety in a contract.
    /// </summary>
    public int Count { get; }

    /// <summary>Whether the risk was measured or reasoned. See <see cref="DocxToPdfRisk"/>.</summary>
    public DocxToPdfRisk Risk { get; }
}
