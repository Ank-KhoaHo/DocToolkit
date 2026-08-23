namespace DocToolkit;

/// <summary>
/// Page setup, remote-image policy and fonts for one HTML to PDF conversion.
/// </summary>
/// <remarks>
/// <b>This type exists because the three settings are independent and the overloads were not.</b>
/// Page setup and remote images already forced a <c>(page, options)</c> overload; fonts made a
/// third axis, and one overload per combination is nine signatures that a fourth axis would turn
/// into eighteen.
///
/// <para><b>Wiring fonts into the old shape would have been worse than the gap it closed.</b> They
/// would have applied only when no page and no remote-image setting were in play - a setting that
/// silently stops taking effect depending on unrelated configuration. That is not hypothetical: it
/// is the bug <c>HtmlToPdfConverterService</c> already had and fixed, where naming a page quietly
/// opted a call back out of remote fetching a consumer had enabled.</para>
///
/// <para><b>Every default matches the overload it replaces</b>, so moving to this type cannot
/// change behaviour: <see cref="PageSetup.A4"/>, no remote fetching, no supplied fonts. In
/// particular <see cref="RemoteImage"/> being <see langword="null"/> is the offline guarantee, not
/// an unset value waiting to be filled in.</para>
/// </remarks>
/// <example>
/// <code source="../../tests/DocToolkit.Tests/DocumentationExamples.cs" region="HtmlToPdfOptions"/>
/// </example>
public sealed class HtmlToPdfOptions
{
    /// <summary>
    /// Paper size and margins. Defaults to <see cref="PageSetup.A4"/>.
    /// </summary>
    /// <remarks>
    /// A4 rather than Letter deliberately. A document that does not describe its own paper renders
    /// on whatever the reader's template chooses, which is the correctness defect 0.13.0 fixed.
    /// </remarks>
    public PageSetup Page { get; init; } = PageSetup.A4;

    /// <summary>
    /// Bounds applied when remote images are fetched. <see langword="null"/> - the default -
    /// fetches nothing.
    /// </summary>
    /// <remarks>
    /// <b>Leaving this null is the offline guarantee</b>, not an omission: no socket is opened on
    /// that path at all. Supplying one opts in, and every bound on
    /// <see cref="RemoteImageOptions"/> then applies.
    /// </remarks>
    public RemoteImageOptions? RemoteImage { get; init; }

    /// <summary>
    /// Fonts supplied for characters the PDF renderer cannot otherwise encode.
    /// <see langword="null"/> - the default - supplies none.
    /// </summary>
    /// <remarks>
    /// <b>Supply fonts covering everything the documents use, not only the script that failed.</b>
    /// They replace the host's own fallbacks rather than adding to them, so too few is worse than
    /// none - see <see cref="PdfFontOptions"/>.
    /// </remarks>
    public PdfFontOptions? Fonts { get; init; }
}
