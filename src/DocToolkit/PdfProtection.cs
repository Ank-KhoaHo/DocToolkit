namespace DocToolkit;

/// <summary>
/// The cipher used when a PDF is encrypted.
/// </summary>
/// <remarks>
/// <b>This is a compatibility decision, not a "more is better" one.</b> Both options below are
/// sound ciphers; they differ in which readers can open the result.
/// </remarks>
public enum PdfEncryptionStrength
{
    /// <summary>
    /// AES-128 (PDF 1.6). <b>The default</b>, because every PDF reader still in service can open
    /// it, and it is what the underlying library produces when nothing is chosen.
    /// </summary>
    Aes128 = 0,

    /// <summary>
    /// AES-256 (PDF 2.0). Stronger, but it needs a reader that understands PDF 2.0 encryption —
    /// Acrobat X and later. An older reader cannot open the file at all, so choose this only when
    /// you know what will open the document.
    /// </summary>
    Aes256 = 1,
}

/// <summary>
/// The password and permissions applied by <see cref="PdfEditor.Protect(byte[], PdfProtection)"/>.
/// </summary>
/// <remarks>
/// <b>The two passwords do different jobs, and mixing them up is the usual mistake.</b>
///
/// <list type="bullet">
/// <item><description>
/// <see cref="UserPassword"/> is needed to <b>open</b> the document. Without it, the file cannot be
/// read at all.
/// </description></item>
/// <item><description>
/// <see cref="OwnerPassword"/> is needed to <b>change</b> the restrictions. A document with only an
/// owner password opens for anyone, and the permissions below are what a reader is asked to honour.
/// </description></item>
/// </list>
///
/// <b>Permissions are a request to the reader, not a guarantee.</b> Measured 2026-08-16: a document
/// carrying only an owner password opens with no password at all — that is the PDF specification
/// working as designed, not a defect here. A cooperative reader greys out printing; an
/// uncooperative one need not. **If content must not be read, use <see cref="UserPassword"/>** —
/// that is enforced by cryptography rather than by convention.
///
/// <b>Every permission defaults to allowed</b>, so adding a password does not silently forbid
/// printing a document that could be printed before.
/// </remarks>
public sealed class PdfProtection
{
    /// <summary>
    /// The password required to <b>open</b> the document. <see langword="null"/> or empty means the
    /// document opens without one.
    /// </summary>
    public string? UserPassword { get; init; }

    /// <summary>
    /// The password required to <b>change</b> the permissions below. <see langword="null"/> or
    /// empty means none is set.
    /// </summary>
    public string? OwnerPassword { get; init; }

    /// <summary>Printing is permitted. Defaults to <see langword="true"/>.</summary>
    public bool AllowPrinting { get; init; } = true;

    /// <summary>
    /// Full-resolution printing is permitted. Defaults to <see langword="true"/>. Ignored by
    /// readers when <see cref="AllowPrinting"/> is <see langword="false"/>.
    /// </summary>
    public bool AllowHighQualityPrinting { get; init; } = true;

    /// <summary>Copying text and graphics out is permitted. Defaults to <see langword="true"/>.</summary>
    public bool AllowCopying { get; init; } = true;

    /// <summary>Changing the content is permitted. Defaults to <see langword="true"/>.</summary>
    public bool AllowModification { get; init; } = true;

    /// <summary>Adding or editing annotations is permitted. Defaults to <see langword="true"/>.</summary>
    public bool AllowAnnotations { get; init; } = true;

    /// <summary>Filling in form fields is permitted. Defaults to <see langword="true"/>.</summary>
    public bool AllowFormFilling { get; init; } = true;

    /// <summary>
    /// Inserting, rotating and deleting pages is permitted. Defaults to <see langword="true"/>.
    /// </summary>
    public bool AllowAssembly { get; init; } = true;

    /// <summary>
    /// The cipher to use. Defaults to <see cref="PdfEncryptionStrength.Aes128"/> — see that member
    /// for why the stronger option is not the default.
    /// </summary>
    public PdfEncryptionStrength Strength { get; init; } = PdfEncryptionStrength.Aes128;
}
