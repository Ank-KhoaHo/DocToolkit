namespace DocToolkit;

/// <summary>
/// How badly a conversion lost something.
/// </summary>
/// <remarks>
/// DocToolkit's own enum rather than the one the underlying converter uses. Re-exporting an
/// upstream enum would put a type this library does not control into its permanent public
/// surface, so an upstream major that renamed a member would become a breaking change for every
/// consumer here.
/// </remarks>
public enum ConversionLossKind
{
    /// <summary>Reported for information; nothing was lost.</summary>
    None = 0,

    /// <summary>The content survived, but not exactly — a style or layout was approximated.</summary>
    Approximation = 1,

    /// <summary>Content the target format cannot express was left out.</summary>
    Omission = 2,

    /// <summary>
    /// Something could not be converted at all. The call still returned, so the output is usable;
    /// this says part of it is missing rather than that the conversion failed. A conversion that
    /// fails outright throws <see cref="DocumentConversionException"/> instead.
    /// </summary>
    Failure = 3,
}

/// <summary>
/// One thing a conversion could not carry across faithfully.
/// </summary>
/// <param name="Code">A short stable identifier from the underlying converter, for grouping.</param>
/// <param name="Message">A human-readable sentence describing what happened.</param>
/// <param name="Kind">How severe the loss was.</param>
public sealed record ConversionWarning(string Code, string Message, ConversionLossKind Kind);

/// <summary>
/// A converted document together with everything the conversion could not carry across.
/// </summary>
/// <remarks>
/// Returned only by the <c>ConvertWithReport</c> overloads. The plain <c>Convert</c> overloads are
/// unchanged and still return the value on its own — this type is opt-in, and adding it broke no
/// existing caller.
///
/// <b>There is no <c>Succeeded</c> property.</b> The underlying converters have one, but a
/// DocToolkit conversion that did not succeed throws <see cref="DocumentConversionException"/>, so
/// it would be <see langword="true"/> on every instance a caller could ever hold — a constant
/// dressed as data.
/// </remarks>
/// <typeparam name="T">The converted document's type.</typeparam>
public sealed class ConversionResult<T>
{
    internal ConversionResult(T value, IReadOnlyList<ConversionWarning> warnings)
    {
        Value = value;
        Warnings = warnings;
    }

    /// <summary>The converted document. Usable whatever <see cref="HasLoss"/> says.</summary>
    public T Value { get; }

    /// <summary>
    /// Everything the conversion could not carry across, in the order the converter reported it.
    /// Empty when nothing was lost.
    /// </summary>
    public IReadOnlyList<ConversionWarning> Warnings { get; }

    /// <summary>
    /// Whether anything was actually lost — that is, whether <see cref="Warnings"/> holds an entry
    /// whose <see cref="ConversionWarning.Kind"/> is not <see cref="ConversionLossKind.None"/>.
    /// </summary>
    /// <remarks>
    /// Derived from <see cref="Warnings"/> rather than copied from the underlying converter's own
    /// flag, so the two can never disagree in front of a caller. A <c>HasLoss</c> of
    /// <see langword="true"/> beside an empty warning list produces bug reports nobody can
    /// reproduce.
    /// </remarks>
    public bool HasLoss
    {
        get
        {
            for (var i = 0; i < Warnings.Count; i++)
            {
                if (Warnings[i].Kind is not ConversionLossKind.None) return true;
            }

            return false;
        }
    }
}
