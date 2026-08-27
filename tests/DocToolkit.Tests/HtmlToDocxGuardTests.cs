namespace DocToolkit.Tests;

/// <summary>
/// The argument guards on <see cref="HtmlToDocxConverter"/>'s Stream and file-path overloads.
/// </summary>
/// <remarks>
/// <b>Every test here exists because a mutation survived.</b> Putting
/// <c>HtmlToDocxConverter.cs</c> into the mutation scope on 2026-08-27 showed that deleting
/// <c>ThrowIfNull</c>, <c>options.Validate()</c> or <c>ThrowIfNullOrWhiteSpace</c> on these
/// overloads changed nothing any test could see.
///
/// <para><b>The discriminator is the exception TYPE and the parameter it names</b>, not merely that
/// something threw. Without the guard the call still fails — later, from inside the conversion,
/// wrapped as a <see cref="DocumentConversionException"/> or naming a parameter the caller never
/// passed. "It throws either way" is exactly the shape this repository has had to correct twice
/// before: the <c>PdfEditor</c> overloads that passed a cancellation suite only because a write
/// refused at the end, and the XLSX exporters whose Stream guards were indistinguishable from no
/// guards at all.</para>
///
/// <para>The <c>Stream</c> overloads' <c>destination</c> guards are NOT here — they belong to
/// <c>StreamOverloadTests</c>' theories, which cover every registered overload uniformly. This file
/// is only the arguments those theories do not vary.</para>
/// </remarks>
public class HtmlToDocxGuardTests
{
    private const string Html = "<p>x</p>";

    private static MemoryStream Sink() => new();

    // ---- html --------------------------------------------------------------------------------

    public static TheoryData<string> HtmlAcceptingOverloads() =>
    [
        "destination", "bool", "options", "page", "page+options", "file", "file+page",
    ];

    private static Task Invoke(string shape, string? html, RemoteImageOptions? options = null,
                               PageSetup? page = null, string? outputPath = null) => shape switch
                               {
                                   "destination" => HtmlToDocxConverter.ConvertAsync(html!, Sink()),
                                   "bool" => HtmlToDocxConverter.ConvertAsync(html!, false, Sink()),
                                   "options" => HtmlToDocxConverter.ConvertAsync(html!, options!, Sink()),
                                   "page" => HtmlToDocxConverter.ConvertAsync(html!, page!, Sink()),
                                   "page+options" => HtmlToDocxConverter.ConvertAsync(html!, page!, options!, Sink()),
                                   "file" => HtmlToDocxConverter.ConvertToFileAsync(html!, outputPath!),
                                   _ => HtmlToDocxConverter.ConvertToFileAsync(html!, page!, outputPath!),
                               };

    [Theory]
    [MemberData(nameof(HtmlAcceptingOverloads))]
    public async Task NullHtmlIsRefusedByName(string shape)
    {
        using var tmp = new TempFile();

        var ex = await Assert.ThrowsAsync<ArgumentNullException>(
            () => Invoke(shape, null, new RemoteImageOptions(), PageSetup.A4, tmp.Path));

        Assert.Equal("html", ex.ParamName);
    }

    // ---- options -----------------------------------------------------------------------------

    [Theory]
    [InlineData("options")]
    [InlineData("page+options")]
    public async Task NullOptionsAreRefusedByName(string shape)
    {
        var ex = await Assert.ThrowsAsync<ArgumentNullException>(
            () => Invoke(shape, Html, null, PageSetup.A4));

        Assert.Equal("options", ex.ParamName);
    }

    [Theory]
    [InlineData("options")]
    [InlineData("page+options")]
    public async Task InvalidOptionsAreRefusedUnwrapped(string shape)
    {
        // options.Validate() is its own guard, separate from the null check.
        //
        // This proves the refusal is UNWRAPPED - an ArgumentOutOfRangeException rather than a
        // DocumentConversionException - which means at or before the choke point. It does NOT
        // prove "before the conversion runs", and an earlier version of this comment claimed it
        // did, along with a claim that a bad timeout "reaches the fetch machinery". Measured: with
        // the entry-point Validate() deleted, BuildPackageAsync's own options?.Validate() throws
        // the identical exception, outside its try, before any parsing. TheFIRSTFaultyArgumentIsTheOneNamed
        // is what actually discriminates the two call sites.
        var invalid = new RemoteImageOptions { Timeout = TimeSpan.Zero };

        var ex = await Assert.ThrowsAsync<ArgumentOutOfRangeException>(
            () => Invoke(shape, Html, invalid, PageSetup.A4));

        Assert.Equal("Timeout", ex.ParamName);
    }

    // ---- page --------------------------------------------------------------------------------

    [Theory]
    [InlineData("page")]
    [InlineData("page+options")]
    [InlineData("file+page")]
    public async Task NullPageIsRefusedByName(string shape)
    {
        using var tmp = new TempFile();

        var ex = await Assert.ThrowsAsync<ArgumentNullException>(
            () => Invoke(shape, Html, new RemoteImageOptions(), null, tmp.Path));

        Assert.Equal("page", ex.ParamName);
    }

    // ---- outputPath --------------------------------------------------------------------------

    [Theory]
    [InlineData("file", "")]
    [InlineData("file", "   ")]
    [InlineData("file+page", "")]
    [InlineData("file+page", "   ")]
    public async Task ABlankOutputPathIsRefusedByName(string shape, string outputPath)
    {
        // The file-path overloads are in none of StreamOverloadTests' name lists - they are not
        // Stream overloads - so nothing covered these at all.
        var ex = await Assert.ThrowsAsync<ArgumentException>(
            () => Invoke(shape, Html, new RemoteImageOptions(), PageSetup.A4, outputPath));

        Assert.Equal("outputPath", ex.ParamName);
    }

    [Theory]
    [InlineData("file")]
    [InlineData("file+page")]
    public async Task ANullOutputPathIsRefusedByName(string shape)
    {
        var ex = await Assert.ThrowsAsync<ArgumentNullException>(
            () => Invoke(shape, Html, new RemoteImageOptions(), PageSetup.A4, null));

        Assert.Equal("outputPath", ex.ParamName);
    }
    // ---- cancellation ---------------------------------------------------------------------------

    [Theory]
    [InlineData("bool")]
    [InlineData("options")]
    [InlineData("page+options")]
    public async Task AnAlreadyCancelledTokenGivesOperationCanceled_NotTaskCanceled(string shape)
    {
        // The EXACT type is the assertion, and it is what discriminates. Measured: deleting
        // ct.ThrowIfCancellationRequested() from these overloads still throws - the token is
        // observed again deeper in the pipeline - but as a TaskCanceledException rather than an
        // OperationCanceledException. A caller catching the latter still catches the former, since
        // it derives from it, so "it throws either way" is true here and useless.
        //
        // What changes is what the guard PROMISES: refused before any work, rather than noticed
        // part-way through. Assert.ThrowsAsync matches the exact type, which is the only reason
        // this fails when the guard is deleted.
        using var cancelled = new CancellationTokenSource();
        await cancelled.CancelAsync();

        await Assert.ThrowsAsync<OperationCanceledException>(() => shape switch
        {
            // "page+options" was MISSING here, and the guard it covers was then labelled
            // equivalent - with a comment claiming BuildPackageAsync re-checks the token. It does
            // not: it re-checks html, page and options and nothing else. The label was wrong
            // because the test was incomplete, which is the worst way for one to be wrong.
            "bool" => HtmlToDocxConverter.ConvertAsync(Html, false, Sink(), cancelled.Token),
            "options" => HtmlToDocxConverter.ConvertAsync(
                Html, new RemoteImageOptions(), Sink(), cancelled.Token),
            _ => HtmlToDocxConverter.ConvertAsync(
                Html, PageSetup.A4, new RemoteImageOptions(), Sink(), cancelled.Token),
        });
    }
    // ---- guards fire in PARAMETER ORDER ---------------------------------------------------------

    /// <summary>
    /// With two faulty arguments, which one the caller is told about is STABLE — and pinning it is
    /// what makes each guard observable.
    /// </summary>
    /// <remarks>
    /// <b>These exist because nine guards were nearly excluded as equivalent, wrongly.</b> Each
    /// entry-point <c>ThrowIfNull</c> is repeated at the choke point in <c>BuildPackageAsync</c>,
    /// so deleting one is invisible <i>when the guarded argument is the only fault</i> — and every
    /// test at the time passed exactly one bad argument. A code review measured the two-fault case
    /// and found all nine observable: delete the <c>html</c> guard and a caller passing null html
    /// AND a null destination is told to check <c>destination</c>, an argument that is not the one
    /// they got wrong first.
    ///
    /// <para>That is a real contract rather than a trick: a caller fixing arguments one at a time
    /// should never be sent to an argument whose fault only became visible because an earlier guard
    /// was skipped.</para>
    ///
    /// <para><b>It is NOT "parameter order", and asserting that was the first version's mistake.</b>
    /// Most overloads report their first declared parameter, but
    /// <c>ConvertToFileAsync(html, outputPath, ct)</c> checks <c>outputPath</c> before delegating,
    /// so it reports that even though <c>html</c> is declared first. Measured, and pinned as it is
    /// rather than bent to a tidier rule — the point is that the answer is stable and changes when
    /// a guard is deleted, not that it follows any particular order.</para>
    ///
    /// <para><b>The near-miss is the lesson.</b> An equivalence label is a claim that no test can
    /// ever tell the difference, and it was reached here by measuring only inputs with a single
    /// fault. This repository has shipped a wrong equivalence label once before.</para>
    /// </remarks>
    [Theory]
    [MemberData(nameof(TwoFaultCases))]
    public async Task WhichArgumentIsNamedIsStable(string label, Func<Task> call, string expected)
    {
        Assert.NotEmpty(label);

        var ex = await Assert.ThrowsAnyAsync<ArgumentException>(call);

        Assert.Equal(expected, ex.ParamName);
    }

    public static TheoryData<string, Func<Task>, string> TwoFaultCases()
    {
        var badOptions = new RemoteImageOptions { Timeout = TimeSpan.Zero };
        var unwritable = new NonWritableForGuardTests();

        return new TheoryData<string, Func<Task>, string>
        {
            // html is declared first, so html is what a caller hears about.
            { "null html, null destination",
              () => HtmlToDocxConverter.ConvertAsync(null!, (Stream)null!), "html" },
            { "null html, unwritable destination",
              () => HtmlToDocxConverter.ConvertAsync(null!, unwritable), "html" },
            { "null html, null destination (bool overload)",
              () => HtmlToDocxConverter.ConvertAsync(null!, false, (Stream)null!), "html" },
            { "null html, null options",
              () => HtmlToDocxConverter.ConvertAsync(null!, (RemoteImageOptions)null!, Sink()), "html" },
            { "null html, null page",
              () => HtmlToDocxConverter.ConvertAsync(null!, (PageSetup)null!, badOptions, Sink()), "html" },
            { "null page, invalid options",
              () => HtmlToDocxConverter.ConvertAsync(Html, (PageSetup)null!, badOptions, Sink()), "page" },
            { "invalid options, null destination",
              () => HtmlToDocxConverter.ConvertAsync(Html, badOptions, (Stream)null!), "Timeout" },
            { "invalid options, null destination (page overload)",
              () => HtmlToDocxConverter.ConvertAsync(Html, PageSetup.A4, badOptions, (Stream)null!), "Timeout" },
            // The byte[]-returning overloads. A sweep of all 29 guard statements in the file
            // showed these surviving too - the first pass only reached the ones that take a
            // destination, because a two-fault case needs a second argument to make faulty and
            // "destination" was the one being reached for.
            { "null html, null page (byte[])",
              () => HtmlToDocxConverter.ConvertAsync(null!, (PageSetup)null!), "html" },
            { "null html, cancelled token (byte[] bool)",
              () => HtmlToDocxConverter.ConvertAsync(null!, false, Cancelled()), "html" },
            { "null html, null options (byte[])",
              () => HtmlToDocxConverter.ConvertAsync(null!, (RemoteImageOptions)null!), "html" },
            { "null html, null page (byte[] page+options)",
              () => HtmlToDocxConverter.ConvertAsync(null!, (PageSetup)null!, new RemoteImageOptions()), "html" },
            { "null page, invalid options (byte[])",
              () => HtmlToDocxConverter.ConvertAsync(Html, (PageSetup)null!, Bad()), "page" },
            { "invalid options, cancelled token (byte[])",
              () => HtmlToDocxConverter.ConvertAsync(Html, Bad(), Cancelled()), "Timeout" },

            // The stream overload whose html guard the first pass missed.
            { "null html, null destination (page overload)",
              () => HtmlToDocxConverter.ConvertAsync(null!, PageSetup.A4, (Stream)null!), "html" },
            // This overload has NO local page guard - page is checked only at the choke point -
            // so RequireWritable fires first. Pinned as measured rather than as expected.
            { "null page, null destination",
              () => HtmlToDocxConverter.ConvertAsync(Html, (PageSetup)null!, (Stream)null!), "destination" },

            // The asymmetric one: outputPath is guarded before the delegation that checks html.
            { "null html, blank outputPath",
              () => HtmlToDocxConverter.ConvertToFileAsync(null!, "   "), "outputPath" },

            // ...and its three-argument sibling, which does report html first.
            { "null html, blank outputPath (page overload)",
              () => HtmlToDocxConverter.ConvertToFileAsync(null!, PageSetup.A4, "   "), "html" },
        };
    }

    private static RemoteImageOptions Bad() => new() { Timeout = TimeSpan.Zero };

    private static CancellationToken Cancelled()
    {
        var source = new CancellationTokenSource();
        source.Cancel();
        return source.Token;
    }

    private sealed class NonWritableForGuardTests : MemoryStream
    {
        public override bool CanWrite => false;
    }
}
