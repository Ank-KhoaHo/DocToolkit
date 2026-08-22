using System.Threading;
using System.Threading.Tasks;
using DocToolkit;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// The token must reach the DOCX to PDF half of an HTML to PDF conversion, not only the first half.
/// </summary>
/// <remarks>
/// <b>AN ALREADY-CANCELLED TOKEN CANNOT PROVE THIS, and that is the whole reason these tests
/// exist.</b> Every public overload already passes the token to
/// <c>HtmlToDocxConverter.ConvertAsync</c>, which refuses immediately - so a suite that only ever
/// cancels up front goes green whether or not the second stage observes anything. That is exactly
/// the trap this repository recorded against <c>PdfEditor</c>, where seven overloads passed the
/// cancellation suite because <c>destination.WriteAsync</c> happened to refuse at the end.
///
/// <para>So these cancel BETWEEN the stages, from inside the conversion delegate. The render is
/// the expensive half and the repair loop can run it several times, so a token that cannot
/// interrupt it does not do the thing its documentation promises.</para>
/// </remarks>
public class HtmlForPdfCancellationTests
{
    private const string Html = "<html><body><p>Cancelling between the stages.</p></body></html>";

    [Fact]
    public async Task RenderAsync_CancelledAfterTheDocxStage_DoesNotGoOnToRender()
    {
        using var cts = new CancellationTokenSource();

        await Assert.ThrowsAnyAsync<System.OperationCanceledException>(() =>
            HtmlForPdf.RenderAsync(
                Html,
                async h =>
                {
                    // The first stage succeeds, then the caller cancels. Deterministic: no timer,
                    // no wall clock, no race - the cancellation happens at a known point.
                    var docx = await HtmlToDocxConverter.ConvertAsync(h, PageSetup.A4, CancellationToken.None);
                    await cts.CancelAsync();
                    return docx;
                },
                fonts: null,
                ct: cts.Token));
    }

    [Fact]
    public async Task RenderAsync_ANotCancelledToken_StillProducesAPdf()
    {
        // The positive control. Without it, a RenderAsync that threw unconditionally would pass
        // the test above and look correct.
        var pdf = await HtmlForPdf.RenderAsync(
            Html,
            h => HtmlToDocxConverter.ConvertAsync(h, PageSetup.A4, CancellationToken.None),
            fonts: null,
            ct: CancellationToken.None);

        Assert.NotEmpty(pdf);
        Assert.Equal("%PDF", System.Text.Encoding.ASCII.GetString(pdf, 0, 4));
    }

    [Fact]
    public async Task RenderAsync_AnAlreadyCancelledToken_DoesNotEvenStartTheFirstStage()
    {
        // Discriminates the check at the TOP of the retry loop, which the test above cannot: there
        // the token is still live when the loop is entered.
        //
        // Today HtmlToDocxConverter would refuse this token itself, so the OUTCOME is the same
        // either way - which is exactly why the assertion is on the delegate never being CALLED
        // rather than on the exception. A guard that works only because its delegate happens to
        // refuse is the failure this file's remarks describe.
        using var cts = new CancellationTokenSource();
        await cts.CancelAsync();

        var started = 0;

        await Assert.ThrowsAnyAsync<System.OperationCanceledException>(() =>
            HtmlForPdf.RenderAsync(
                Html,
                h =>
                {
                    Interlocked.Increment(ref started);
                    return HtmlToDocxConverter.ConvertAsync(h, PageSetup.A4, CancellationToken.None);
                },
                fonts: null,
                ct: cts.Token));

        Assert.Equal(0, started);
    }
}
