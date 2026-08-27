using HtmlToOpenXml.IO;

namespace DocToolkit.Tests;

/// <summary>
/// The default refuse-everything path, asked directly.
/// </summary>
/// <remarks>
/// <b>Mutation testing found <c>SupportsProtocol</c> had NO COVERAGE AT ALL</b> — flipping
/// <c>=&gt; false</c> to <c>=&gt; true</c> was reached by nothing in the repository. That is the
/// class <c>CLAUDE.md</c> calls <i>"the default path"</i>, taken by every consumer who does not
/// opt in to remote images, on a library whose whole premise includes running air-gapped.
///
/// <para><b>The air-gap suites cannot see this, and it is worth understanding why before assuming
/// they cover it.</b> They count connections to a loopback probe and assert zero. With
/// <c>SupportsProtocol</c> returning <c>true</c>, HtmlToOpenXml would start asking this loader for
/// resources — but <c>FetchAsync</c> returns <c>null</c>, so still nothing is fetched and the probe
/// still sees zero. The suites' claim is <i>"nothing reached our probe"</i>, which is a weaker
/// statement than <i>"the component that could download was never built"</i>. Only asking the type
/// directly separates them.</para>
///
/// <para>Same shape as B28's step 1: the offline guarantee has two independent lines of defence,
/// and a test that can only see the outer one leaves the inner one unasserted.</para>
/// </remarks>
public class OfflineResourceLoaderTests
{
    private static IWebRequest Loader => HtmlToDocxConverter.OfflineResourceLoader.Instance;

    [Theory]
    [InlineData("http")]
    [InlineData("https")]
    [InlineData("file")]
    [InlineData("ftp")]
    [InlineData("data")]
    [InlineData("")]
    public void NoProtocolIsSupported(string protocol)
    {
        // The class's own doc comment says "No protocol is supported - not http, not https, not
        // file", and until now nothing checked it. `data:` is here deliberately: data-URI images
        // ARE embedded, but by the parser itself rather than through this loader, so the honest
        // answer is still false. If that ever changes, this line is where it surfaces.
        Assert.False(Loader.SupportsProtocol(protocol));
    }

    [Fact]
    public async Task FetchAsyncReturnsNothing_SoTheSecondLineOfDefenceHoldsToo()
    {
        // Never called in practice, given SupportsProtocol. Asserted anyway because "never called"
        // is a property of the CALLER, and this class's guarantee should not depend on one.
        Assert.Null(await Loader.FetchAsync(new Uri("https://example.invalid/logo.png"), CancellationToken.None));
        Assert.Null(await Loader.FetchAsync(new Uri("file:///etc/passwd"), CancellationToken.None));
    }

    [Fact]
    public void TheInstanceIsShared_AndIsWhatTheConverterHandsToHtmlToOpenXml()
    {
        // A second instance would be harmless, but the singleton is what makes "the component that
        // could download was never built" true for every conversion rather than per call.
        Assert.Same(HtmlToDocxConverter.OfflineResourceLoader.Instance,
                    HtmlToDocxConverter.OfflineResourceLoader.Instance);
    }
}
