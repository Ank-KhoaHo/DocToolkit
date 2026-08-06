using DocToolkit.Extensions.DependencyInjection;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class DocToolkitOptionsTests
{
    [Fact]
    public void AllowRemoteImageDownload_DefaultsToFalse()
    {
        var options = new DocToolkitOptions();

        Assert.False(options.AllowRemoteImageDownload);
    }

    [Fact]
    public void AllowRemoteImageDownload_IsSettable()
    {
        var options = new DocToolkitOptions { AllowRemoteImageDownload = true };

        Assert.True(options.AllowRemoteImageDownload);
    }

    [Fact]
    public void RemoteImage_IsPresentAndRestrictiveByDefault()
    {
        // Restated here rather than deferred to the core package's own tests, because the value
        // this asserts is that *the DI defaults* are the restrictive ones. A future change that
        // handed DocToolkitOptions a pre-loosened RemoteImageOptions would leave core's tests
        // green and silently opt every consumer of AddDocToolkit into a wider reach.
        var options = new DocToolkitOptions();

        Assert.NotNull(options.RemoteImage);
        Assert.False(options.RemoteImage.AllowPrivateAddresses);
        Assert.Equal(TimeSpan.FromSeconds(10), options.RemoteImage.Timeout);
        Assert.Equal(5 * 1024 * 1024, options.RemoteImage.MaxBytesPerImage);
        Assert.Empty(options.RemoteImage.AllowedHosts);
    }

    [Fact]
    public void RemoteImage_IsConfiguredInPlace_NotReplaced()
    {
        // The property is get-only by design: a settable one lets a caller drop in an object that
        // missed one of the restrictive defaults. Mutating in place cannot lose a default that was
        // not deliberately changed.
        var options = new DocToolkitOptions();
        var original = options.RemoteImage;

        options.RemoteImage.AllowedHosts.Add("cdn.example.com");
        options.RemoteImage.Timeout = TimeSpan.FromSeconds(3);

        Assert.Same(original, options.RemoteImage);
        Assert.Contains("cdn.example.com", options.RemoteImage.AllowedHosts);
        Assert.Equal(TimeSpan.FromSeconds(3), options.RemoteImage.Timeout);
    }
}
