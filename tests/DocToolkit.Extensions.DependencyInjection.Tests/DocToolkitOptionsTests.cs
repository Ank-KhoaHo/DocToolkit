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
}
