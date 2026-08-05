using System.Net;
using Xunit;

namespace DocToolkit.Tests;

public class GuardedResourceLoaderTests
{
    [Theory]
    [InlineData("127.0.0.1")]        // loopback
    [InlineData("::1")]              // loopback, v6
    [InlineData("10.0.0.5")]         // RFC1918
    [InlineData("172.16.0.1")]       // RFC1918
    [InlineData("192.168.1.1")]      // RFC1918
    [InlineData("169.254.169.254")]  // link-local - the cloud metadata endpoint
    [InlineData("fe80::1")]          // link-local, v6
    [InlineData("fc00::1")]          // unique-local, v6
    [InlineData("::ffff:10.0.0.5")]  // v4-mapped RFC1918 - the obvious bypass
    public void IsBlockedAddress_RefusesEveryPrivateForm(string address)
    {
        Assert.True(GuardedResourceLoader.IsBlockedAddress(IPAddress.Parse(address)));
    }

    [Theory]
    [InlineData("93.184.216.34")]
    [InlineData("2606:2800:220:1:248:1893:25c8:1946")]
    public void IsBlockedAddress_AllowsPublicAddresses(string address)
    {
        Assert.False(GuardedResourceLoader.IsBlockedAddress(IPAddress.Parse(address)));
    }

    [Theory]
    [InlineData("file")]
    [InlineData("ftp")]
    [InlineData("data")]
    [InlineData("gopher")]
    public void SupportsProtocol_RefusesEverythingButHttpAndHttps(string protocol)
    {
        var loader = new GuardedResourceLoader(new RemoteImageOptions());

        Assert.False(loader.SupportsProtocol(protocol));
    }

    [Theory]
    [InlineData("http")]
    [InlineData("https")]
    public void SupportsProtocol_AcceptsHttpAndHttps(string protocol)
    {
        var loader = new GuardedResourceLoader(new RemoteImageOptions());

        Assert.True(loader.SupportsProtocol(protocol));
    }
}
