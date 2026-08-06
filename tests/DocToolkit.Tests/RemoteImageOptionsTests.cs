using Xunit;

namespace DocToolkit.Tests;

public class RemoteImageOptionsTests
{
    [Fact]
    public void Defaults_AreTheSafeOnes()
    {
        var options = new RemoteImageOptions();

        Assert.Equal(TimeSpan.FromSeconds(10), options.Timeout);
        Assert.Equal(5 * 1024 * 1024, options.MaxBytesPerImage);
        Assert.Empty(options.AllowedHosts);
        Assert.False(options.AllowPrivateAddresses);
    }

    [Fact]
    public void AllowedHosts_IgnoresCase()
    {
        var options = new RemoteImageOptions();
        options.AllowedHosts.Add("CDN.Example.COM");

        Assert.Contains("cdn.example.com", options.AllowedHosts);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Validate_RejectsANonPositiveTimeout(int seconds)
    {
        var options = new RemoteImageOptions { Timeout = TimeSpan.FromSeconds(seconds) };

        Assert.Throws<ArgumentOutOfRangeException>(() => options.Validate());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Validate_RejectsANonPositiveSizeCap(long bytes)
    {
        var options = new RemoteImageOptions { MaxBytesPerImage = bytes };

        Assert.Throws<ArgumentOutOfRangeException>(() => options.Validate());
    }

    [Fact]
    public void Validate_RejectsABlankAllowedHost()
    {
        var options = new RemoteImageOptions();
        options.AllowedHosts.Add(" ");

        Assert.Throws<ArgumentException>(() => options.Validate());
    }

    [Fact]
    public void Validate_AcceptsTheDefaults()
    {
        new RemoteImageOptions().Validate();   // must not throw
    }
}
