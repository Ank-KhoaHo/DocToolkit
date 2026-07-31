using DocToolkit.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Options;
using Xunit;

namespace DocToolkit.Extensions.DependencyInjection.Tests;

public class ServiceCollectionExtensionsTests
{
    /// <summary>
    /// How long to wait after a conversion returns before declaring "nothing connected". Matches
    /// AirGapGuardTests.SettleWindow in the core test project - long enough to catch a fetch that
    /// was fired and abandoned late, even though it costs test suite time on every run.
    /// </summary>
    private static readonly TimeSpan SettleWindow = TimeSpan.FromMilliseconds(750);

    [Fact]
    public void AddDocToolkit_ResolvesAllSixInterfaces()
    {
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();

        Assert.NotNull(provider.GetRequiredService<IHtmlToDocxConverter>());
        Assert.NotNull(provider.GetRequiredService<IDocxToPdfConverter>());
        Assert.NotNull(provider.GetRequiredService<IHtmlToPdfConverter>());
        Assert.NotNull(provider.GetRequiredService<IDocxEditor>());
        Assert.NotNull(provider.GetRequiredService<IWorkbookEditor>());
        Assert.NotNull(provider.GetRequiredService<IPresentationEditor>());
    }

    [Fact]
    public void AddDocToolkit_RegistersEachInterfaceAsASingleton()
    {
        var services = new ServiceCollection().AddDocToolkit();

        var registeredTypes = new[]
        {
            typeof(IHtmlToDocxConverter), typeof(IDocxToPdfConverter), typeof(IHtmlToPdfConverter),
            typeof(IDocxEditor), typeof(IWorkbookEditor), typeof(IPresentationEditor),
        };

        foreach (var serviceType in registeredTypes)
        {
            var descriptor = Assert.Single(services, d => d.ServiceType == serviceType);
            Assert.Equal(ServiceLifetime.Singleton, descriptor.Lifetime);
        }
    }

    [Fact]
    public void AddDocToolkit_WithNoConfigure_DefaultsToNoRemoteImageDownload()
    {
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();

        Assert.False(provider.GetRequiredService<IOptions<DocToolkitOptions>>().Value.AllowRemoteImageDownload);
    }

    [Fact]
    public void AddDocToolkit_WithConfigure_MakesTheValueObservableViaIOptions()
    {
        var provider = new ServiceCollection()
            .AddDocToolkit(o => o.AllowRemoteImageDownload = true)
            .BuildServiceProvider();

        Assert.True(provider.GetRequiredService<IOptions<DocToolkitOptions>>().Value.AllowRemoteImageDownload);
    }

    [Fact]
    public async Task AddDocToolkit_WithAllowRemoteImageDownloadFalse_NeverConnectsOutbound()
    {
        using var probe = new LoopbackProbe();
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();
        var sut = provider.GetRequiredService<IHtmlToDocxConverter>();

        await sut.ConvertAsync($"<img src=\"{probe.ImageUrl}\">");
        await Task.Delay(SettleWindow);

        Assert.Equal(0, probe.Connections);
    }

    [Fact]
    public async Task AddDocToolkit_WithAllowRemoteImageDownloadTrue_DoesConnectOutbound()
    {
        using var probe = new LoopbackProbe();
        var provider = new ServiceCollection()
            .AddDocToolkit(o => o.AllowRemoteImageDownload = true)
            .BuildServiceProvider();
        var sut = provider.GetRequiredService<IHtmlToDocxConverter>();

        try
        {
            await sut.ConvertAsync($"<img src=\"{probe.ImageUrl}\">");
        }
        catch (DocToolkit.DocumentConversionException)
        {
            // The connection is what's under test here, not a fully successful conversion.
        }

        Assert.True(await probe.WaitForConnectionAsync(TimeSpan.FromSeconds(5)));
    }

    [Fact]
    public async Task AddDocToolkit_WithAllowRemoteImageDownloadFalse_HtmlToPdfNeverConnectsOutbound()
    {
        using var probe = new LoopbackProbe();
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();
        var sut = provider.GetRequiredService<IHtmlToPdfConverter>();

        await sut.ConvertAsync($"<img src=\"{probe.ImageUrl}\">");
        await Task.Delay(SettleWindow);

        Assert.Equal(0, probe.Connections);
    }

    [Fact]
    public async Task AddDocToolkit_WithAllowRemoteImageDownloadTrue_HtmlToPdfDoesConnectOutbound()
    {
        using var probe = new LoopbackProbe();
        var provider = new ServiceCollection()
            .AddDocToolkit(o => o.AllowRemoteImageDownload = true)
            .BuildServiceProvider();
        var sut = provider.GetRequiredService<IHtmlToPdfConverter>();

        try
        {
            await sut.ConvertAsync($"<img src=\"{probe.ImageUrl}\">");
        }
        catch (DocToolkit.DocumentConversionException)
        {
            // The connection is what's under test here, not a fully successful conversion.
        }

        Assert.True(await probe.WaitForConnectionAsync(TimeSpan.FromSeconds(5)));
    }
}
