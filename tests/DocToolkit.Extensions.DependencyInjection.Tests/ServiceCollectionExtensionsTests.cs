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
    public async Task AddDocToolkit_ResolvedWorkbookEditor_SheetNamesAndReadSheetMatchTheStaticApi()
    {
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();
        var sut = provider.GetRequiredService<IWorkbookEditor>();

        var xlsx = DocToolkit.WorkbookEditor.Create("Sales", new object?[][]
        {
            new object?[] { "Region", "Total" },
            new object?[] { "North", 1200 },
        });

        Assert.Equal(DocToolkit.WorkbookEditor.SheetNames(xlsx), sut.SheetNames(xlsx));
        Assert.Equal(
            await DocToolkit.WorkbookEditor.SheetNamesAsync(new MemoryStream(xlsx)),
            await sut.SheetNamesAsync(new MemoryStream(xlsx)));

        Assert.Equal(DocToolkit.WorkbookEditor.ReadSheet(xlsx, "Sales"), sut.ReadSheet(xlsx, "Sales"));
        Assert.Equal(
            await DocToolkit.WorkbookEditor.ReadSheetAsync(new MemoryStream(xlsx), "Sales"),
            await sut.ReadSheetAsync(new MemoryStream(xlsx), "Sales"));
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
            .AddDocToolkit(o =>
            {
                o.AllowRemoteImageDownload = true;

                // Core 0.8.0's guard refuses loopback, private and link-local addresses by
                // default, which would refuse this probe too. Naming that escape hatch here is
                // what keeps this test discriminating - and this test is what stops every
                // Assert.Equal(0, probe.Connections) in this file from passing vacuously,
                // because nothing connected rather than because connecting was blocked.
                o.RemoteImage.AllowPrivateAddresses = true;
            })
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
    public async Task AddDocToolkit_WithAllowRemoteImageDownloadFalse_StreamOverloadNeverConnectsOutbound()
    {
        using var probe = new LoopbackProbe();
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();
        var sut = provider.GetRequiredService<IHtmlToDocxConverter>();

        using var destination = new MemoryStream();
        await sut.ConvertAsync($"<img src=\"{probe.ImageUrl}\">", destination);
        await Task.Delay(SettleWindow);

        Assert.Equal(0, probe.Connections);
    }

    [Fact]
    public async Task AddDocToolkit_WithAllowRemoteImageDownloadTrue_StreamOverloadDoesConnectOutbound()
    {
        using var probe = new LoopbackProbe();
        var provider = new ServiceCollection()
            .AddDocToolkit(o =>
            {
                o.AllowRemoteImageDownload = true;

                // Core 0.8.0's guard refuses loopback, private and link-local addresses by
                // default, which would refuse this probe too. Naming that escape hatch here is
                // what keeps this test discriminating - and this test is what stops every
                // Assert.Equal(0, probe.Connections) in this file from passing vacuously,
                // because nothing connected rather than because connecting was blocked.
                o.RemoteImage.AllowPrivateAddresses = true;
            })
            .BuildServiceProvider();
        var sut = provider.GetRequiredService<IHtmlToDocxConverter>();

        try
        {
            using var destination = new MemoryStream();
            await sut.ConvertAsync($"<img src=\"{probe.ImageUrl}\">", destination);
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
            .AddDocToolkit(o =>
            {
                o.AllowRemoteImageDownload = true;

                // Core 0.8.0's guard refuses loopback, private and link-local addresses by
                // default, which would refuse this probe too. Naming that escape hatch here is
                // what keeps this test discriminating - and this test is what stops every
                // Assert.Equal(0, probe.Connections) in this file from passing vacuously,
                // because nothing connected rather than because connecting was blocked.
                o.RemoteImage.AllowPrivateAddresses = true;
            })
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

    [Fact]
    public async Task AddDocToolkit_WithAllowRemoteImageDownloadFalse_HtmlToPdfStreamOverloadNeverConnectsOutbound()
    {
        using var probe = new LoopbackProbe();
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();
        var sut = provider.GetRequiredService<IHtmlToPdfConverter>();

        using var destination = new MemoryStream();
        await sut.ConvertAsync($"<img src=\"{probe.ImageUrl}\">", destination);
        await Task.Delay(SettleWindow);

        Assert.Equal(0, probe.Connections);
    }

    [Fact]
    public async Task AddDocToolkit_WithAllowRemoteImageDownloadTrue_HtmlToPdfStreamOverloadDoesConnectOutbound()
    {
        using var probe = new LoopbackProbe();
        var provider = new ServiceCollection()
            .AddDocToolkit(o =>
            {
                o.AllowRemoteImageDownload = true;

                // Core 0.8.0's guard refuses loopback, private and link-local addresses by
                // default, which would refuse this probe too. Naming that escape hatch here is
                // what keeps this test discriminating - and this test is what stops every
                // Assert.Equal(0, probe.Connections) in this file from passing vacuously,
                // because nothing connected rather than because connecting was blocked.
                o.RemoteImage.AllowPrivateAddresses = true;
            })
            .BuildServiceProvider();
        var sut = provider.GetRequiredService<IHtmlToPdfConverter>();

        try
        {
            using var destination = new MemoryStream();
            await sut.ConvertAsync($"<img src=\"{probe.ImageUrl}\">", destination);
        }
        catch (DocToolkit.DocumentConversionException)
        {
            // The connection is what's under test here, not a fully successful conversion.
        }

        Assert.True(await probe.WaitForConnectionAsync(TimeSpan.FromSeconds(5)));
    }

    // The four DoesConnectOutbound tests above are, between them, the proof that RemoteImage
    // actually reaches the core converter rather than being an inert property: a service that
    // ignored it and kept passing the bool would send core `new RemoteImageOptions()`, whose
    // defaults block loopback, and all four would fail. This test pins a second field for the
    // same reason - AllowPrivateAddresses alone could be threaded through by accident, a whole
    // options object being passed cannot.

    [Fact]
    public async Task AddDocToolkit_RemoteImageAllowedHosts_NarrowsWhatIsFetched()
    {
        using var probe = new LoopbackProbe();
        var provider = new ServiceCollection()
            .AddDocToolkit(o =>
            {
                o.AllowRemoteImageDownload = true;
                o.RemoteImage.AllowPrivateAddresses = true;   // the address check would refuse it...

                // ...but the allow-list names a host this probe is not, so nothing should connect.
                // Same registration as the tests above, one field different, opposite outcome.
                o.RemoteImage.AllowedHosts.Add("images.example.invalid");
            })
            .BuildServiceProvider();
        var sut = provider.GetRequiredService<IHtmlToDocxConverter>();

        await sut.ConvertAsync($"<img src=\"{probe.ImageUrl}\">");
        await Task.Delay(SettleWindow);

        Assert.Equal(0, probe.Connections);
    }
}
