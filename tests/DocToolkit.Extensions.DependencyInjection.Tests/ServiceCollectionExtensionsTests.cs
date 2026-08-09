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
    public void AddDocToolkit_ResolvesAllTenInterfaces()
    {
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();

        Assert.NotNull(provider.GetRequiredService<IHtmlToDocxConverter>());
        Assert.NotNull(provider.GetRequiredService<IDocxToPdfConverter>());
        Assert.NotNull(provider.GetRequiredService<IHtmlToPdfConverter>());
        Assert.NotNull(provider.GetRequiredService<IDocxEditor>());
        Assert.NotNull(provider.GetRequiredService<IWorkbookEditor>());
        Assert.NotNull(provider.GetRequiredService<IPresentationEditor>());
        Assert.NotNull(provider.GetRequiredService<IXlsxToPdfConverter>());
        Assert.NotNull(provider.GetRequiredService<IPptxToPdfConverter>());
        Assert.NotNull(provider.GetRequiredService<IDocxToHtmlConverter>());
        Assert.NotNull(provider.GetRequiredService<IDocxToMarkdownConverter>());
        Assert.NotNull(provider.GetRequiredService<IPdfEditor>());
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

        using var namesStaticSource = new MemoryStream(xlsx);
        using var namesWrapperSource = new MemoryStream(xlsx);

        Assert.Equal(DocToolkit.WorkbookEditor.SheetNames(xlsx), sut.SheetNames(xlsx));
        Assert.Equal(
            await DocToolkit.WorkbookEditor.SheetNamesAsync(namesStaticSource),
            await sut.SheetNamesAsync(namesWrapperSource));

        using var sheetStaticSource = new MemoryStream(xlsx);
        using var sheetWrapperSource = new MemoryStream(xlsx);

        Assert.Equal(DocToolkit.WorkbookEditor.ReadSheet(xlsx, "Sales"), sut.ReadSheet(xlsx, "Sales"));
        Assert.Equal(
            await DocToolkit.WorkbookEditor.ReadSheetAsync(sheetStaticSource, "Sales"),
            await sut.ReadSheetAsync(sheetWrapperSource, "Sales"));
    }

    [Fact]
    public void AddDocToolkit_RegistersEachInterfaceAsASingleton()
    {
        var services = new ServiceCollection().AddDocToolkit();

        var registeredTypes = new[]
        {
            typeof(IHtmlToDocxConverter), typeof(IDocxToPdfConverter), typeof(IHtmlToPdfConverter),
            typeof(IDocxEditor), typeof(IWorkbookEditor), typeof(IPresentationEditor),
            typeof(IXlsxToPdfConverter), typeof(IPptxToPdfConverter),
            typeof(IDocxToHtmlConverter), typeof(IDocxToMarkdownConverter),
            typeof(IPdfEditor),
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

    // The mirror going stale is this package's most-repeated defect - seven times now. Resolving
    // an interface proves it was REGISTERED; these prove each new member actually reaches the
    // static method behind it, which is the part that has silently drifted before.
    [Fact]
    public void ResolvedConverters_DelegateToTheStaticApi()
    {
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();

        byte[] docx = DocToolkit.DocxEditor.Create(new[] { DocToolkit.DocxBlock.Heading("Report", 1) });
        byte[] xlsx = DocToolkit.WorkbookEditor.Create("Sales", new[] { new object?[] { "A", 1 } });
        byte[] pptx = DocToolkit.PresentationEditor.Create(new[] { DocToolkit.PptxSlide.Titled("T", "B") });

        Assert.Contains("<h1", provider.GetRequiredService<IDocxToHtmlConverter>().Convert(docx),
            StringComparison.OrdinalIgnoreCase);
        Assert.Contains("# Report", provider.GetRequiredService<IDocxToMarkdownConverter>().Convert(docx),
            StringComparison.Ordinal);
        Assert.NotEmpty(provider.GetRequiredService<IXlsxToPdfConverter>().Convert(xlsx));
        Assert.NotEmpty(provider.GetRequiredService<IPptxToPdfConverter>().Convert(pptx));
    }

    // PageSetup is the A9 mirror. Letter is asserted rather than the A4 default precisely because
    // A4 would pass even if the parameter were dropped on the way through.
    [Fact]
    public async Task ResolvedServices_HonourThePageSetupTheyAreGiven()
    {
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();
        var blocks = new[] { DocToolkit.DocxBlock.Paragraph("Hello.") };

        byte[] viaInterface = provider.GetRequiredService<IDocxEditor>()
            .Create(blocks, DocToolkit.PageSetup.Letter);
        byte[] viaStatic = DocToolkit.DocxEditor.Create(blocks, DocToolkit.PageSetup.Letter);
        Assert.Equal(DocToolkit.DocxEditor.ExtractText(viaStatic),
            DocToolkit.DocxEditor.ExtractText(viaInterface));

        byte[] pdf = await provider.GetRequiredService<IHtmlToPdfConverter>()
            .ConvertAsync("<p>x</p>", DocToolkit.PageSetup.Letter);
        Assert.NotEmpty(pdf);
    }

    // Every Stream overload on the newly mirrored interfaces. The coverage gate caught these
    // missing on the first CI run - the DI package is held at 100% precisely because it is pure
    // delegation, so an uncovered member IS an untested one and the fix is always a short test.
    [Fact]
    public async Task ResolvedServices_StreamOverloadsDelegateToo()
    {
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();
        var blocks = new[] { DocToolkit.DocxBlock.Heading("Report", 1) };
        byte[] docx = DocToolkit.DocxEditor.Create(blocks);
        byte[] xlsx = DocToolkit.WorkbookEditor.Create("Sales", new[] { new object?[] { "A", 1 } });
        byte[] pptx = DocToolkit.PresentationEditor.Create(new[] { DocToolkit.PptxSlide.Titled("T", "B") });

        using (var destination = new MemoryStream())
        {
            await provider.GetRequiredService<IDocxEditor>()
                .CreateAsync(blocks, DocToolkit.PageSetup.Letter, destination);
            Assert.NotEmpty(destination.ToArray());
        }

        using (var source = new MemoryStream(docx))
            Assert.Contains("<h1", await provider.GetRequiredService<IDocxToHtmlConverter>()
                .ConvertAsync(source), StringComparison.OrdinalIgnoreCase);

        using (var source = new MemoryStream(docx))
            Assert.Contains("# Report", await provider.GetRequiredService<IDocxToMarkdownConverter>()
                .ConvertAsync(source), StringComparison.Ordinal);

        using (var source = new MemoryStream(xlsx))
        using (var destination = new MemoryStream())
        {
            await provider.GetRequiredService<IXlsxToPdfConverter>().ConvertAsync(source, destination);
            Assert.NotEmpty(destination.ToArray());
        }

        using (var source = new MemoryStream(pptx))
        using (var destination = new MemoryStream())
        {
            await provider.GetRequiredService<IPptxToPdfConverter>().ConvertAsync(source, destination);
            Assert.NotEmpty(destination.ToArray());
        }
    }

    [Fact]
    public async Task ResolvedHtmlConverters_PageSetupOverloadsDelegateToo()
    {
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();
        const string Html = "<p>x</p>";

        Assert.NotEmpty(await provider.GetRequiredService<IHtmlToDocxConverter>()
            .ConvertAsync(Html, DocToolkit.PageSetup.Letter));

        using (var destination = new MemoryStream())
        {
            await provider.GetRequiredService<IHtmlToDocxConverter>()
                .ConvertAsync(Html, DocToolkit.PageSetup.Letter, destination);
            Assert.NotEmpty(destination.ToArray());
        }

        using (var destination = new MemoryStream())
        {
            await provider.GetRequiredService<IHtmlToPdfConverter>()
                .ConvertAsync(Html, DocToolkit.PageSetup.Letter, destination);
            Assert.NotEmpty(destination.ToArray());
        }
    }

    // A13. The services are singletons, so IOptions<T>.Value would be resolved once for the life
    // of the container and a configuration reload would silently do nothing. That is unremarkable
    // for most settings and not for this one: AllowRemoteImageDownload is the ONLY switch that
    // lets this library open a socket, so turning it off in configuration - as an incident
    // response, say - looked like it worked and did not take effect until a restart.
    //
    // Asserted by SOCKET COUNT rather than by reading the option back. Reading it back would pass
    // against a service that reads CurrentValue and then ignores it; only the probe proves the
    // reloaded value reached the conversion.
    [Fact]
    public async Task AddDocToolkit_WhenOptionsReload_TheNewValueTakesEffectWithoutARestart()
    {
        using var probe = new LoopbackProbe();

        // The configure delegate closes over `allowRemote`, so clearing the options cache re-runs
        // it and produces a genuinely different value - the same thing an appsettings reload does.
        var allowRemote = false;
        var provider = new ServiceCollection()
            .AddDocToolkit(o =>
            {
                o.AllowRemoteImageDownload = allowRemote;
                o.RemoteImage.AllowPrivateAddresses = true;   // see the note above; the probe is loopback
            })
            .BuildServiceProvider();

        var sut = provider.GetRequiredService<IHtmlToDocxConverter>();

        // Resolved and used while the option is false: nothing dials.
        await sut.ConvertAsync($"<img src=\"{probe.ImageUrl}\">");
        Assert.Equal(0, probe.Connections);

        // Now the configuration changes, on the SAME resolved singleton.
        allowRemote = true;
        provider.GetRequiredService<IOptionsMonitorCache<DocToolkitOptions>>().Clear();

        try
        {
            await sut.ConvertAsync($"<img src=\"{probe.ImageUrl}\">");
        }
        catch (DocToolkit.DocumentConversionException)
        {
            // The connection is what is under test, not a successful conversion.
        }

        Assert.True(
            await probe.WaitForConnectionAsync(TimeSpan.FromSeconds(5)),
            "The reloaded AllowRemoteImageDownload never reached the conversion. The service is "
            + "still holding the value it captured when it was constructed.");
    }

    // The other direction, and the one that actually matters operationally: turning the switch OFF
    // must take effect immediately. A guard that can only be relaxed at runtime would be worse than
    // none at all.
    [Fact]
    public async Task AddDocToolkit_WhenOptionsReloadToOff_TheFetchStops()
    {
        using var probe = new LoopbackProbe();

        var allowRemote = true;
        var provider = new ServiceCollection()
            .AddDocToolkit(o =>
            {
                o.AllowRemoteImageDownload = allowRemote;
                o.RemoteImage.AllowPrivateAddresses = true;
            })
            .BuildServiceProvider();

        var sut = provider.GetRequiredService<IHtmlToDocxConverter>();

        try { await sut.ConvertAsync($"<img src=\"{probe.ImageUrl}\">"); }
        catch (DocToolkit.DocumentConversionException) { }
        Assert.True(await probe.WaitForConnectionAsync(TimeSpan.FromSeconds(5)),
            "Precondition failed: the probe never saw the opt-in connect, so the assertion below "
            + "would pass vacuously.");

        var before = probe.Connections;
        allowRemote = false;
        provider.GetRequiredService<IOptionsMonitorCache<DocToolkitOptions>>().Clear();

        await sut.ConvertAsync($"<img src=\"{probe.ImageUrl}\">");

        Assert.Equal(before, probe.Connections);
    }


    [Fact]
    public async Task AddDocToolkit_ResolvedPdfEditor_MatchesTheStaticApi()
    {
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();
        var sut = provider.GetRequiredService<IPdfEditor>();

        var first = await DocToolkit.HtmlToPdfConverter.ConvertAsync("<h1>First</h1>");
        var second = await DocToolkit.HtmlToPdfConverter.ConvertAsync("<h1>Second</h1>");

        var merged = sut.Merge([first, second]);

        Assert.Equal(DocToolkit.PdfEditor.PageCount(merged), sut.PageCount(merged));
        Assert.Equal(2, sut.PageCount(merged));
        Assert.Equal(1, sut.PageCount(sut.ExtractPages(merged, 2, 1)));
    }

    [Fact]
    public async Task AddDocToolkit_ResolvedPdfEditor_RoundTripsMetadata()
    {
        var provider = new ServiceCollection().AddDocToolkit().BuildServiceProvider();
        var sut = provider.GetRequiredService<IPdfEditor>();

        var pdf = await DocToolkit.HtmlToPdfConverter.ConvertAsync("<h1>Report</h1>");

        var stamped = sut.WithMetadata(pdf, new DocToolkit.PdfMetadata { Title = "Quarterly" });

        Assert.Equal("Quarterly", sut.ReadMetadata(stamped).Title);

        // Absent stays absent rather than arriving as "" - the distinction the core API makes, and
        // the one an interface is most likely to lose by round-tripping through a DTO.
        Assert.Null(sut.ReadMetadata(stamped).Author);
    }
}
