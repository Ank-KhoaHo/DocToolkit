using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using Xunit.Abstractions;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace DocToolkit.Tests;

/// <summary>
/// The air-gap guarantee: <b>no code path reachable from the public API opens a socket</b>, unless
/// the caller explicitly opts in.
///
/// This matters because DocToolkit's users have no internet access at all. Their machines reach a
/// NuGet feed and nothing else, so an outbound request does not merely degrade a feature — it
/// hangs on a SYN that is never answered and then fails. "We set a safe default" is an intention;
/// these tests are the guarantee.
///
/// The method is deliberately blunt. A real TCP listener is started on 127.0.0.1 with an
/// OS-assigned port, every public API is fed content that names that listener's URL in every way a
/// document is capable of pulling a subresource, and the accepted-connection count must be exactly
/// zero. Counting sockets rather than mocking an HTTP client means the assertion holds no matter
/// which dependency, which transport or which future version does the fetching.
///
/// A guard that can only ever say "zero" proves nothing, so
/// <see cref="OptIn_DoesReachTheListener_SoTheGuardCanTellTheTwoApart"/> drives the same probe
/// through the one opt-in that is allowed to fetch and requires it to land.
/// </summary>
public class AirGapGuardTests
{
    private readonly ITestOutputHelper _output;

    public AirGapGuardTests(ITestOutputHelper output) => _output = output;

    /// <summary>
    /// TEST-NET-3 (RFC 5737). Reserved for documentation and guaranteed not to be routed, so a
    /// connection attempt to it stalls on SYN retransmits exactly the way an air-gapped machine
    /// stalls. That is the shape of the failure being guarded against: not a fast refusal, a hang.
    /// </summary>
    private const string UnroutableHost = "203.0.113.1";

    /// <summary>
    /// A closed port on loopback — the baseline for <see cref="UnroutableCeiling"/>.
    ///
    /// On the default path nothing dials at all, so a conversion pointed here costs exactly what
    /// the conversion costs. What makes it the right baseline is the worst case: <b>even if the
    /// offline guarantee broke</b>, a closed loopback port refuses rather than hanging, so this
    /// measurement can never itself contain a stall.
    ///
    /// "Refuses rather than hanging" is not the same as "instantly", which is what an earlier
    /// version of this comment claimed. Measured 2026-08-08 on Windows: a connect to
    /// <c>127.0.0.1:1</c> takes <b>2.07 s</b> — Windows retries a refused connection — against
    /// <b>21.03 s</b> for <see cref="UnroutableHost"/>. The margin is 18.9 s, which is what the
    /// 5 s ceiling is spending. POSIX platforms return <c>ECONNREFUSED</c> faster still, so
    /// Windows is the worst case and the one worth quoting.
    /// </summary>
    private const string RefusedBaselineHost = "127.0.0.1:1";

    /// <summary>
    /// Ceiling on how much SLOWER a default-path conversion is against an unroutable host than the
    /// same conversion against <see cref="RefusedBaselineHost"/>. A delta, not an absolute time,
    /// and that distinction is load-bearing rather than fussy.
    ///
    /// It was an absolute 5 s ceiling until 2026-08-08, when adding macOS to the CI matrix made it
    /// fail: `HtmlToPdf_Default_DoesNotStallOnAnUnroutableHost` took 11.1 s on a macOS runner —
    /// while its DOCX-only sibling passed, and all 37 zero-connection assertions passed. Nothing
    /// had dialled out. The whole 11 s was DOCX → PDF render on a cold runner, and the absolute
    /// ceiling could not tell that apart from a connect timeout, which is precisely the thing it
    /// exists to detect.
    ///
    /// Raising the number would have been the wrong repair: it degrades the signal without
    /// understanding the failure, and the headroom needed for a slow runner starts to overlap
    /// Windows' ~21 s connect timeout. Measuring the delta removes runner speed from the question
    /// entirely, and the margin it spends is measured rather than guessed: 21.03 s to connect to
    /// an unrouted address against 2.07 s to a refused loopback port, both on Windows — 18.9 s
    /// of headroom against render costs that differ between runners by a couple of seconds.
    /// </summary>
    private static readonly TimeSpan UnroutableCeiling = TimeSpan.FromSeconds(5);

    /// <summary>
    /// Attempts allowed before a stall is believed, of which <see cref="MeasureStallAsync"/> keeps
    /// the <b>smallest</b>. Taking the minimum is valid because the noise here is one-sided:
    /// contention on a shared runner can only make a conversion slower, never faster, while a
    /// genuine connect timeout inflates <b>every</b> attempt and so survives the minimum intact.
    ///
    /// This exists because the delta alone was not enough. Measured on a macOS runner: a baseline
    /// of 5.55 s against 13.47 s for the unroutable host — 7.92 s of apparent stall — on a test
    /// that had already passed twice on the same code, and whose absolute cost across runs ranged
    /// from under 5 s to 13.5 s. A fixed dial-out would be a constant; that spread is a 3-core
    /// runner with the rest of the suite running beside it. Note also that the SECOND call was the
    /// slower one, which is backwards from warm-up and so cannot be explained by it.
    ///
    /// Only paid when an attempt looks bad — a first attempt under the ceiling ends the loop.
    /// </summary>
    private const int StallAttempts = 3;

    /// <summary>
    /// Bound on the whole timed call, so a regression that reintroduces a network fetch shows up
    /// as a failed assertion rather than as a test suite that never finishes.
    /// </summary>
    private static readonly TimeSpan HangGuard = TimeSpan.FromSeconds(30);

    // =====================================================================================
    // The probe has to be able to say "yes" before "no" from it means anything.
    // =====================================================================================

    [Fact]
    public async Task TheProbeCountsARealConnection()
    {
        using var probe = new LoopbackProbe(_output);

        using (var client = new TcpClient())
        {
            await client.ConnectAsync(IPAddress.Loopback, probe.Port);
        }

        Assert.True(await probe.WaitForConnectionAsync(TimeSpan.FromSeconds(5)),
            "The probe failed to count a connection made directly to it, so every zero it " +
            "reports elsewhere in this class is meaningless.");
    }

    [Fact]
    public async Task OptIn_DoesReachTheListener_SoTheGuardCanTellTheTwoApart()
    {
        using var probe = new LoopbackProbe(_output);

        await RemoteDownloadGate.RunAsync(async () =>
        {
            try
            {
                // GuardedResourceLoader blocks loopback, private and link-local addresses by
                // default (RemoteImageOptions.AllowPrivateAddresses is false), which would refuse
                // this listener too. AllowPrivateAddresses = true is the escape hatch this test
                // depends on to reach it - naming that opt-in here is more honest than the old
                // form, which relied on there being no restriction at all.
                await HtmlToDocxConverter.ConvertAsync(
                    $"""<p>Report <img src="{probe.BaseUrl}/logo.bmp" alt="logo" /></p>""",
                    new RemoteImageOptions { AllowPrivateAddresses = true });
            }
            catch (DocumentConversionException)
            {
                // A guarded fetch can still fail for other reasons (a malformed response, a
                // cancelled read). Irrelevant here: what is under test is that the opt-in reaches
                // the wire at all, and the connection count below settles that either way.
            }
        });

        Assert.True(await probe.WaitForConnectionAsync(TimeSpan.FromSeconds(10)),
            "The RemoteImageOptions opt-in made no outbound connection. Either it no longer " +
            "reaches HtmlToOpenXml's image processing mode, or this probe cannot detect a fetch " +
            "- and if it cannot detect one, the zero-connection assertions below are vacuous.");
    }

    [Fact]
    public async Task OptIn_OnTheHtmlToPdfPath_AlsoReachesTheListener()
    {
        using var probe = new LoopbackProbe(_output);

        await RemoteDownloadGate.RunAsync(async () =>
        {
            try
            {
                // As above: the guard blocks loopback by default, so AllowPrivateAddresses = true
                // is the escape hatch this test names to prove HtmlToPdfConverter forwards the
                // options through to the HTML stage rather than dropping them.
                await HtmlToPdfConverter.ConvertAsync(
                    $"""<p>Report <img src="{probe.BaseUrl}/logo.bmp" alt="logo" /></p>""",
                    new RemoteImageOptions { AllowPrivateAddresses = true });
            }
            catch (DocumentConversionException)
            {
                // See above.
            }
        });

        Assert.True(await probe.WaitForConnectionAsync(TimeSpan.FromSeconds(10)),
            "HtmlToPdfConverter did not forward the RemoteImageOptions opt-in to the HTML stage.");
    }

    // =====================================================================================
    // Default path: HTML converters. These carry the whole subresource surface.
    // =====================================================================================

    [Fact]
    public async Task HtmlToDocx_ConvertAsync_ContactsNothing()
    {
        using var probe = new LoopbackProbe(_output);

        var docx = await HtmlToDocxConverter.ConvertAsync(SubresourceMarkup(probe.BaseUrl));

        Assert.NotEmpty(docx);
        await probe.AssertSilentAsync("HtmlToDocxConverter.ConvertAsync");
    }

    [Fact]
    public async Task HtmlToDocx_ExplicitlyDisallowedOverload_ContactsNothing()
    {
        using var probe = new LoopbackProbe(_output);

        await HtmlToDocxConverter.ConvertAsync(
            SubresourceMarkup(probe.BaseUrl), allowRemoteImageDownload: false);

        await probe.AssertSilentAsync("HtmlToDocxConverter.ConvertAsync(allowRemoteImageDownload: false)");
    }

    [Fact]
    public async Task HtmlToDocx_ConvertToFileAsync_ContactsNothing()
    {
        using var probe = new LoopbackProbe(_output);
        using var output = new TempFile(".docx");

        await HtmlToDocxConverter.ConvertToFileAsync(SubresourceMarkup(probe.BaseUrl), output.Path);

        Assert.True(new FileInfo(output.Path).Length > 0);
        await probe.AssertSilentAsync("HtmlToDocxConverter.ConvertToFileAsync");
    }

    [Fact]
    public async Task HtmlToPdf_ConvertAsync_ContactsNothing()
    {
        using var probe = new LoopbackProbe(_output);

        var pdf = await HtmlToPdfConverter.ConvertAsync(SubresourceMarkup(probe.BaseUrl));

        Assert.NotEmpty(pdf);
        await probe.AssertSilentAsync("HtmlToPdfConverter.ConvertAsync");
    }

    [Fact]
    public async Task HtmlToPdf_ExplicitlyDisallowedOverload_ContactsNothing()
    {
        using var probe = new LoopbackProbe(_output);

        await HtmlToPdfConverter.ConvertAsync(
            SubresourceMarkup(probe.BaseUrl), allowRemoteImageDownload: false);

        await probe.AssertSilentAsync("HtmlToPdfConverter.ConvertAsync(allowRemoteImageDownload: false)");
    }

    [Fact]
    public async Task HtmlToPdf_ConvertToFileAsync_ContactsNothing()
    {
        using var probe = new LoopbackProbe(_output);
        using var output = new TempFile(".pdf");

        await HtmlToPdfConverter.ConvertToFileAsync(SubresourceMarkup(probe.BaseUrl), output.Path);

        Assert.True(new FileInfo(output.Path).Length > 0);
        await probe.AssertSilentAsync("HtmlToPdfConverter.ConvertToFileAsync");
    }

    /// <summary>
    /// The other half of the promise: with the network shut off, the images that <i>are</i>
    /// self-contained still make it into the document.
    ///
    /// This is the canary for any tightening of the no-network default. "Nothing was fetched" is
    /// trivially satisfiable by embedding nothing at all, so the guard needs a companion that
    /// fails if the default ever degrades from "does not reach out" to "drops images".
    /// </summary>
    [Fact]
    public async Task DataUriImagesAreStillEmbedded_WithNothingFetchable()
    {
        using var probe = new LoopbackProbe(_output);
        var dataUri = "data:image/bmp;base64," + Convert.ToBase64String(ImageFixtures.Bmp());

        var docx = await HtmlToDocxConverter.ConvertAsync(
            $"""
             <p><img src="{dataUri}" alt="inline" /></p>
             <p><img src="{probe.BaseUrl}/remote.png" alt="remote" /></p>
             """);

        Assert.Equal(1, DocxFixtures.Read(docx, main => main.ImageParts.Count()));
        await probe.AssertSilentAsync("HtmlToDocxConverter.ConvertAsync (data URI fixture)");
    }

    /// <summary>
    /// Each subresource form on its own, so a failure names the vector instead of the fixture.
    /// A single kitchen-sink document tells you something leaked; this tells you what.
    /// </summary>
    [Theory]
    [InlineData("<img src>", """<p><img src="{0}/x.png" alt="logo" /></p>""")]
    [InlineData("<link rel=stylesheet>", """<link rel="stylesheet" type="text/css" href="{0}/x.css" /><p>text</p>""")]
    [InlineData("@import", """<style>@import url("{0}/y.css");</style><p>text</p>""")]
    [InlineData("<img> in a table", """<table><tr><td><img src="{0}/cell.png" alt="c" /></td></tr></table>""")]
    [InlineData("<a href>", """<p><a href="{0}/page.html">never fetched</a></p>""")]
    [InlineData("inline background-image", """<p style="background-image: url('{0}/bg.png')">shaded</p>""")]
    [InlineData("background shorthand", """<div style="background: url({0}/bg2.png) no-repeat">x</div>""")]
    [InlineData("stylesheet background-image", """<style>p { background-image: url("{0}/bg3.png"); }</style><p>x</p>""")]
    [InlineData("<base href> + relative img", """<base href="{0}/assets/" /><p><img src="rel.png" alt="r" /></p>""")]
    [InlineData("<img srcset>", """<p><img srcset="{0}/a-2x.png 2x, {0}/a-1x.png 1x" src="{0}/a.png" alt="s" /></p>""")]
    [InlineData("<object data>", """<object data="{0}/o.bin"></object>""")]
    [InlineData("<embed src>", """<embed src="{0}/e.bin" />""")]
    [InlineData("<iframe src>", """<iframe src="{0}/frame.html"></iframe>""")]
    [InlineData("<video poster>", """<video poster="{0}/poster.png"><source src="{0}/v.mp4" /></video>""")]
    [InlineData("<input type=image>", """<input type="image" src="{0}/button.png" />""")]
    [InlineData("<script src>", """<script src="{0}/app.js"></script><p>x</p>""")]
    public async Task HtmlToDocx_ContactsNothing_ForEverySubresourceForm(string vector, string template)
    {
        using var probe = new LoopbackProbe(_output);

        // Replace rather than string.Format: several of these templates are CSS, and CSS braces
        // would have to be doubled to survive a composite format string.
        await HtmlToDocxConverter.ConvertAsync(template.Replace("{0}", probe.BaseUrl));

        await probe.AssertSilentAsync($"HtmlToDocxConverter.ConvertAsync with {vector}");
    }

    // =====================================================================================
    // Default path: the OpenXml APIs. Office formats can carry external relationships - a
    // linked (not embedded) picture, a hyperlink, an external workbook reference - so each
    // fixture below is a real package with those relationships pointing at the probe.
    // =====================================================================================

    [Fact]
    public async Task DocxToPdf_Convert_ContactsNothing()
    {
        using var probe = new LoopbackProbe(_output);

        // A hand-built package with an external image relationship (Word's "link to file"
        // picture) and an external hyperlink. A renderer that resolves linked images would
        // fetch here; whether OfficeIMO renders it or refuses it, it must not dial out.
        try
        {
            DocxToPdfConverter.Convert(DocxWithExternalReferences(probe.BaseUrl));
        }
        catch (DocumentConversionException)
        {
            // Refusing to render a linked picture is a perfectly good outcome. Fetching it is not.
        }

        await probe.AssertSilentAsync("DocxToPdfConverter.Convert");
    }

    [Fact]
    public async Task DocxToPdf_Convert_ContactsNothing_ForAConvertedHtmlDocument()
    {
        using var probe = new LoopbackProbe(_output);

        // HtmlToOpenXml turns every <a href> into a real external hyperlink relationship, so this
        // exercises the same path with a package produced by the library itself rather than by
        // the test.
        var docx = await HtmlToDocxConverter.ConvertAsync(SubresourceMarkup(probe.BaseUrl));
        var pdf = DocxToPdfConverter.Convert(docx);

        Assert.NotEmpty(pdf);
        await probe.AssertSilentAsync("DocxToPdfConverter.Convert (HTML-derived package)");
    }

    [Fact]
    public async Task DocxToPdf_ConvertFile_ContactsNothing()
    {
        using var probe = new LoopbackProbe(_output);
        using var input = new TempFile(".docx");
        using var output = new TempFile(".pdf");

        await File.WriteAllBytesAsync(
            input.Path, await HtmlToDocxConverter.ConvertAsync(SubresourceMarkup(probe.BaseUrl)));
        DocxToPdfConverter.ConvertFile(input.Path, output.Path);

        Assert.True(new FileInfo(output.Path).Length > 0);
        await probe.AssertSilentAsync("DocxToPdfConverter.ConvertFile");
    }

    [Fact]
    public async Task DocxEditor_ContactsNothing()
    {
        using var probe = new LoopbackProbe(_output);

        var docx = DocxWithExternalReferences(probe.BaseUrl);
        var replaced = DocxEditor.ReplaceText(
            docx, new Dictionary<string, string> { ["{{label}}"] = probe.BaseUrl + "/replaced.png" });

        Assert.NotEmpty(DocxEditor.ExtractText(replaced));
        DocxEditor.ExtractText(replaced, includeHeadersAndFooters: true);

        await probe.AssertSilentAsync("DocxEditor.ReplaceText / ExtractText");
    }

    [Fact]
    public async Task DocxEditorFillRows_ContactsNothing()
    {
        using var probe = new LoopbackProbe(_output);

        // The template row, the substituted values and the surrounding document all name the
        // listener, so a fetch triggered by any of the three would show up.
        var docx = DocxFixtures.Build(
            DocxFixtures.P(DocxFixtures.R($"See {probe.BaseUrl}/index.html")),
            DocxFixtures.Tbl(
                DocxFixtures.Row(DocxFixtures.R("Description")),
                DocxFixtures.Row(DocxFixtures.R($"{{{{item.Desc}}}} — {probe.BaseUrl}/row.png"))));

        var filled = DocxEditor.FillRows(docx, "item", new[]
        {
            new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["Desc"] = $"{probe.BaseUrl}/value.png",
            },
        });

        Assert.NotEmpty(DocxEditor.ExtractText(filled));

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        await DocxEditor.FillRowsAsync(source, "item", new[]
        {
            new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["Desc"] = $"{probe.BaseUrl}/async.png",
            },
        }, destination);

        Assert.NotEmpty(destination.ToArray());

        await probe.AssertSilentAsync("DocxEditor.FillRows / FillRowsAsync");
    }

    [Fact]
    public async Task DocxEditorReplaceImage_ContactsNothing()
    {
        using var probe = new LoopbackProbe(_output);

        // The surrounding text names the listener, so anything that tried to resolve document
        // content while inserting the image would show up as a connection.
        var docx = DocxFixtures.Build(
            DocxFixtures.P(DocxFixtures.R($"Logo {{{{logo}}}} from {probe.BaseUrl}/logo.png")));

        var filled = DocxEditor.ReplaceImage(docx, "{{logo}}", ImageFixtures.Png());
        Assert.NotEmpty(filled);

        using var source = new MemoryStream(docx);
        using var destination = new MemoryStream();
        await DocxEditor.ReplaceImageAsync(
            source, "{{logo}}", ImageFixtures.Jpeg(), destination);
        Assert.NotEmpty(destination.ToArray());

        await probe.AssertSilentAsync("DocxEditor.ReplaceImage / ReplaceImageAsync");
    }

    // The two PDF renderers added for A6-BUILD. These reuse the SAME external-reference
    // fixtures the editors above are driven with - a workbook carrying a real external
    // hyperlink relationship and a formula referencing another workbook over http, and a deck
    // carrying external relationships - because the question is identical: given a document
    // that asks to be told what is on another machine, does anything dial out.
    //
    // MEASURED LIMIT OF THESE TWO, recorded so nobody credits them with more than they prove:
    // flipping BOTH policy flags to true and re-running leaves them passing. Nothing dials out
    // either way, because an Excel external-link relationship and a PPTX external relationship
    // are not resources this renderer resolves - it renders cells and shapes, not link targets.
    //
    // So these assert that nothing dials TODAY, and would catch a future renderer that started
    // fetching. They are NOT evidence that PdfRenderPolicy is what prevents it. The flag is
    // pinned separately and mutation-verified by
    // XlsxPptxToPdfTests.ResourcePolicy_RefusesRemoteAndLocalResources, which DOES fail when a
    // flag is flipped. Two tests, two different claims; neither substitutes for the other.
    // DOCX -> text formats. This direction reads a package and writes a string, and
    // WordToHtmlOptions carries no HttpClient at all - unlike its HtmlToWordOptions counterpart.
    // Covered anyway: this suite's value is that it does not depend on that reasoning staying
    // true across an upstream version.
    [Fact]
    public async Task DocxToTextFormatConverters_ContactNothing()
    {
        using var probe = new LoopbackProbe(_output);
        byte[] docx = DocxWithExternalReferences(probe.BaseUrl);

        Assert.NotEmpty(DocxToHtmlConverter.Convert(docx));
        Assert.NotEmpty(DocxToMarkdownConverter.Convert(docx));

        using (var source = new MemoryStream(docx))
            Assert.NotEmpty(await DocxToHtmlConverter.ConvertAsync(source));
        using (var source = new MemoryStream(docx))
            Assert.NotEmpty(await DocxToMarkdownConverter.ConvertAsync(source));

        await probe.AssertSilentAsync("DocxToHtmlConverter / DocxToMarkdownConverter");
    }

    [Fact]
    public async Task XlsxToPdfConverter_ContactsNothing()
    {
        using var probe = new LoopbackProbe(_output);

        byte[] pdf = XlsxToPdfConverter.Convert(WorkbookWithExternalReferences(probe.BaseUrl));
        Assert.True(PdfProbe.IsPdf(pdf));

        using var source = new MemoryStream(WorkbookWithExternalReferences(probe.BaseUrl));
        using var destination = new MemoryStream();
        await XlsxToPdfConverter.ConvertAsync(source, destination);
        Assert.NotEmpty(destination.ToArray());

        await probe.AssertSilentAsync("XlsxToPdfConverter.Convert / ConvertAsync");
    }

    [Fact]
    public async Task PptxToPdfConverter_ContactsNothing()
    {
        using var probe = new LoopbackProbe(_output);

        byte[] pdf = PptxToPdfConverter.Convert(PptxWithExternalReferences(probe.BaseUrl));
        Assert.True(PdfProbe.IsPdf(pdf));

        using var source = new MemoryStream(PptxWithExternalReferences(probe.BaseUrl));
        using var destination = new MemoryStream();
        await PptxToPdfConverter.ConvertAsync(source, destination);
        Assert.NotEmpty(destination.ToArray());

        await probe.AssertSilentAsync("PptxToPdfConverter.Convert / ConvertAsync");
    }

    [Fact]
    public async Task WorkbookEditor_ContactsNothing()
    {
        using var probe = new LoopbackProbe(_output);

        var created = WorkbookEditor.Create("Sales", new[]
        {
            new object?[] { "Region", "Link" },
            new object?[] { "North", $"{probe.BaseUrl}/north.xlsx" },
        });

        // ...and again on a workbook that carries a real external hyperlink relationship plus a
        // formula referencing an external workbook over http, which is how a spreadsheet asks to
        // be told what is on another machine.
        var withLinks = WorkbookWithExternalReferences(probe.BaseUrl);

        Assert.Equal("Region", WorkbookEditor.ReadCell(created, "Sales", "A1"));
        Assert.Equal($"{probe.BaseUrl}/north.xlsx", WorkbookEditor.ReadCell(created, "Sales", "B2"));

        try
        {
            WorkbookEditor.ReadCell(withLinks, "Sales", "A1");
            WorkbookEditor.ReadCell(withLinks, "Sales", "C1");   // the external-workbook formula
            var updated = WorkbookEditor.SetCell(withLinks, "Sales", "B1", 1500);
            Assert.Equal("1500", WorkbookEditor.ReadCell(updated, "Sales", "B1"));

            // Bulk reads walk every cell, so they touch the external-workbook formula and the
            // hyperlink relationship together — the two things in a workbook that ask to be told
            // what is on another machine.
            WorkbookEditor.SheetNames(withLinks);
            WorkbookEditor.ReadSheet(withLinks, "Sales");
        }
        catch (DocumentConversionException)
        {
            // Refusing to open a workbook that links out is a fine outcome. Resolving the link is
            // not, and the connection count below is what actually decides this test.
        }

        await probe.AssertSilentAsync(
            "WorkbookEditor.Create / ReadCell / SetCell / SheetNames / ReadSheet");
    }

    [Fact]
    public async Task PresentationEditor_ContactsNothing()
    {
        using var probe = new LoopbackProbe(_output);

        var pptx = PptxWithExternalReferences(probe.BaseUrl);

        Assert.Equal(1, PresentationEditor.SlideCount(pptx));
        Assert.NotEmpty(PresentationEditor.ExtractText(pptx));
        PresentationEditor.ReplaceText(
            pptx, new Dictionary<string, string> { ["{{who}}"] = probe.BaseUrl + "/who.png" });

        await probe.AssertSilentAsync("PresentationEditor.SlideCount / ExtractText / ReplaceText");
    }

    // =====================================================================================
    // No hang against an unroutable host.
    //
    // An air-gapped machine does not refuse a connection, it swallows it: the SYN goes out and
    // nothing comes back, so the caller sits on the OS connect timeout (~21 s on Windows) per
    // reference. A loopback probe cannot show that, because a refused loopback connection is
    // instant - which is exactly why "the probe saw nothing" and "the call returned promptly"
    // are two separate assertions.
    // =====================================================================================

    [Fact]
    public async Task HtmlToDocx_Default_DoesNotStallOnAnUnroutableHost()
    {
        var stall = await MeasureStallAsync(
            baseUrl => HtmlToDocxConverter.ConvertAsync(SubresourceMarkup(baseUrl)),
            "HtmlToDocxConverter.ConvertAsync");

        Assert.True(stall < UnroutableCeiling,
            $"Was {stall.TotalSeconds:0.000} s slower against {UnroutableHost} than against " +
            $"{RefusedBaselineHost}, over the {UnroutableCeiling.TotalSeconds:0.#} s ceiling - " +
            "that is the shape of a connect timeout, i.e. something tried to dial out.");
    }

    [Fact]
    public async Task HtmlToPdf_Default_DoesNotStallOnAnUnroutableHost()
    {
        var stall = await MeasureStallAsync(
            baseUrl => HtmlToPdfConverter.ConvertAsync(SubresourceMarkup(baseUrl)),
            "HtmlToPdfConverter.ConvertAsync");

        Assert.True(stall < UnroutableCeiling,
            $"Was {stall.TotalSeconds:0.000} s slower against {UnroutableHost} than against " +
            $"{RefusedBaselineHost}, over the {UnroutableCeiling.TotalSeconds:0.#} s ceiling.");
    }

    [Fact]
    public async Task DocxToPdf_Default_DoesNotStallOnAnUnroutableHost()
    {
        var stall = await MeasureStallAsync(
            baseUrl =>
            {
                var docx = DocxWithExternalReferences(baseUrl);
                try { DocxToPdfConverter.Convert(docx); }
                catch (DocumentConversionException) { /* refusing to render is fine; stalling is not */ }
                return Task.CompletedTask;
            },
            "DocxToPdfConverter.Convert");

        Assert.True(stall < UnroutableCeiling,
            $"Was {stall.TotalSeconds:0.000} s slower against {UnroutableHost} than against " +
            $"{RefusedBaselineHost}, over the {UnroutableCeiling.TotalSeconds:0.#} s ceiling.");
    }

    /// <summary>
    /// How much slower <paramref name="convert"/> is against <see cref="UnroutableHost"/> than
    /// against <see cref="RefusedBaselineHost"/> — the stall, with the conversion's own cost
    /// subtracted out.
    ///
    /// The baseline runs first deliberately, so it absorbs JIT and font warm-up. That biases the
    /// delta DOWNWARDS, which is the safe direction to be wrong in only because the failure being
    /// detected is enormous: a connect timeout is ~21 s on Windows and over a minute on Linux,
    /// against a warm-up of a second or two. A tighter ceiling than that gap would need a warm-up
    /// call before the baseline; this one does not.
    /// </summary>
    private async Task<TimeSpan> MeasureStallAsync(Func<string, Task> convert, string what)
    {
        var best = TimeSpan.MaxValue;

        for (var attempt = 1; attempt <= StallAttempts; attempt++)
        {
            var baseline = await TimeBoundedAsync(
                () => convert($"http://{RefusedBaselineHost}"),
                $"{what} [baseline]", RefusedBaselineHost);

            var elapsed = await TimeBoundedAsync(
                () => convert($"http://{UnroutableHost}"), what, UnroutableHost);

            var stall = elapsed - baseline;
            if (stall < best) best = stall;

            _output.WriteLine(
                $"{what} attempt {attempt}/{StallAttempts}: {elapsed.TotalMilliseconds:0.0} ms " +
                $"against {UnroutableHost} minus {baseline.TotalMilliseconds:0.0} ms baseline = " +
                $"{stall.TotalMilliseconds:0.0} ms of stall (best so far " +
                $"{best.TotalMilliseconds:0.0} ms, ceiling {UnroutableCeiling.TotalSeconds:0.#} s).");

            // Retry only to disprove a suspected stall. A first attempt already under the ceiling
            // has answered the question, and repeating it would triple the cost of the common case.
            if (best < UnroutableCeiling) break;
        }

        return best;
    }

    /// <summary>
    /// Runs <paramref name="action"/> with a hard wall-clock bound, so a regression that
    /// reintroduces a blocking network call fails this test instead of wedging the suite.
    /// </summary>
    private async Task<TimeSpan> TimeBoundedAsync(Func<Task> action, string what, string host)
    {
        var stopwatch = Stopwatch.StartNew();
        var work = Task.Run(action);

        if (await Task.WhenAny(work, Task.Delay(HangGuard)) != work)
        {
            // Nothing will ever await the orphan; observe its exception so it does not resurface
            // as an unobserved-task escalation later in the run.
            _ = work.ContinueWith(static t => _ = t.Exception, TaskScheduler.Default);
            Assert.Fail(
                $"{what} was still running after {HangGuard.TotalSeconds:0} s against " +
                $"{host}. It is waiting on a TCP connect that will never complete.");
        }

        await work;
        stopwatch.Stop();

        _output.WriteLine(
            $"{what} against http://{host} completed in " +
            $"{stopwatch.Elapsed.TotalMilliseconds:0.0} ms.");

        return stopwatch.Elapsed;
    }

    // =====================================================================================
    // Fixtures
    // =====================================================================================

    /// <summary>
    /// One document naming <paramref name="baseUrl"/> in every way HTML can ask for a
    /// subresource: stylesheets by link and by <c>@import</c>, backgrounds inline and from a
    /// stylesheet, images plain, in a table, via <c>srcset</c>, via <c>&lt;base&gt;</c>-relative
    /// paths, plus the embedding elements and a plain hyperlink, which must never be fetched
    /// because a hyperlink is a destination, not a dependency.
    /// </summary>
    private static string SubresourceMarkup(string baseUrl) => $$"""
        <!DOCTYPE html>
        <html>
        <head>
          <base href="{{baseUrl}}/assets/" />
          <link rel="stylesheet" type="text/css" href="{{baseUrl}}/x.css" />
          <style>
            @import url("{{baseUrl}}/y.css");
            body { background-image: url("{{baseUrl}}/bg-sheet.png"); }
            p.note { background: url({{baseUrl}}/bg-note.png) no-repeat; }
          </style>
          <script src="{{baseUrl}}/app.js"></script>
        </head>
        <body>
          <h1 style="background-image: url('{{baseUrl}}/bg-inline.png')">Air-gap check</h1>
          <p class="note">Styled from a stylesheet rule.</p>
          <p><img src="{{baseUrl}}/x.png" alt="absolute" /></p>
          <p><img src="relative.png" alt="relative to base" /></p>
          <p><img srcset="{{baseUrl}}/x-2x.png 2x, {{baseUrl}}/x-1x.png 1x"
                  src="{{baseUrl}}/x-fallback.png" alt="srcset" /></p>
          <p><a href="{{baseUrl}}/page.html">a hyperlink is a destination, not a dependency</a></p>
          <table border="1">
            <tr><th>Region</th><th>Logo</th></tr>
            <tr><td>North</td><td><img src="{{baseUrl}}/cell.png" alt="in a table" /></td></tr>
          </table>
          <div style="background: url({{baseUrl}}/bg-shorthand.png) no-repeat">shorthand</div>
          <object data="{{baseUrl}}/o.bin"></object>
          <embed src="{{baseUrl}}/e.bin" />
          <iframe src="{{baseUrl}}/frame.html"></iframe>
          <video poster="{{baseUrl}}/poster.png"><source src="{{baseUrl}}/v.mp4" /></video>
          <input type="image" src="{{baseUrl}}/button.png" />
        </body>
        </html>
        """;

    /// <summary>
    /// A .docx carrying an external hyperlink relationship and an externally *linked* picture -
    /// the <c>a:blip/@r:link</c> form Word writes for "Insert &gt; Picture &gt; Link to File",
    /// where the image bytes live at a URL instead of inside the package.
    /// </summary>
    private static byte[] DocxWithExternalReferences(string baseUrl)
    {
        const string ImageRelationshipType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var body = new Body();
            main.Document = new Document(body);

            var hyperlink = main.AddHyperlinkRelationship(new Uri($"{baseUrl}/page.html"), true);
            body.AppendChild(new Paragraph(
                new Hyperlink(new Run(new Text("{{label}} - linked document"))) { Id = hyperlink.Id }));

            var picture = main.AddExternalRelationship(
                ImageRelationshipType, new Uri($"{baseUrl}/linked.png"));
            body.AppendChild(new Paragraph(new Run(LinkedPicture(picture.Id!))));

            main.Document.Save();
        }

        return ms.ToArray();
    }

    private static Drawing LinkedPicture(string relationshipId) => new(
        new DW.Inline(
            new DW.Extent { Cx = 990000L, Cy = 792000L },
            new DW.DocProperties { Id = 1U, Name = "Linked picture" },
            new A.Graphic(
                new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = 0U, Name = "linked.png" },
                            new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(
                            new A.Blip { Link = relationshipId },
                            new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0L, Y = 0L },
                                new A.Extents { Cx = 990000L, Cy = 792000L }),
                            new A.PresetGeometry(new A.AdjustValueList())
                            {
                                Preset = A.ShapeTypeValues.Rectangle,
                            })))
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture",
                })));

    /// <summary>
    /// An .xlsx with a cell hyperlinked to <paramref name="baseUrl"/> <i>and</i> a genuine external
    /// workbook link to it - the two ways a spreadsheet asks to be told what is on another machine.
    ///
    /// ClosedXML has no API for external workbook links (and its formula parser rejects the
    /// <c>'http://host/[book.xlsx]Sheet1'!A1</c> shorthand outright), so the link is grafted on
    /// with the raw OpenXml SDK the way Excel actually stores it: an <c>externalLink</c> part whose
    /// <c>externalLinkPath</c> relationship is an http URL, listed in <c>workbook.xml</c> and used
    /// by a cell formula of the form <c>[1]Sheet1!A1</c>.
    /// </summary>
    private static byte[] WorkbookWithExternalReferences(string baseUrl)
    {
        const string ExternalLinkPathRelationshipType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";

        using var ms = new MemoryStream();

        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Sales");
            sheet.Cell("A1").Value = "Linked";
            sheet.Cell("A1").SetHyperlink(new XLHyperlink(new Uri($"{baseUrl}/report.xlsx")));
            sheet.Cell("B1").Value = 1200;
            workbook.SaveAs(ms);
        }

        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, true))
        {
            var workbookPart = doc.WorkbookPart!;

            var externalPart = workbookPart.AddNewPart<ExternalWorkbookPart>();
            var linkPath = externalPart.AddExternalRelationship(
                ExternalLinkPathRelationshipType, new Uri($"{baseUrl}/book.xlsx"));
            externalPart.ExternalLink = new X.ExternalLink(
                new X.ExternalBook(new X.SheetNames(new X.SheetName { Val = "Sheet1" }))
                {
                    Id = linkPath.Id,
                });
            externalPart.ExternalLink.Save();

            // CT_Workbook orders externalReferences straight after sheets.
            var workbook = workbookPart.Workbook!;
            workbook.InsertAfter(
                new X.ExternalReferences(new X.ExternalReference
                {
                    Id = workbookPart.GetIdOfPart(externalPart),
                }),
                workbook.Sheets!);

            var worksheet = workbookPart.WorksheetParts.First().Worksheet!;
            worksheet.GetFirstChild<X.SheetData>()!
                     .Elements<X.Row>()
                     .First()
                     .AppendChild(new X.Cell
                     {
                         CellReference = "C1",
                         CellFormula = new X.CellFormula("[1]Sheet1!A1"),
                         CellValue = new X.CellValue("42"),
                     });

            worksheet.Save();
            workbook.Save();
        }

        return ms.ToArray();
    }

    /// <summary>
    /// The committed sample deck plus an externally linked picture relationship and an external
    /// hyperlink relationship on its slide part.
    /// </summary>
    private static byte[] PptxWithExternalReferences(string baseUrl)
    {
        const string ImageRelationshipType =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

        var sample = PptxFixtures.Sample();
        using var ms = new MemoryStream();
        ms.Write(sample, 0, sample.Length);
        ms.Position = 0;

        using (var doc = PresentationDocument.Open(ms, true))
        {
            var slidePart = doc.PresentationPart!.SlideParts.First();
            slidePart.AddExternalRelationship(
                ImageRelationshipType, new Uri($"{baseUrl}/slide-picture.png"));
            slidePart.AddHyperlinkRelationship(new Uri($"{baseUrl}/slide-link.html"), true);

            slidePart.Slide!.Descendants<A.Text>().First().Text = $"{{{{who}}}} {baseUrl}/inline.png";
            slidePart.Slide.Save();
        }

        return ms.ToArray();
    }

    /// <summary>A temp file that deletes itself, so a failing assertion does not litter %TEMP%.</summary>
    private sealed class TempFile : IDisposable
    {
        public TempFile(string extension) =>
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(), $"doctoolkit-airgap-{Guid.NewGuid():N}{extension}");

        public string Path { get; }

        public void Dispose()
        {
            try { File.Delete(Path); } catch (IOException) { /* best effort */ }
        }
    }

    // LoopbackProbe lives in its own file now: GuardedResourceLoaderTests reuses it (with a custom
    // responder) to exercise FetchAsync's byte cap, timeout and cancellation behaviour, and a
    // second copy of a raw-socket HTTP responder is not something to maintain twice.
}

/// <summary>
/// Serialises the handful of tests that opt in to remote image downloads.
///
/// This is no longer about HtmlToOpenXml 3.5.0's own process-wide, header-mutating static
/// HttpClient: both opt-in paths now hand <c>HtmlConverter</c> an explicit
/// <see cref="GuardedResourceLoader"/> (or <c>OfflineResourceLoader</c> on the default path), so
/// that DefaultWebRequest - and its shared HttpClient - is never constructed at all; see the
/// comment in <see cref="HtmlToDocxConverter.BuildPackageAsync"/>. GuardedResourceLoader has its
/// own static HttpClient, but a benign one: its configuration is fixed at construction and every
/// request builds its own <c>HttpRequestMessage</c>, so nothing about it is mutated per call.
///
/// The gate remains for a more conservative reason: nothing here has proven that
/// <c>HtmlConverter.ParseBody</c> itself - the parsing and document-building logic, independent of
/// the network layer - is safe to run on multiple threads at once. A plain semaphore is the right
/// tool rather than an xUnit collection: a collection would also reorder the suite, and at least
/// one neighbouring test measures wall-clock time and is sensitive to whether the assembly is warm
/// by the time it runs.
///
/// Only the opt-in path needs this; the no-network default never constructs an HtmlConverter
/// pointed at anything but the refusing OfflineResourceLoader.
/// </summary>
internal static class RemoteDownloadGate
{
    private static readonly SemaphoreSlim Gate = new(1, 1);

    public static async Task RunAsync(Func<Task> optInConversion)
    {
        await Gate.WaitAsync();
        try { await optInConversion(); }
        finally { Gate.Release(); }
    }
}
