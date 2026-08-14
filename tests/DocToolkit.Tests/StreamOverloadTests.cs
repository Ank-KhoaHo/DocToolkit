using System.Reflection;
using System.Text;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// The <see cref="Stream"/> overloads, tested as one surface rather than one class at a time.
///
/// Every one of them has the same shape — inputs first (a <c>Stream source</c> wherever the
/// <c>byte[]</c> overload took bytes), then <c>Stream destination</c>, then
/// <c>CancellationToken ct = default</c> — so the properties that matter are properties of the
/// shape, not of any one converter: the caller's streams are read and written but never closed,
/// never sought, never required to be seekable; a stream handed in the wrong way round fails with
/// a sentence rather than a <c>NotSupportedException</c> from three libraries down; and the token
/// is honoured while the source is being consumed, not merely glanced at on the way in.
///
/// Testing them one class at a time would let one of the six quietly drift. The theories below
/// enumerate the surface by name, so <b>adding an overload without adding it to these lists is the
/// only way to escape them</b> — and the round-trip facts underneath pin each one to the
/// <c>byte[]</c> overload it has to agree with.
/// </summary>
public class StreamOverloadTests
{
    private const string Html = """
        <h1>Quarterly Report</h1>
        <p>Revenue was <strong>up 12%</strong> and costs were <em>flat</em>.</p>
        <table border="1"><tr><th>Region</th><th>Total</th></tr>
        <tr><td>North</td><td>1200</td></tr></table>
        """;

    private static readonly byte[] Docx = DocxFixtures.Build(
        "Header for {{customer}}",
        "Footer text",
        DocxFixtures.P(DocxFixtures.R("Dear {{customer}}, your invoice is ready.")));

    /// <summary>A .docx whose table holds a repeating-row template, for FillRowsAsync.</summary>
    private static readonly byte[] TableDocx = DocxFixtures.Build(DocxFixtures.Tbl(
        DocxFixtures.Row(DocxFixtures.R("Description")),
        DocxFixtures.Row(DocxFixtures.R("{{item.Desc}}"))));

    /// <summary>A .docx holding an image placeholder, for ReplaceImageAsync.</summary>
    private static readonly byte[] ImageDocx = DocxFixtures.Build(
        DocxFixtures.P(DocxFixtures.R("Logo: {{logo}} end")));

    /// <summary>
    /// A .pptx holding a shape whose text is nothing but the placeholder, for
    /// PresentationEditor.ReplaceImageAsync — unlike DocxEditor's, it swaps the whole shape, so it
    /// cannot share <see cref="Pptx"/>'s "Hello {{who}}" text.
    /// </summary>
    private static readonly byte[] ImagePptx = PptxFixtures.DeckWithPlaceholderBox("{{chart}}");

    /// <summary>Blocks for DocxEditor.CreateAsync, which takes no source.</summary>
    private static readonly DocxBlock[] Blocks =
    {
        DocxBlock.Heading("Quarterly Report", 1),
        DocxBlock.Paragraph("Revenue was up 12%."),
    };

    /// <summary>Slides for PresentationEditor.CreateAsync, which takes no source.</summary>
    private static readonly PptxSlide[] Slides =
    {
        PptxSlide.Titled("Quarterly Report", "Revenue was up 12%."),
    };

    private static readonly IReadOnlyList<IReadOnlyDictionary<string, string>> FillRowsRecords =
        new[]
        {
            new Dictionary<string, string>(StringComparer.Ordinal) { ["Desc"] = "Widget" },
            new Dictionary<string, string>(StringComparer.Ordinal) { ["Desc"] = "Gadget" },
        };

    private static readonly byte[] Xlsx = WorkbookEditor.Create("Sales", new[]
    {
        new object?[] { "Region", "Total" },
        new object?[] { "North", 1200 },
    });

    private static readonly byte[] Pptx = PptxFixtures.Sample();

    /// <summary>Markdown for MarkdownToDocxConverter.ConvertAsync, which takes no source.</summary>
    private const string Md = """
        # Quarterly Report

        Revenue was **up 12%**.
        """;

    /// <summary>A .pdf, for PdfEditor.ExtractTextAsync. Built from <see cref="Docx"/> rather than
    /// through HtmlToPdfConverter because that path is async and these fields are initialised
    /// eagerly.</summary>
    private static readonly byte[] Pdf = DocxToPdfConverter.Convert(Docx);

    /// <summary>
    /// A two-page .pdf, for RemovePagesAsync — which refuses to remove every page, so it is the one
    /// PdfEditor overload that cannot run against the single-page <see cref="Pdf"/>.
    ///
    /// Declared after <see cref="Pdf"/> because static field initialisers run in declaration order.
    /// </summary>
    private static readonly byte[] TwoPagePdf = PdfEditor.Merge(new[] { Pdf, Pdf });

    /// <summary>Keys for all three formats, so one dictionary drives every ReplaceText overload.</summary>
    private static readonly Dictionary<string, string> Replacements = new()
    {
        ["{{customer}}"] = "Contoso Ltd",
        ["{{who}}"] = "World",
    };

    private static readonly object?[][] Rows =
    {
        new object?[] { "Region", "Total" },
        new object?[] { "North", 1200 },
    };

    /// <summary>Sheets for WorkbookEditor.CreateAsync(sheets), which takes no source.</summary>
    private static readonly XlsxSheet[] Sheets =
    {
        XlsxSheet.Named("Sales", Rows),
    };

    // =====================================================================================
    // The surface, by name. Every Stream overload appears in at least one of these lists.
    // =====================================================================================

    /// <summary>Overloads that take a <c>Stream destination</c>.</summary>
    private static readonly string[] DestinationWriterNames =
    {
        "HtmlToDocxConverter.ConvertAsync",
        "HtmlToDocxConverter.ConvertAsync(allowRemoteImageDownload)",
        "HtmlToDocxConverter.ConvertAsync(RemoteImageOptions)",
        "HtmlToDocxConverter.ConvertAsync(PageSetup)",
        "HtmlToPdfConverter.ConvertAsync",
        "HtmlToPdfConverter.ConvertAsync(allowRemoteImageDownload)",
        "HtmlToPdfConverter.ConvertAsync(RemoteImageOptions)",
        "HtmlToPdfConverter.ConvertAsync(PageSetup)",
        "DocxToPdfConverter.ConvertAsync",
        "XlsxToPdfConverter.ConvertAsync",
        "PptxToPdfConverter.ConvertAsync",
        "DocxEditor.ReplaceTextAsync",
        "DocxEditor.FillRowsAsync",
        "DocxEditor.ReplaceImageAsync",
        "DocxEditor.CreateAsync",
        "DocxEditor.CreateAsync(PageSetup)",
        "WorkbookEditor.CreateAsync",
        "WorkbookEditor.CreateAsync(sheets)",
        "WorkbookEditor.SetCellAsync",
        "WorkbookEditor.AppendRowsAsync",
        "WorkbookEditor.FormatAsync",
        "PresentationEditor.ReplaceTextAsync",
        "PresentationEditor.ReplaceImageAsync",
        "PresentationEditor.CreateAsync",
        "MarkdownToDocxConverter.ConvertAsync",
        "MarkdownToPdfConverter.ConvertAsync",
        "PdfEditor.MergeAsync",
        "PdfEditor.ExtractPagesAsync",
        "PdfEditor.RemovePagesAsync",
        "PdfEditor.RotatePagesAsync",
        "PdfEditor.ReorderPagesAsync",
        "PdfEditor.InsertPagesAsync",
    };

    /// <summary>Overloads that take a <c>Stream source</c>.</summary>
    private static readonly string[] SourceReaderNames =
    {
        "DocxToPdfConverter.ConvertAsync",
        "DocxEditor.ReplaceTextAsync",
        "DocxEditor.FillRowsAsync",
        "DocxEditor.ReplaceImageAsync",
        "DocxEditor.ExtractTextAsync",
        "PdfEditor.ExtractTextAsync",
        "DocxToHtmlConverter.ConvertAsync",
        "DocxToMarkdownConverter.ConvertAsync",
        "DocxToHtmlConverter.ConvertWithReportAsync",
        "DocxToMarkdownConverter.ConvertWithReportAsync",
        "XlsxToCsvConverter.ConvertAsync",
        "XlsxToHtmlConverter.ConvertAsync",
        "DocxEditor.ExtractTextAsync(includeHeadersAndFooters)",
        "DocxEditor.TableCountAsync",
        "DocxEditor.ReadTableAsync",
        "WorkbookEditor.ReadCellAsync",
        "WorkbookEditor.SheetNamesAsync",
        "WorkbookEditor.ReadSheetAsync",
        "WorkbookEditor.SetCellAsync",
        "WorkbookEditor.AppendRowsAsync",
        "WorkbookEditor.FormatAsync",
        "PresentationEditor.SlideCountAsync",
        "PresentationEditor.ExtractTextAsync",
        "PresentationEditor.ReplaceTextAsync",
        "PresentationEditor.ReplaceImageAsync",
        "PdfEditor.PageCountAsync",
        "PdfEditor.MergeAsync",
        "PdfEditor.ExtractPagesAsync",
        "PdfEditor.RemovePagesAsync",
        "PdfEditor.RotatePagesAsync",
        "PdfEditor.ReorderPagesAsync",
        "PdfEditor.InsertPagesAsync",
    };

    /// <summary>
    /// The parameter a source reader names when it refuses the stream it was handed.
    ///
    /// "source" for all but two, and the two exceptions are why this is a lookup rather than a
    /// literal in the assertion: <c>MergeAsync</c> takes an <c>IEnumerable&lt;Stream&gt; sources</c>
    /// and <c>InsertPagesAsync</c> reads a <c>target</c> before its <c>source</c>. Asserting the
    /// literal "source" would have forced both out of the theory — and being outside the theory is
    /// exactly how all seven PdfEditor overloads escaped this suite in the first place.
    /// </summary>
    private static string SourceParamName(string api) => api switch
    {
        "PdfEditor.MergeAsync" => "sources",
        "PdfEditor.InsertPagesAsync" => "target",
        _ => "source",
    };

    /// <summary>
    /// Destination writers whose output is assembled and then copied out with
    /// <c>CopyToAsync</c>. Excludes the two PDF paths, which hand the caller's destination to
    /// OfficeIMO's own writer instead of buffering the PDF — see
    /// <see cref="DocxToPdf_StreamsThePdfToTheDestinationInPieces_RatherThanBufferingItWhole"/>.
    /// </summary>
    private static readonly string[] BufferedDestinationWriterNames =
    {
        "HtmlToDocxConverter.ConvertAsync",
        "HtmlToDocxConverter.ConvertAsync(allowRemoteImageDownload)",
        "HtmlToDocxConverter.ConvertAsync(RemoteImageOptions)",
        "HtmlToDocxConverter.ConvertAsync(PageSetup)",
        "DocxEditor.ReplaceTextAsync",
        "DocxEditor.FillRowsAsync",
        "DocxEditor.ReplaceImageAsync",
        "DocxEditor.CreateAsync",
        "DocxEditor.CreateAsync(PageSetup)",
        "WorkbookEditor.CreateAsync",
        "WorkbookEditor.CreateAsync(sheets)",
        "WorkbookEditor.SetCellAsync",
        "WorkbookEditor.AppendRowsAsync",
        "WorkbookEditor.FormatAsync",
        "PresentationEditor.ReplaceTextAsync",
        "PresentationEditor.ReplaceImageAsync",
        "PresentationEditor.CreateAsync",
        "MarkdownToDocxConverter.ConvertAsync",
        "PdfEditor.MergeAsync",
        "PdfEditor.ExtractPagesAsync",
        "PdfEditor.RemovePagesAsync",
        "PdfEditor.RotatePagesAsync",
        "PdfEditor.ReorderPagesAsync",
        "PdfEditor.InsertPagesAsync",
    };

    public static TheoryData<string> DestinationWriters => Cases(DestinationWriterNames);

    public static TheoryData<string> SourceReaders => Cases(SourceReaderNames);

    public static TheoryData<string> BufferedDestinationWriters => Cases(BufferedDestinationWriterNames);

    /// <summary>Every Stream overload, writers and readers alike, each exactly once.</summary>
    public static TheoryData<string> AllOverloads
        => Cases(DestinationWriterNames.Union(SourceReaderNames, StringComparer.Ordinal).ToArray());

    /// <summary>
    /// The lists above are an inventory, and an inventory drifts. This DERIVES the surface from the
    /// shipped assembly and fails naming anything missing.
    ///
    /// <b>It has now caught the same class of gap twice.</b> B17: every one of <c>PdfEditor</c>'s
    /// eight Stream overloads was absent, and registering them failed 17 cases, all real defects in
    /// shipped code. Then, one day after that fix,
    /// <c>DocxToHtmlConverter.ConvertWithReportAsync</c> and its Markdown twin shipped in 0.25.0
    /// without being registered either — harmless, as it turned out, but invisible.
    ///
    /// The class doc comment above says "adding an overload without adding it to these lists is the
    /// only way to escape them". That was true, and it is precisely why a hand-maintained list is
    /// the wrong shape for it. Same principle as <c>gen-third-party-notices.py</c> reading the
    /// lockfile and <c>automerge-eligible.py</c> reading the workflows: derive, do not remember.
    ///
    /// <b>Known limitation, stated so it is a choice.</b> This matches on <c>Class.Method</c>, not
    /// on signature, because that is how the lists are keyed. A NEW OVERLOAD of an already-listed
    /// name is therefore still invisible here — the same limitation `CLAUDE.md` records for the 1:1
    /// interface check. What it does catch is a whole method, or a whole class, going unregistered,
    /// which is what happened both times.
    /// </summary>
    [Fact]
    public void EveryPublicStreamOverloadIsRegisteredInTheListsAbove()
    {
        var listed = DestinationWriterNames
            .Concat(SourceReaderNames)
            .Concat(BufferedDestinationWriterNames)
            .Select(n => n.Split('(')[0])
            .ToHashSet(StringComparer.Ordinal);

        var shipped =
            from type in typeof(DocxEditor).Assembly.GetExportedTypes()
            where type is { IsAbstract: true, IsSealed: true }        // C# static class
            from method in type.GetMethods(BindingFlags.Public | BindingFlags.Static
                                           | BindingFlags.DeclaredOnly)
            where method.GetParameters().Any(p => typeof(Stream).IsAssignableFrom(p.ParameterType))
            select $"{type.Name}.{method.Name}";

        var missing = shipped.Distinct().Where(m => !listed.Contains(m)).OrderBy(m => m).ToList();

        Assert.True(missing.Count == 0,
            "These public Stream overloads are not in this file's name lists, so every theory here "
            + "skips them silently:\n  " + string.Join("\n  ", missing));
    }

    private static TheoryData<string> Cases(IEnumerable<string> names)
    {
        var data = new TheoryData<string>();
        foreach (var name in names) data.Add(name);
        return data;
    }

    // =====================================================================================
    // Destinations
    // =====================================================================================

    /// <summary>
    /// The whole point of the overload: a destination that is write-only, forward-only and not
    /// seekable — an HTTP response body — receives a complete document.
    ///
    /// <see cref="ForwardOnlySink"/> throws on <c>Read</c>, <c>Seek</c>, <c>Length</c> and
    /// <c>Position</c>, so an implementation that rewinds the destination to patch a header, or
    /// reads back what it wrote, fails here rather than in production against a socket. The
    /// stream is also left open: DocToolkit did not open it and must not close it.
    /// </summary>
    [Theory]
    [MemberData(nameof(DestinationWriters))]
    public async Task EveryDestinationWriter_WritesAWholeDocumentToAForwardOnlySink_AndLeavesItOpen(string api)
    {
        var sink = new ForwardOnlySink();
        var destination = new TrackingStream(sink);
        using var source = NewSource(api);

        await InvokeAsync(api, source, destination);

        var written = sink.ToArray();
        Assert.True(written.Length > 0, $"{api} wrote nothing to the destination.");
        AssertLooksLikeADocument(api, written);
        Assert.Equal(0, destination.Seeks);
        Assert.False(destination.IsDisposed, $"{api} disposed a destination stream it does not own.");
        Assert.False(sink.IsDisposed, $"{api} disposed a destination stream it does not own.");
    }

    /// <summary>
    /// Bytes reach the caller's destination through <c>WriteAsync</c>, not <c>Write</c>: these
    /// overloads exist so a caller can push a document at a socket without pinning a thread while
    /// it drains.
    /// </summary>
    [Theory]
    [MemberData(nameof(BufferedDestinationWriters))]
    public async Task EveryBufferedWriter_WritesToTheDestinationAsynchronously(string api)
    {
        var destination = new TrackingStream(new ForwardOnlySink());
        using var source = NewSource(api);

        await InvokeAsync(api, source, destination);

        Assert.True(destination.AsyncWrites > 0, $"{api} never called WriteAsync on the destination.");
        Assert.Equal(0, destination.SyncWrites);
    }

    /// <summary>A destination that cannot be written is named as such, not left to fail later.</summary>
    [Theory]
    [MemberData(nameof(DestinationWriters))]
    public async Task EveryDestinationWriter_RejectsADestinationItCannotWriteTo(string api)
    {
        using var forNull = NewSource(api);
        await Assert.ThrowsAsync<ArgumentNullException>(() => InvokeAsync(api, forNull, null));

        using var forUnwritable = NewSource(api);
        var ex = await Assert.ThrowsAsync<ArgumentException>(
            () => InvokeAsync(api, forUnwritable, new NonWritableStream()));

        Assert.Equal("destination", ex.ParamName);
        Assert.Contains("writable", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    // =====================================================================================
    // Sources
    // =====================================================================================

    /// <summary>
    /// A source that is forward-only and not seekable — an HTTP request body — is consumed, and
    /// consumed with <c>ReadAsync</c>, never <c>Read</c>, and never rewound. And it is left open.
    /// </summary>
    [Theory]
    [MemberData(nameof(SourceReaders))]
    public async Task EverySourceReader_ConsumesAForwardOnlySourceAsynchronously_AndLeavesItOpen(string api)
    {
        var forwardOnly = new ForwardOnlySource(SourceBytesFor(api));
        var source = new TrackingStream(forwardOnly);
        using var destination = new MemoryStream();

        await InvokeAsync(api, source, destination);

        Assert.True(source.AsyncReads > 0, $"{api} never called ReadAsync on the source.");
        Assert.Equal(0, source.SyncReads);
        Assert.Equal(0, source.Seeks);
        Assert.False(source.IsDisposed, $"{api} disposed a source stream it does not own.");
        Assert.False(forwardOnly.IsDisposed, $"{api} disposed a source stream it does not own.");
    }

    /// <summary>
    /// Null, unreadable and empty sources each produce their own sentence. Empty matters because
    /// a zero-byte "document" is the shape of a truncated upload, and the <c>byte[]</c> overloads
    /// already reject it by name.
    /// </summary>
    [Theory]
    [MemberData(nameof(SourceReaders))]
    public async Task EverySourceReader_RejectsASourceItCannotRead(string api)
    {
        using var nullCaseDestination = new MemoryStream();
        using var unreadableCaseDestination = new MemoryStream();
        using var emptyCaseSource = new MemoryStream();
        using var emptyCaseDestination = new MemoryStream();

        await Assert.ThrowsAsync<ArgumentNullException>(
            () => InvokeAsync(api, null, nullCaseDestination));

        var expectedParam = SourceParamName(api);

        var unreadable = await Assert.ThrowsAsync<ArgumentException>(
            () => InvokeAsync(api, new NonReadableStream(), unreadableCaseDestination));
        Assert.Equal(expectedParam, unreadable.ParamName);
        Assert.Contains("readable", unreadable.Message, StringComparison.OrdinalIgnoreCase);

        var empty = await Assert.ThrowsAsync<ArgumentException>(
            () => InvokeAsync(api, emptyCaseSource, emptyCaseDestination));
        Assert.Equal(expectedParam, empty.ParamName);
        Assert.Contains("empty", empty.Message, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// The token is observed <i>while the source is being read</i>, not only at the guard on the
    /// way in. <see cref="CancelsOnFirstReadSource"/> cancels from inside the first read, by which
    /// point the entry check has already passed.
    /// </summary>
    [Theory]
    [MemberData(nameof(SourceReaders))]
    public async Task EverySourceReader_HonoursATokenCancelledWhileTheSourceIsBeingRead(string api)
    {
        using var cts = new CancellationTokenSource();
        var source = new CancelsOnFirstReadSource(SourceBytesFor(api), cts);
        using var destination = new MemoryStream();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => InvokeAsync(api, source, destination, cts.Token));
    }

    /// <summary>Nothing starts on a token that is already cancelled.</summary>
    [Theory]
    [MemberData(nameof(AllOverloads))]
    public async Task EveryOverload_ThrowsForAnAlreadyCancelledToken(string api)
    {
        using var cts = new CancellationTokenSource();
        await cts.CancelAsync();
        using var source = NewSource(api);
        using var destination = new MemoryStream();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => InvokeAsync(api, source, destination, cts.Token));
    }

    // =====================================================================================
    // Round-trip equivalence: the Stream overload must agree with the byte[] overload it
    // shadows, or the two APIs are two different products.
    // =====================================================================================

    /// <summary>
    /// Compared on <c>word/document.xml</c> rather than on the raw package bytes: an OOXML package
    /// is a ZIP, and a ZIP stamps every entry with the time it was written, so two byte-identical
    /// documents produced a second apart are not byte-identical files. The part XML is the thing
    /// that actually has to agree.
    /// </summary>
    [Fact]
    public async Task HtmlToDocx_StreamOverload_ProducesTheSamePackageAsTheByteArrayOverload()
    {
        var expected = await HtmlToDocxConverter.ConvertAsync(Html);

        using var destination = new MemoryStream();
        await HtmlToDocxConverter.ConvertAsync(Html, destination);

        Assert.Equal(DocumentXml(expected), DocumentXml(destination.ToArray()));
    }

    [Fact]
    public async Task HtmlToDocx_StreamOverloadWithTheRemoteFlag_ProducesTheSamePackage()
    {
        var expected = await HtmlToDocxConverter.ConvertAsync(Html, allowRemoteImageDownload: false);

        using var destination = new MemoryStream();
        await HtmlToDocxConverter.ConvertAsync(Html, allowRemoteImageDownload: false, destination);

        Assert.Equal(DocumentXml(expected), DocumentXml(destination.ToArray()));
    }

    [Fact]
    public async Task DocxToPdf_StreamOverload_ProducesTheSamePdfAsTheByteArrayOverload()
    {
        var expected = DocxToPdfConverter.Convert(Docx);

        using var source = StreamDoubles.Seekable(Docx);
        using var destination = new MemoryStream();
        await DocxToPdfConverter.ConvertAsync(source, destination);

        var streamed = destination.ToArray();
        Assert.True(PdfProbe.IsPdf(streamed));
        Assert.Equal(PdfProbe.ExtractText(expected), PdfProbe.ExtractText(streamed));
        Assert.Equal(PdfProbe.PageCount(expected), PdfProbe.PageCount(streamed));
    }

    [Fact]
    public async Task HtmlToPdf_StreamOverload_ProducesTheSamePdfAsTheByteArrayOverload()
    {
        var expected = await HtmlToPdfConverter.ConvertAsync(Html);

        using var destination = new MemoryStream();
        await HtmlToPdfConverter.ConvertAsync(Html, destination);

        var streamed = destination.ToArray();
        Assert.True(PdfProbe.IsPdf(streamed));
        Assert.Equal(PdfProbe.ExtractText(expected), PdfProbe.ExtractText(streamed));
        Assert.Contains("Quarterly Report", PdfProbe.ExtractText(streamed));
    }

    [Fact]
    public async Task DocxEditor_ReplaceTextAsync_MatchesTheByteArrayOverload()
    {
        var expected = DocxEditor.ReplaceText(Docx, Replacements);

        using var source = StreamDoubles.Seekable(Docx);
        using var destination = new MemoryStream();
        await DocxEditor.ReplaceTextAsync(source, Replacements, destination);

        var streamed = destination.ToArray();
        Assert.Equal(DocumentXml(expected), DocumentXml(streamed));
        Assert.Equal(
            DocxEditor.ExtractText(expected, includeHeadersAndFooters: true),
            DocxEditor.ExtractText(streamed, includeHeadersAndFooters: true));
        Assert.Contains("Contoso Ltd", DocxEditor.ExtractText(streamed));
        Assert.DoesNotContain("{{customer}}", DocxEditor.ExtractText(streamed, includeHeadersAndFooters: true));
    }

    [Fact]
    public async Task DocxEditor_ExtractTextAsync_MatchesTheByteArrayOverloads()
    {
        using var bodyOnly = StreamDoubles.Seekable(Docx);
        Assert.Equal(DocxEditor.ExtractText(Docx), await DocxEditor.ExtractTextAsync(bodyOnly));

        using var withHeaders = StreamDoubles.Seekable(Docx);
        Assert.Equal(
            DocxEditor.ExtractText(Docx, includeHeadersAndFooters: true),
            await DocxEditor.ExtractTextAsync(withHeaders, includeHeadersAndFooters: true));
    }

    [Fact]
    public async Task WorkbookEditor_CreateAsync_MatchesTheByteArrayOverload()
    {
        var expected = WorkbookEditor.Create("Sales", Rows);

        using var destination = new MemoryStream();
        await WorkbookEditor.CreateAsync("Sales", Rows, destination);

        var streamed = destination.ToArray();
        Assert.Equal(
            WorkbookEditor.ReadCell(expected, "Sales", "A1"),
            WorkbookEditor.ReadCell(streamed, "Sales", "A1"));
        Assert.Equal("1200", WorkbookEditor.ReadCell(streamed, "Sales", "B2"));
    }

    [Fact]
    public async Task WorkbookEditor_CreateAsync_WithSheets_DoesNotDisposeTheDestination()
    {
        var sheets = new[] { XlsxSheet.Named("Sales", new[] { new object?[] { "a", 1 } }) };

        using var destination = new ForwardOnlySink();
        await WorkbookEditor.CreateAsync(sheets, destination);

        Assert.False(destination.IsDisposed);
        Assert.True(destination.ToArray().Length > 0);
    }

    [Fact]
    public async Task WorkbookEditor_ReadCellAsync_MatchesTheByteArrayOverload()
    {
        using var source = StreamDoubles.Seekable(Xlsx);

        Assert.Equal(
            WorkbookEditor.ReadCell(Xlsx, "Sales", "B2"),
            await WorkbookEditor.ReadCellAsync(source, "Sales", "B2"));
    }

    [Fact]
    public async Task WorkbookEditor_SetCellAsync_MatchesTheByteArrayOverload()
    {
        var expected = WorkbookEditor.SetCell(Xlsx, "Sales", "B2", 1500);

        using var source = StreamDoubles.Seekable(Xlsx);
        using var destination = new MemoryStream();
        await WorkbookEditor.SetCellAsync(source, "Sales", "B2", 1500, destination);

        var streamed = destination.ToArray();
        Assert.Equal(
            WorkbookEditor.ReadCell(expected, "Sales", "B2"),
            WorkbookEditor.ReadCell(streamed, "Sales", "B2"));
        Assert.Equal("1500", WorkbookEditor.ReadCell(streamed, "Sales", "B2"));
        Assert.Equal("Region", WorkbookEditor.ReadCell(streamed, "Sales", "A1"));
    }

    [Fact]
    public async Task PresentationEditor_StreamOverloads_MatchTheByteArrayOverloads()
    {
        // B16: every assertion in this test used to compare a production method against itself.
        // SlideCount == SlideCount holds if both return 0; ExtractText == ExtractText holds if
        // both return an empty list. The literals below are what make the parity lines mean
        // something, and they are checked first for that reason.
        using var forCount = StreamDoubles.Seekable(Pptx);
        Assert.Equal(1, await PresentationEditor.SlideCountAsync(forCount));

        using var forCountParity = StreamDoubles.Seekable(Pptx);
        Assert.Equal(PresentationEditor.SlideCount(Pptx), await PresentationEditor.SlideCountAsync(forCountParity));

        using var forText = StreamDoubles.Seekable(Pptx);
        var streamedText = await PresentationEditor.ExtractTextAsync(forText);
        Assert.Contains("{{who}}", Assert.Single(streamedText), StringComparison.Ordinal);
        Assert.Equal(PresentationEditor.ExtractText(Pptx), streamedText);

        var expected = PresentationEditor.ReplaceText(Pptx, Replacements);
        using var forReplace = StreamDoubles.Seekable(Pptx);
        using var destination = new MemoryStream();
        await PresentationEditor.ReplaceTextAsync(forReplace, Replacements, destination);

        var replaced = PresentationEditor.ExtractText(destination.ToArray());
        Assert.Contains("World", Assert.Single(replaced), StringComparison.Ordinal);
        Assert.DoesNotContain("{{who}}", replaced[0], StringComparison.Ordinal);

        Assert.Equal(PresentationEditor.ExtractText(expected), replaced);
    }

    // =====================================================================================
    // Proof that the PDF path really streams
    // =====================================================================================

    /// <summary>
    /// The PDF is handed to the destination as it is produced, not assembled into one buffer and
    /// posted at the end.
    ///
    /// This is the assertion that a <c>byte[]</c> round trip wearing a <c>Stream</c> parameter
    /// cannot pass: writing a finished array costs exactly one write, whereas OfficeIMO's writer
    /// emits a document of this size in scores of them. It is also why
    /// <see cref="DocxToPdfConverter"/> is absent from <see cref="BufferedDestinationWriters"/> —
    /// those writes are synchronous, because OfficeIMO's stream writer is, and buffering the whole
    /// PDF to make them asynchronous would give up precisely the property under test here.
    /// </summary>
    [Fact]
    public async Task DocxToPdf_StreamsThePdfToTheDestinationInPieces_RatherThanBufferingItWhole()
    {
        var body = new StringBuilder();
        for (var i = 0; i < 2500; i++)
            body.Append("<p>Line ").Append(i).Append(" of a report long enough to need many pages.</p>");

        var docx = await HtmlToDocxConverter.ConvertAsync(body.ToString());

        using var source = StreamDoubles.Seekable(docx);
        var sink = new ForwardOnlySink();
        await DocxToPdfConverter.ConvertAsync(source, sink);

        Assert.True(sink.ToArray().Length > 100_000, $"expected a sizeable PDF, got {sink.ToArray().Length} bytes");
        Assert.True(sink.Writes > 10,
            $"The PDF reached the destination in {sink.Writes} write(s). One write means it was " +
            "materialised in full first, which is the byte[] behaviour these overloads exist to avoid.");
    }

    // =====================================================================================
    // Argument guards that are specific to one overload
    // =====================================================================================

    [Fact]
    public async Task HtmlToDocx_StreamOverload_RejectsNullHtml()
    {
        using var destination = new MemoryStream();
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToDocxConverter.ConvertAsync(null!, destination));
    }

    [Fact]
    public async Task HtmlToPdf_StreamOverload_RejectsNullHtml()
    {
        using var destination = new MemoryStream();
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => HtmlToPdfConverter.ConvertAsync(null!, destination));
    }

    [Fact]
    public async Task DocxEditor_ReplaceTextAsync_RejectsNullReplacements()
    {
        using var source = StreamDoubles.Seekable(Docx);
        using var destination = new MemoryStream();
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => DocxEditor.ReplaceTextAsync(source, null!, destination));
    }

    [Fact]
    public async Task PresentationEditor_ReplaceTextAsync_RejectsNullReplacements()
    {
        using var source = StreamDoubles.Seekable(Pptx);
        using var destination = new MemoryStream();
        await Assert.ThrowsAsync<ArgumentNullException>(
            () => PresentationEditor.ReplaceTextAsync(source, null!, destination));
    }

    [Fact]
    public async Task WorkbookEditor_CreateAsync_RejectsABlankSheetName()
    {
        using var destination = new MemoryStream();
        await Assert.ThrowsAsync<ArgumentException>(
            () => WorkbookEditor.CreateAsync(" ", Rows, destination));
    }

    [Fact]
    public async Task WorkbookEditor_ReadCellAsync_ReportsAMissingSheetAsAConversionFailure()
    {
        using var source = StreamDoubles.Seekable(Xlsx);

        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => WorkbookEditor.ReadCellAsync(source, "Nope", "A1"));

        Assert.Contains("Nope", ex.Message);
    }

    /// <summary>
    /// The Stream overloads agree with the <c>byte[]</c> overloads they shadow. Without this the
    /// theories above would be satisfied by an overload that guarded its streams correctly and then
    /// produced the wrong document.
    /// </summary>
    [Fact]
    public async Task PdfEditor_StreamOverloads_MatchTheByteArrayOverloads()
    {
        using var forCount = StreamDoubles.Seekable(TwoPagePdf);
        Assert.Equal(2, await PdfEditor.PageCountAsync(forCount));

        using var forExtract = StreamDoubles.Seekable(TwoPagePdf);
        using var extracted = new MemoryStream();
        await PdfEditor.ExtractPagesAsync(forExtract, 2, 1, extracted);
        Assert.Equal(
            PdfEditor.ExtractText(PdfEditor.ExtractPages(TwoPagePdf, 2, 1)),
            PdfEditor.ExtractText(extracted.ToArray()));
        Assert.Equal(1, PdfEditor.PageCount(extracted.ToArray()));

        using var forRemove = StreamDoubles.Seekable(TwoPagePdf);
        using var removed = new MemoryStream();
        await PdfEditor.RemovePagesAsync(forRemove, 1, 1, removed);
        Assert.Equal(1, PdfEditor.PageCount(removed.ToArray()));

        using var forInsert = StreamDoubles.Seekable(Pdf);
        using var insertSource = StreamDoubles.Seekable(TwoPagePdf);
        using var inserted = new MemoryStream();
        await PdfEditor.InsertPagesAsync(forInsert, insertSource, 1, inserted);
        Assert.Equal(3, PdfEditor.PageCount(inserted.ToArray()));
    }

    /// <summary>
    /// A <see cref="MemoryStream"/> source is consumed like any other stream.
    ///
    /// It used to take a fast path that called <c>ToArray()</c>: that returns the whole buffer
    /// regardless of <c>Position</c> and never advances the stream, so PdfEditor's two ways of
    /// reading a source disagreed about where reading starts — one of them silently handing back
    /// bytes the caller had already consumed. No theory above can see this, because they all pass a
    /// stream positioned at 0.
    /// </summary>
    [Fact]
    public async Task PdfEditor_ConsumesAMemoryStreamSource_RatherThanReadingItsBufferBehindItsBack()
    {
        using var source = new MemoryStream(Pdf);

        Assert.Equal(PdfEditor.PageCount(Pdf), await PdfEditor.PageCountAsync(source));
        Assert.Equal(source.Length, source.Position);
    }

    /// <summary>
    /// <c>InsertPagesAsync</c> is the one overload here with two sources, and the theory above can
    /// only drive one of them. The second gets the same three guards, named for itself.
    /// </summary>
    [Fact]
    public async Task PdfEditor_InsertPagesAsync_GuardsItsSecondSourceToo()
    {
        using var target = StreamDoubles.Seekable(Pdf);
        using var destination = new MemoryStream();

        await Assert.ThrowsAsync<ArgumentNullException>(
            () => PdfEditor.InsertPagesAsync(target, null!, 1, destination));

        using var forUnreadable = StreamDoubles.Seekable(Pdf);
        var unreadable = await Assert.ThrowsAsync<ArgumentException>(
            () => PdfEditor.InsertPagesAsync(forUnreadable, new NonReadableStream(), 1, destination));
        Assert.Equal("source", unreadable.ParamName);

        using var forEmpty = StreamDoubles.Seekable(Pdf);
        using var emptySource = new MemoryStream();
        var empty = await Assert.ThrowsAsync<ArgumentException>(
            () => PdfEditor.InsertPagesAsync(forEmpty, emptySource, 1, destination));
        Assert.Equal("source", empty.ParamName);
        Assert.Contains("empty", empty.Message, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Merging nothing names <c>sources</c> — the parameter the caller actually passed — rather
    /// than <c>pdfs</c>, which is the byte[] overload's parameter and is not in this signature.
    /// </summary>
    [Fact]
    public async Task PdfEditor_MergeAsync_NamesItsOwnParameterWhenGivenNothingToMerge()
    {
        using var destination = new MemoryStream();

        var ex = await Assert.ThrowsAsync<ArgumentException>(
            () => PdfEditor.MergeAsync(Array.Empty<Stream>(), destination));

        Assert.Equal("sources", ex.ParamName);
    }

    [Fact]
    public async Task DocxToPdf_ReportsARubbishSourceAsAConversionFailure()
    {
        using var source = StreamDoubles.Seekable(Encoding.ASCII.GetBytes("this is not a docx"));
        using var destination = new MemoryStream();

        await Assert.ThrowsAsync<DocumentConversionException>(
            () => DocxToPdfConverter.ConvertAsync(source, destination));
    }

    // =====================================================================================
    // Dispatch
    // =====================================================================================

    private static Task InvokeAsync(
        string api, Stream? source, Stream? destination, CancellationToken ct = default) => api switch
        {
            "HtmlToDocxConverter.ConvertAsync" =>
                HtmlToDocxConverter.ConvertAsync(Html, destination!, ct),
            "HtmlToDocxConverter.ConvertAsync(allowRemoteImageDownload)" =>
                HtmlToDocxConverter.ConvertAsync(Html, false, destination!, ct),
            "HtmlToDocxConverter.ConvertAsync(RemoteImageOptions)" =>
                HtmlToDocxConverter.ConvertAsync(Html, new RemoteImageOptions(), destination!, ct),
            "HtmlToDocxConverter.ConvertAsync(PageSetup)" =>
                HtmlToDocxConverter.ConvertAsync(Html, PageSetup.Letter, destination!, ct),
            "HtmlToPdfConverter.ConvertAsync" =>
                HtmlToPdfConverter.ConvertAsync(Html, destination!, ct),
            "HtmlToPdfConverter.ConvertAsync(allowRemoteImageDownload)" =>
                HtmlToPdfConverter.ConvertAsync(Html, false, destination!, ct),
            "HtmlToPdfConverter.ConvertAsync(RemoteImageOptions)" =>
                HtmlToPdfConverter.ConvertAsync(Html, new RemoteImageOptions(), destination!, ct),
            "HtmlToPdfConverter.ConvertAsync(PageSetup)" =>
                HtmlToPdfConverter.ConvertAsync(Html, PageSetup.Letter, destination!, ct),
            "DocxToPdfConverter.ConvertAsync" =>
                DocxToPdfConverter.ConvertAsync(source!, destination!, ct),
            "XlsxToPdfConverter.ConvertAsync" =>
                XlsxToPdfConverter.ConvertAsync(source!, destination!, ct),
            "PptxToPdfConverter.ConvertAsync" =>
                PptxToPdfConverter.ConvertAsync(source!, destination!, ct),
            "DocxEditor.ReplaceTextAsync" =>
                DocxEditor.ReplaceTextAsync(source!, Replacements, destination!, ct),
            "DocxEditor.FillRowsAsync" =>
                DocxEditor.FillRowsAsync(source!, "item", FillRowsRecords, destination!, ct),
            "DocxEditor.ReplaceImageAsync" =>
                DocxEditor.ReplaceImageAsync(source!, "{{logo}}", ImageFixtures.Png(), destination!, ct: ct),
            "DocxEditor.ExtractTextAsync" =>
                DocxEditor.ExtractTextAsync(source!, ct),
            "PdfEditor.ExtractTextAsync" =>
                PdfEditor.ExtractTextAsync(source!, ct),
            "DocxToHtmlConverter.ConvertAsync" =>
                DocxToHtmlConverter.ConvertAsync(source!, ct),
            "DocxToMarkdownConverter.ConvertAsync" =>
                DocxToMarkdownConverter.ConvertAsync(source!, ct),
            "DocxToHtmlConverter.ConvertWithReportAsync" =>
                DocxToHtmlConverter.ConvertWithReportAsync(source!, ct),
            "DocxToMarkdownConverter.ConvertWithReportAsync" =>
                DocxToMarkdownConverter.ConvertWithReportAsync(source!, ct),
            "XlsxToCsvConverter.ConvertAsync" =>
                XlsxToCsvConverter.ConvertAsync(source!, "Sales", ct),
            "XlsxToHtmlConverter.ConvertAsync" =>
                XlsxToHtmlConverter.ConvertAsync(source!, "Sales", ct),
            "DocxEditor.ExtractTextAsync(includeHeadersAndFooters)" =>
                DocxEditor.ExtractTextAsync(source!, true, ct),
            "DocxEditor.TableCountAsync" =>
                DocxEditor.TableCountAsync(source!, ct),
            "DocxEditor.ReadTableAsync" =>
                DocxEditor.ReadTableAsync(source!, 0, ct),
            "DocxEditor.CreateAsync" =>
                DocxEditor.CreateAsync(Blocks, destination!, ct),
            // PageSetup.Letter rather than A4: A4 is the default, so an arm passing it
            // would still pass if the parameter were ignored entirely.
            "DocxEditor.CreateAsync(PageSetup)" =>
                DocxEditor.CreateAsync(Blocks, PageSetup.Letter, destination!, ct),
            "WorkbookEditor.CreateAsync" =>
                WorkbookEditor.CreateAsync("Sales", Rows, destination!, ct),
            "WorkbookEditor.CreateAsync(sheets)" =>
                WorkbookEditor.CreateAsync(Sheets, destination!, ct),
            "WorkbookEditor.ReadCellAsync" =>
                WorkbookEditor.ReadCellAsync(source!, "Sales", "A1", ct),
            "WorkbookEditor.SheetNamesAsync" =>
                WorkbookEditor.SheetNamesAsync(source!, ct),
            "WorkbookEditor.ReadSheetAsync" =>
                WorkbookEditor.ReadSheetAsync(source!, "Sales", ct),
            "WorkbookEditor.SetCellAsync" =>
                WorkbookEditor.SetCellAsync(source!, "Sales", "B2", 1500, destination!, ct),
            "WorkbookEditor.AppendRowsAsync" =>
                WorkbookEditor.AppendRowsAsync(source!, "Sales", Rows, destination!, ct),
            "WorkbookEditor.FormatAsync" =>
                WorkbookEditor.FormatAsync(source!, "Sales", XlsxFormat.Report, destination!, ct),
            "PresentationEditor.SlideCountAsync" =>
                PresentationEditor.SlideCountAsync(source!, ct),
            "PresentationEditor.ExtractTextAsync" =>
                PresentationEditor.ExtractTextAsync(source!, ct),
            "PresentationEditor.ReplaceTextAsync" =>
                PresentationEditor.ReplaceTextAsync(source!, Replacements, destination!, ct),
            "PresentationEditor.ReplaceImageAsync" =>
                PresentationEditor.ReplaceImageAsync(source!, "{{chart}}", ImageFixtures.Png(), destination!, ct),
            "PresentationEditor.CreateAsync" =>
                PresentationEditor.CreateAsync(Slides, destination!, ct),
            "MarkdownToDocxConverter.ConvertAsync" =>
                MarkdownToDocxConverter.ConvertAsync(Md, destination!, ct),
            // Absent from BufferedDestinationWriters on purpose: like
            // DocxToPdfConverter, this hands the destination to OfficeIMO's own writer,
            // whose writes are synchronous - buffering to make them async would give up
            // the streaming this converter exists to preserve.
            "MarkdownToPdfConverter.ConvertAsync" =>
                MarkdownToPdfConverter.ConvertAsync(Md, destination!, ct),
            "PdfEditor.PageCountAsync" =>
                PdfEditor.PageCountAsync(source!, ct),
            // The theory drives ONE stream; MergeAsync takes a collection, so it gets a one-element
            // one. Merging a single document is a real call, not a degenerate case.
            "PdfEditor.MergeAsync" =>
                PdfEditor.MergeAsync(new[] { source! }, destination!, ct),
            "PdfEditor.ExtractPagesAsync" =>
                PdfEditor.ExtractPagesAsync(source!, 1, 1, destination!, ct),
            // Against TwoPagePdf: RemovePages refuses to remove every page.
            "PdfEditor.RemovePagesAsync" =>
                PdfEditor.RemovePagesAsync(source!, 1, 1, destination!, ct),
            "PdfEditor.RotatePagesAsync" =>
                PdfEditor.RotatePagesAsync(source!, 1, 1, 90, destination!, ct),
            "PdfEditor.ReorderPagesAsync" =>
                PdfEditor.ReorderPagesAsync(source!, new[] { 1 }, destination!, ct),
            // The theory's stream is the TARGET here — the first one read, and the one
            // SourceParamName reports. InsertPagesAsync's own `source` gets its guards from
            // PdfEditor_InsertPagesAsync_GuardsItsSecondSourceToo below.
            "PdfEditor.InsertPagesAsync" =>
                PdfEditor.InsertPagesAsync(source!, StreamDoubles.Seekable(Pdf), 1, destination!, ct),
            _ => throw new ArgumentOutOfRangeException(nameof(api), api, "Unknown Stream overload."),
        };

    /// <summary>The bytes an overload's <c>source</c> parameter expects, by format.</summary>
    private static byte[] SourceBytesFor(string api) => api switch
    {
        // FillRowsAsync throws unless the document holds a matching template row, and
        // ReadTableAsync throws unless index 0 exists, so neither can share the plain Docx
        // fixture the other DocxEditor overloads use.
        "DocxEditor.FillRowsAsync" => TableDocx,
        "DocxEditor.ReadTableAsync" => TableDocx,
        "DocxEditor.ReplaceImageAsync" => ImageDocx,
        "PresentationEditor.ReplaceImageAsync" => ImagePptx,
        "XlsxToPdfConverter.ConvertAsync" => Xlsx,
        "PptxToPdfConverter.ConvertAsync" => Pptx,
        "PdfEditor.RemovePagesAsync" => TwoPagePdf,
        _ when api.StartsWith("XlsxTo", StringComparison.Ordinal) => Xlsx,
        _ when api.StartsWith("PdfEditor", StringComparison.Ordinal) => Pdf,
        _ when api.StartsWith("WorkbookEditor", StringComparison.Ordinal) => Xlsx,
        _ when api.StartsWith("PresentationEditor", StringComparison.Ordinal) => Pptx,
        _ => Docx,
    };

    /// <summary>A fresh, valid source for <paramref name="api"/>, or an empty one if it takes none.</summary>
    private static MemoryStream NewSource(string api)
        => api.StartsWith("HtmlTo", StringComparison.Ordinal)
            || api == "WorkbookEditor.CreateAsync"
            || api == "WorkbookEditor.CreateAsync(sheets)"
            || api == "DocxEditor.CreateAsync"
            || api == "PresentationEditor.CreateAsync"
            || api == "MarkdownToDocxConverter.ConvertAsync"
            || api == "MarkdownToPdfConverter.ConvertAsync"
            ? new MemoryStream()
            : StreamDoubles.Seekable(SourceBytesFor(api));

    private static void AssertLooksLikeADocument(string api, byte[] written)
    {
        if (api.Contains("Pdf", StringComparison.Ordinal))
        {
            Assert.True(PdfProbe.IsPdf(written), $"{api} did not write a PDF.");
            return;
        }

        // An OOXML package is a ZIP: local file header magic "PK\x03\x04".
        Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, written.Take(4).ToArray());
    }

    /// <summary>The main document part's XML — the deterministic part of a .docx.</summary>
    private static string DocumentXml(byte[] docx)
        => DocxFixtures.Read(docx, main => main.Document!.OuterXml);

    /// <summary>
    /// Pins <see cref="DocumentXml"/> itself (B16). Every other use of it in this file compares
    /// <c>DocumentXml(a)</c> against <c>DocumentXml(b)</c> - the helper against itself - which
    /// holds however broken the helper is. If it returned an empty string, the byte[]-versus-Stream
    /// equivalence assertions it backs would all pass while proving nothing, and those assertions
    /// are the whole reason this file exists.
    ///
    /// Same shape as A26, which hid behind Assert.Contains for eight releases.
    /// </summary>
    [Fact]
    public async Task DocumentXml_ReturnsRealMarkup_SoTheEquivalenceAssertionsMeanSomething()
    {
        var xml = DocumentXml(await HtmlToDocxConverter.ConvertAsync(Html));

        Assert.Contains("<w:document", xml, StringComparison.Ordinal);
        Assert.Contains("<w:body", xml, StringComparison.Ordinal);
        Assert.True(xml.Length > 200, $"expected real markup, got {xml.Length} characters");
    }
}
