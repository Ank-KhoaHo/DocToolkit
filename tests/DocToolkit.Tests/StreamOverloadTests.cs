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

    /// <summary>
    /// No values. The shared <c>Docx</c> fixture carries no content controls at all, so filling it
    /// writes nothing and succeeds — which is what this suite wants, because it tests stream
    /// plumbing rather than form semantics.
    /// </summary>
    private static readonly IReadOnlyDictionary<string, DocxFormValue> NoFormValues =
        new Dictionary<string, DocxFormValue>();

    /// <summary>
    /// No values, deliberately. The shared <c>Docx</c> fixture carries <c>{{placeholder}}</c> text
    /// and not one MERGEFIELD, so a merge over it fills nothing and succeeds — which is what this
    /// suite wants, because it tests stream plumbing rather than merge semantics. A fixture WITH
    /// fields would make the strict overload throw before the plumbing was exercised at all.
    /// </summary>
    private static readonly IReadOnlyDictionary<string, string> NoMergeValues =
        new Dictionary<string, string>();

    /// <summary>
    /// No conditions, deliberately — the shared <c>Docx</c> fixture carries no <c>{{#Name}}</c>
    /// conditional marker at all, so resolving it against an empty dictionary changes nothing and
    /// succeeds, for the same plumbing-not-semantics reason <see cref="NoMergeValues"/> exists.
    /// </summary>
    private static readonly IReadOnlyDictionary<string, bool> NoConditions =
        new Dictionary<string, bool>();

    /// <summary>
    /// No regions, deliberately — the shared <c>Docx</c> fixture carries no <c>{{#each Name}}</c>
    /// repeating marker at all, so expanding it against an empty dictionary changes nothing and
    /// succeeds, for the same plumbing-not-semantics reason <see cref="NoConditions"/> exists.
    /// </summary>
    private static readonly IReadOnlyDictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>> NoRegions =
        new Dictionary<string, IEnumerable<IReadOnlyDictionary<string, string>>>();

    /// <summary>
    /// The nested twin of <see cref="NoRegions"/>, for the <c>MergeRepeatingRegions*</c> overloads,
    /// and empty for the same reason.
    /// </summary>
    private static readonly IReadOnlyDictionary<string, IEnumerable<DocxMailMergeBlockData>> NoBlockRegions =
        new Dictionary<string, IEnumerable<DocxMailMergeBlockData>>();

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

    /// <summary>
    /// A .docx whose only paragraph's text is exactly the placeholder, for AddTableOfContentsAsync
    /// — unlike AddFootnoteAsync/AddEndnoteAsync, which splice an inline reference the way
    /// ReplaceImageAsync does and so can share <see cref="ImageDocx"/>, AddTableOfContentsAsync
    /// replaces a whole paragraph and refuses one carrying anything besides the placeholder text.
    /// Built through raw OpenXml (via <see cref="DocxFixtures.Build"/>), not OfficeIMO, so this
    /// also exercises the "not OfficeIMO-authored" composition the ticket's fix specifically
    /// closed (no mc:Ignorable="w14 ..." declared at the document root).
    /// </summary>
    private static readonly byte[] TocDocx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("{{toc}}")));

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

    /// <summary>A two-slide deck, for overloads a one-slide source can't meaningfully exercise (ReorderSlidesAsync's permutation, RemoveSlidesAsync's non-empty-set removal).</summary>
    private static readonly byte[] MultiSlidePptx =
        PptxFixtures.MultiSlideDeck(new[] { "Slide 1", "Slide 2" }, reverseDeckOrder: false);

    /// <summary>
    /// A real Word 97-2003 binary .doc, for the DocToDocxConverter overloads.
    ///
    /// <b>The LOSSLESS fixture deliberately, not the one with a table.</b> A .doc holding a table
    /// carries a binary stream a .docx cannot take, so ConvertAsync refuses it without an explicit
    /// opt-in - and every theory in this file drives the overload with no options and expects it to
    /// succeed. Using the blocking fixture would fail them for a reason that has nothing to do with
    /// stream handling, which is all this suite is about.
    /// </summary>
    private static readonly byte[] LegacyDoc =
        File.ReadAllBytes(Path.Join(AppContext.BaseDirectory, "assets", "legacy-lossless.doc"));

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

    /// <summary>
    /// An ENCRYPTED .pdf, for UnprotectAsync — the one overload here whose input must already be
    /// protected. Feeding it the plain <see cref="Pdf"/> would fail for a reason that has nothing
    /// to do with stream handling, which is all this suite tests.
    ///
    /// Declared after <see cref="Pdf"/> because static field initialisers run in declaration order.
    /// </summary>
    private static readonly byte[] ProtectedPdf =
        PdfEditor.Protect(Pdf, new PdfProtection { UserPassword = "pw" });

    /// <summary>
    /// Encrypted Office documents, for the three Unprotect overloads. Each must be encrypted with
    /// the same password the dispatch table passes, or the overload fails for a reason that has
    /// nothing to do with stream handling.
    /// </summary>
    private static readonly byte[] ProtectedDocx = DocxEditor.Protect(Docx, "pw");
    private static readonly byte[] ProtectedXlsx = WorkbookEditor.Protect(Xlsx, "pw");
    private static readonly byte[] ProtectedPptx = PresentationEditor.Protect(Pptx, "pw");

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
        "HtmlToDocxConverter.ConvertAsync(PageSetup, RemoteImageOptions)",
        "HtmlToPdfConverter.ConvertAsync",
        "HtmlToPdfConverter.ConvertAsync(allowRemoteImageDownload)",
        "HtmlToPdfConverter.ConvertAsync(RemoteImageOptions)",
        "HtmlToPdfConverter.ConvertAsync(PageSetup)",
        "HtmlToPdfConverter.ConvertAsync(PageSetup, RemoteImageOptions)",
        "HtmlToPdfConverter.ConvertAsync(HtmlToPdfOptions)",
        "DocxToPdfConverter.ConvertAsync",
        "DocxToPdfConverter.ConvertAsync(PdfFontOptions)",
        "XlsxToPdfConverter.ConvertAsync",
        "PptxToPdfConverter.ConvertAsync",
        "DocxEditor.ReplaceTextAsync",
        "DocxEditor.FillRowsAsync",
        "DocxEditor.ReplaceImageAsync",
        "DocxEditor.AddFootnoteAsync",
        "DocxEditor.AddEndnoteAsync",
        "DocxEditor.AddTableOfContentsAsync",
        "DocxEditor.CreateAsync",
        "DocxEditor.CreateAsync(PageSetup)",
        "WorkbookEditor.CreateAsync",
        "WorkbookEditor.CreateAsync(sheets)",
        "WorkbookEditor.SetCellAsync",
        "WorkbookEditor.AppendRowsAsync",
        "WorkbookEditor.FormatAsync",
        "WorkbookEditor.AddChartAsync",
        "PresentationEditor.ReplaceTextAsync",
        "PresentationEditor.InsertSlidesAsync",
        "PresentationEditor.ReorderSlidesAsync",
        "PresentationEditor.RemoveSlidesAsync",
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
        "DocToDocxConverter.ConvertAsync",
        "DocToDocxConverter.ConvertAsync(LegacyDocOptions)",
        "PdfEditor.ProtectAsync",
        "PdfEditor.UnprotectAsync",
        "DocxEditor.ProtectAsync",
        "DocxEditor.UnprotectAsync",
        "WorkbookEditor.ProtectAsync",
        "WorkbookEditor.UnprotectAsync",
        "PresentationEditor.ProtectAsync",
        "PresentationEditor.UnprotectAsync",
        "DocxReview.RemoveCommentsAsync",
        "DocxReview.AcceptRevisionsAsync",
        "DocxReview.RejectRevisionsAsync",
        "DocxMailMerge.MergeAsync",
        "DocxMailMerge.MergeWithReportAsync",
        "DocxMailMerge.MergeConditionalAsync",
        "DocxMailMerge.MergeConditionalWithReportAsync",
        "DocxMailMerge.MergeRepeatingAsync",
        "DocxMailMerge.MergeRepeatingWithReportAsync",
        "DocxMailMerge.MergeRepeatingRegionsAsync",
        "DocxMailMerge.MergeRepeatingRegionsWithReportAsync",
        "DocxMailMerge.MergeTableRowsAsync",
        "DocxMailMerge.MergeTableRowGroupsAsync",
        "DocxForm.FillAsync",
    };

    /// <summary>Overloads that take a <c>Stream source</c>.</summary>
    private static readonly string[] SourceReaderNames =
    {
        "DocxToPdfConverter.ConvertAsync",
        "DocxToPdfConverter.ConvertAsync(PdfFontOptions)",
        "DocxEditor.ReplaceTextAsync",
        "DocxEditor.FillRowsAsync",
        "DocxEditor.ReplaceImageAsync",
        "DocxEditor.AddFootnoteAsync",
        "DocxEditor.AddEndnoteAsync",
        "DocxEditor.AddTableOfContentsAsync",
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
        "WorkbookEditor.AddChartAsync",
        "PresentationEditor.SlideCountAsync",
        "PresentationEditor.ReadSlideAsync",
        "PresentationEditor.ReadSmartArtAsync",
        "PresentationEditor.ExtractTextAsync",
        "PresentationEditor.ReplaceTextAsync",
        "PresentationEditor.InsertSlidesAsync",
        "PresentationEditor.ReorderSlidesAsync",
        "PresentationEditor.RemoveSlidesAsync",
        "PresentationEditor.ReplaceImageAsync",
        "PdfEditor.PageCountAsync",
        "PdfEditor.MergeAsync",
        "PdfEditor.ExtractPagesAsync",
        "PdfEditor.RemovePagesAsync",
        "PdfEditor.RotatePagesAsync",
        "PdfEditor.ReorderPagesAsync",
        "PdfEditor.InsertPagesAsync",
        "DocToDocxConverter.ConvertAsync",
        "DocToDocxConverter.ConvertAsync(LegacyDocOptions)",
        "DocToDocxConverter.ExtractTextAsync",
        "PdfEditor.ProtectAsync",
        "PdfEditor.UnprotectAsync",
        "DocxEditor.ProtectAsync",
        "DocxEditor.UnprotectAsync",
        "WorkbookEditor.ProtectAsync",
        "WorkbookEditor.UnprotectAsync",
        "PresentationEditor.ProtectAsync",
        "PresentationEditor.UnprotectAsync",
        "DocxReview.InspectAsync",
        "DocxReview.RemoveCommentsAsync",
        "DocxReview.AcceptRevisionsAsync",
        "DocxReview.RejectRevisionsAsync",
        "DocxToPdfPreflight.InspectAsync",
        "DocxMailMerge.InspectTemplateAsync",
        "DocxMailMerge.MergeAsync",
        "DocxMailMerge.MergeWithReportAsync",
        "DocxMailMerge.MergeConditionalAsync",
        "DocxMailMerge.MergeConditionalWithReportAsync",
        "DocxMailMerge.MergeRepeatingAsync",
        "DocxMailMerge.MergeRepeatingWithReportAsync",
        "DocxMailMerge.MergeRepeatingRegionsAsync",
        "DocxMailMerge.MergeRepeatingRegionsWithReportAsync",
        "DocxMailMerge.MergeTableRowsAsync",
        "DocxMailMerge.MergeTableRowGroupsAsync",
        "DocxForm.InspectAsync",
        "DocxForm.ValidateAsync",
        "DocxForm.FillAsync",
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
        "HtmlToDocxConverter.ConvertAsync(PageSetup, RemoteImageOptions)",
        "DocxEditor.ReplaceTextAsync",
        "DocxEditor.FillRowsAsync",
        "DocxEditor.ReplaceImageAsync",
        "DocxEditor.AddFootnoteAsync",
        "DocxEditor.AddEndnoteAsync",
        "DocxEditor.AddTableOfContentsAsync",
        "DocxEditor.CreateAsync",
        "DocxEditor.CreateAsync(PageSetup)",
        "WorkbookEditor.CreateAsync",
        "WorkbookEditor.CreateAsync(sheets)",
        "WorkbookEditor.SetCellAsync",
        "WorkbookEditor.AppendRowsAsync",
        "WorkbookEditor.FormatAsync",
        "WorkbookEditor.AddChartAsync",
        "PresentationEditor.ReplaceTextAsync",
        "PresentationEditor.InsertSlidesAsync",
        "PresentationEditor.ReorderSlidesAsync",
        "PresentationEditor.RemoveSlidesAsync",
        "PresentationEditor.ReplaceImageAsync",
        "PresentationEditor.CreateAsync",
        "MarkdownToDocxConverter.ConvertAsync",
        "PdfEditor.MergeAsync",
        "PdfEditor.ExtractPagesAsync",
        "PdfEditor.RemovePagesAsync",
        "PdfEditor.RotatePagesAsync",
        "PdfEditor.ReorderPagesAsync",
        "PdfEditor.InsertPagesAsync",
        "DocToDocxConverter.ConvertAsync",
        "DocToDocxConverter.ConvertAsync(LegacyDocOptions)",
        "PdfEditor.ProtectAsync",
        "PdfEditor.UnprotectAsync",
        "DocxEditor.ProtectAsync",
        "DocxEditor.UnprotectAsync",
        "WorkbookEditor.ProtectAsync",
        "WorkbookEditor.UnprotectAsync",
        "PresentationEditor.ProtectAsync",
        "PresentationEditor.UnprotectAsync",
        "DocxReview.RemoveCommentsAsync",
        "DocxReview.AcceptRevisionsAsync",
        "DocxReview.RejectRevisionsAsync",
        "DocxMailMerge.MergeAsync",
        "DocxMailMerge.MergeWithReportAsync",
        "DocxMailMerge.MergeConditionalAsync",
        "DocxMailMerge.MergeConditionalWithReportAsync",
        "DocxMailMerge.MergeRepeatingAsync",
        "DocxMailMerge.MergeRepeatingWithReportAsync",
        "DocxMailMerge.MergeRepeatingRegionsAsync",
        "DocxMailMerge.MergeRepeatingRegionsWithReportAsync",
        "DocxMailMerge.MergeTableRowsAsync",
        "DocxMailMerge.MergeTableRowGroupsAsync",
        "DocxForm.FillAsync",
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
    /// <b>It counts overloads, not just names, and it did not always.</b> Matching on
    /// <c>Class.Method</c> alone left "a NEW OVERLOAD of an already-listed name" invisible — a
    /// limitation this comment used to state as a deliberate choice. It stopped being defensible on
    /// 2026-08-27, when mutation testing found that
    /// <c>HtmlToDocxConverter.ConvertAsync(html, PageSetup, RemoteImageOptions, destination, ct)</c>
    /// was in neither list: <b>six guards on it survived every test in the repository</b>, including
    /// the <c>RequireWritable</c> and already-cancelled-token theories that exist to cover exactly
    /// those. The name <c>HtmlToDocxConverter.ConvertAsync</c> was listed four times, so the check
    /// was satisfied while a fifth overload went untested.
    ///
    /// <para>So the count per method name must match. A stated limitation that has since cost real
    /// coverage is a bug wearing a comment.</para>
    /// </summary>
    [Fact]
    public void EveryPublicStreamOverloadIsRegisteredInTheListsAbove()
    {
        // DISTINCT entries per method name: the two destination lists deliberately overlap, so a
        // name appearing in both is one registered overload rather than two. Counting raw entries
        // would let a duplicate stand in for a missing overload.
        var listedByMethod = DestinationWriterNames
            .Concat(SourceReaderNames)
            .Concat(BufferedDestinationWriterNames)
            .Distinct(StringComparer.Ordinal)
            .GroupBy(n => n.Split('(')[0], StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.Count(), StringComparer.Ordinal);

        // ALL SIX SHIPPED ASSEMBLIES, not just the one this file's oldest entries came from.
        // Until the per-concern project split this was `typeof(DocxEditor).Assembly` and that WAS
        // the whole library; afterwards it was one assembly of six, so a new Stream overload on
        // PdfEditor, WorkbookEditor or PresentationEditor could go unregistered and every theory
        // here would skip it in silence - reopening the exact hole this test exists to close.
        // Same list, and the same reasoning, as ArgumentExceptionNamesADeclaredParameterTests.
        var assemblies = new[]
        {
            typeof(HtmlToPdfConverter).Assembly,   // DocToolkit
            typeof(PageSetup).Assembly,            // DocToolkit.Primitives
            typeof(DocxEditor).Assembly,           // DocToolkit.Docx
            typeof(WorkbookEditor).Assembly,       // DocToolkit.Xlsx
            typeof(PresentationEditor).Assembly,   // DocToolkit.Pptx
            typeof(PdfEditor).Assembly,            // DocToolkit.Pdf
        }.Distinct();

        var shipped =
            from type in assemblies.SelectMany(a => a.GetExportedTypes())
            where type is { IsAbstract: true, IsSealed: true }        // C# static class
            from method in type.GetMethods(BindingFlags.Public | BindingFlags.Static
                                           | BindingFlags.DeclaredOnly)
            where method.GetParameters().Any(p => typeof(Stream).IsAssignableFrom(p.ParameterType))
            select $"{type.Name}.{method.Name}";

        var shippedByMethod = shipped
            .GroupBy(m => m, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.Count(), StringComparer.Ordinal);

        var gaps = shippedByMethod
            .Select(kv => (Method: kv.Key, Shipped: kv.Value,
                           Listed: listedByMethod.GetValueOrDefault(kv.Key)))
            .Where(x => x.Listed < x.Shipped)
            .OrderBy(x => x.Method, StringComparer.Ordinal)
            .ToList();

        Assert.True(gaps.Count == 0,
            "These public Stream overloads are not all registered in this file's name lists, so "
            + "every theory here skips the unregistered ones silently:\n  "
            + string.Join("\n  ", gaps.Select(
                g => $"{g.Method}: {g.Shipped} shipped, {g.Listed} listed")));
    }

    /// <summary>
    /// A PdfFontOptions carrying bytes that are NOT a real TrueType file.
    /// </summary>
    /// <remarks>
    /// These theories assert plumbing - that a destination is written, refused or cancelled - not
    /// that a glyph rendered. There is no font asset in this repository and no test converts with a
    /// real one, so a synthetic instance is what registers this overload. If it ever turns out the
    /// renderer rejects it, that is a finding about the renderer worth its own test, not a reason
    /// to leave the overload unregistered again.
    /// </remarks>
    private static PdfFontOptions SampleFonts => new("DocToolkitTest", [0x00, 0x01, 0x00, 0x00]);

    /// <summary>
    /// APIs whose theories cannot assert a successful CONVERSION, and why.
    /// </summary>
    /// <remarks>
    /// <b>An exemption stated in one place beats an overload left unregistered.</b> Registering
    /// <c>DocxToPdfConverter.ConvertAsync(PdfFontOptions)</c> is what gives it the guard theories —
    /// a refused destination, an already-cancelled token, an unreadable source — which is what
    /// mutation testing showed were missing. What it cannot have is the two theories that require
    /// bytes to come out the other end.
    ///
    /// <para>The reason is measured, not assumed: a <c>PdfFontOptions</c> must carry a real
    /// TrueType file, and a synthetic one fails with <c>NotSupportedException: TrueType font data
    /// is too small to embed</c>, wrapped as a <c>DocumentConversionException</c>. That is the
    /// renderer behaving correctly. This repository ships no font asset, and fabricating a valid
    /// TTF to satisfy a plumbing test would be a worse trade than naming the gap.</para>
    ///
    /// <para><b>If a font asset is ever added, delete this set rather than growing it.</b></para>
    /// </remarks>
    private static readonly HashSet<string> CannotAssertARenderedDocument = new(StringComparer.Ordinal)
    {
        "DocxToPdfConverter.ConvertAsync(PdfFontOptions)",
    };

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
        // See CannotAssertARenderedDocument: this one is registered for its guard
        // theories, and cannot produce a document from synthetic test input.
        if (CannotAssertARenderedDocument.Contains(api)) return;

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
        // See CannotAssertARenderedDocument: this one is registered for its guard
        // theories, and cannot produce a document from synthetic test input.
        if (CannotAssertARenderedDocument.Contains(api)) return;

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
    /// The PDF is rendered whole and then emitted, and the destination is only ever touched by a
    /// conversion that succeeded.
    /// </summary>
    /// <remarks>
    /// <b>This test asserted the opposite until 2026-08-20, and it was right about the old
    /// contract.</b> It required the PDF to reach the destination in many writes, because a
    /// <c>byte[]</c> round trip wearing a <c>Stream</c> signature costs exactly one - a real hazard
    /// this file exists to catch.
    ///
    /// <b>The contract changed on measurement, not on taste.</b> Writing straight through means a
    /// repair cannot retry a failed render - you cannot un-write bytes already on somebody's
    /// response body - so the stream overloads applied NO repairs. Measured over real files: 4 of
    /// 99 real Word documents converted through <c>Convert(byte[])</c> and were refused here, and
    /// on the HTML path a construct present in 27 of 181 real .gov pages did the same. The
    /// maintainer chose parity over streaming.
    ///
    /// <b>So the assertion is inverted rather than deleted</b>, and the write count is still
    /// load-bearing - it now pins the buffering that makes the repairs possible. What replaces the
    /// old guarantee is stronger for a caller: a failure leaves the destination untouched instead
    /// of carrying a truncated PDF. <see cref="StreamPathParityTests"/> holds that half.
    ///
    /// <b>This does not license buffering elsewhere.</b> Every other <c>Stream</c> overload in this
    /// suite is still held to <see cref="BufferedDestinationWriters"/>' rules; the PDF render is the
    /// one place a documented property was traded, with a number attached.
    /// </remarks>
    [Fact]
    public async Task DocxToPdf_RendersWholeThenEmits_SoARetryCanRepairAFailedRender()
    {
        var body = new StringBuilder();
        for (var i = 0; i < 2500; i++)
            body.Append("<p>Line ").Append(i).Append(" of a report long enough to need many pages.</p>");

        var docx = await HtmlToDocxConverter.ConvertAsync(body.ToString());

        using var source = StreamDoubles.Seekable(docx);
        var sink = new ForwardOnlySink();
        await DocxToPdfConverter.ConvertAsync(source, sink);

        Assert.True(sink.ToArray().Length > 100_000, $"expected a sizeable PDF, got {sink.ToArray().Length} bytes");
        Assert.True(PdfProbe.IsPdf(sink.ToArray()));
        Assert.False(sink.IsDisposed, "ConvertAsync disposed a destination it does not own");
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
            "DocxReview.InspectAsync" =>
                DocxReview.InspectAsync(source!, ct),
            "DocxToPdfPreflight.InspectAsync" =>
                DocxToPdfPreflight.InspectAsync(source!, ct),
            "DocxMailMerge.InspectTemplateAsync" =>
                DocxMailMerge.InspectTemplateAsync(source!, ct),
            "DocxForm.InspectAsync" =>
                DocxForm.InspectAsync(source!, ct: ct),
            "DocxForm.ValidateAsync" =>
                DocxForm.ValidateAsync(source!, NoFormValues, ct: ct),
            "DocxForm.FillAsync" =>
                DocxForm.FillAsync(source!, destination!, NoFormValues, ct: ct),
            "DocxMailMerge.MergeAsync" =>
                DocxMailMerge.MergeAsync(source!, destination!, NoMergeValues, ct),
            "DocxMailMerge.MergeWithReportAsync" =>
                DocxMailMerge.MergeWithReportAsync(source!, destination!, NoMergeValues, ct),
            "DocxMailMerge.MergeConditionalAsync" =>
                DocxMailMerge.MergeConditionalAsync(source!, destination!, NoConditions, ct),
            "DocxMailMerge.MergeConditionalWithReportAsync" =>
                DocxMailMerge.MergeConditionalWithReportAsync(source!, destination!, NoConditions, ct),
            "DocxMailMerge.MergeRepeatingAsync" =>
                DocxMailMerge.MergeRepeatingAsync(source!, destination!, NoRegions, ct),
            "DocxMailMerge.MergeRepeatingWithReportAsync" =>
                DocxMailMerge.MergeRepeatingWithReportAsync(source!, destination!, NoRegions, ct),
            "DocxMailMerge.MergeRepeatingRegionsAsync" =>
                DocxMailMerge.MergeRepeatingRegionsAsync(source!, destination!, NoBlockRegions, ct),
            "DocxMailMerge.MergeRepeatingRegionsWithReportAsync" =>
                DocxMailMerge.MergeRepeatingRegionsWithReportAsync(source!, destination!, NoBlockRegions, ct),
            "DocxMailMerge.MergeTableRowsAsync" =>
                DocxMailMerge.MergeTableRowsAsync(source!, destination!, 0, 1, Array.Empty<IReadOnlyDictionary<string, string>>(), ct),
            "DocxMailMerge.MergeTableRowGroupsAsync" =>
                DocxMailMerge.MergeTableRowGroupsAsync(source!, destination!, 0, 0, 1, Array.Empty<DocxMailMergeTableRowGroup>(), ct),
            "DocxReview.RemoveCommentsAsync" =>
                DocxReview.RemoveCommentsAsync(source!, destination!, ct),
            "DocxReview.AcceptRevisionsAsync" =>
                DocxReview.AcceptRevisionsAsync(source!, destination!, ct),
            "DocxReview.RejectRevisionsAsync" =>
                DocxReview.RejectRevisionsAsync(source!, destination!, ct),
            "DocxEditor.ProtectAsync" =>
                DocxEditor.ProtectAsync(source!, destination!, "pw", ct),
            "DocxEditor.UnprotectAsync" =>
                DocxEditor.UnprotectAsync(source!, destination!, "pw", ct),
            "WorkbookEditor.ProtectAsync" =>
                WorkbookEditor.ProtectAsync(source!, destination!, "pw", ct),
            "WorkbookEditor.UnprotectAsync" =>
                WorkbookEditor.UnprotectAsync(source!, destination!, "pw", ct),
            "PresentationEditor.ProtectAsync" =>
                PresentationEditor.ProtectAsync(source!, destination!, "pw", ct),
            "PresentationEditor.UnprotectAsync" =>
                PresentationEditor.UnprotectAsync(source!, destination!, "pw", ct),
            "PdfEditor.ProtectAsync" =>
                PdfEditor.ProtectAsync(source!, destination!, new PdfProtection { UserPassword = "pw" }, ct),
            "PdfEditor.UnprotectAsync" =>
                PdfEditor.UnprotectAsync(source!, destination!, "pw", ct),
            "DocToDocxConverter.ConvertAsync" =>
                DocToDocxConverter.ConvertAsync(source!, destination!, ct),
            "DocToDocxConverter.ConvertAsync(LegacyDocOptions)" =>
                DocToDocxConverter.ConvertAsync(source!, destination!, new LegacyDocOptions(), ct),
            "DocToDocxConverter.ExtractTextAsync" =>
                DocToDocxConverter.ExtractTextAsync(source!, ct),
            "HtmlToDocxConverter.ConvertAsync" =>
                HtmlToDocxConverter.ConvertAsync(Html, destination!, ct),
            "HtmlToDocxConverter.ConvertAsync(allowRemoteImageDownload)" =>
                HtmlToDocxConverter.ConvertAsync(Html, false, destination!, ct),
            "HtmlToDocxConverter.ConvertAsync(RemoteImageOptions)" =>
                HtmlToDocxConverter.ConvertAsync(Html, new RemoteImageOptions(), destination!, ct),
            "HtmlToDocxConverter.ConvertAsync(PageSetup)" =>
                HtmlToDocxConverter.ConvertAsync(Html, PageSetup.Letter, destination!, ct),
            "HtmlToDocxConverter.ConvertAsync(PageSetup, RemoteImageOptions)" =>
                HtmlToDocxConverter.ConvertAsync(
                    Html, PageSetup.Letter, new RemoteImageOptions(), destination!, ct),
            "HtmlToPdfConverter.ConvertAsync" =>
                HtmlToPdfConverter.ConvertAsync(Html, destination!, ct),
            "HtmlToPdfConverter.ConvertAsync(allowRemoteImageDownload)" =>
                HtmlToPdfConverter.ConvertAsync(Html, false, destination!, ct),
            "HtmlToPdfConverter.ConvertAsync(RemoteImageOptions)" =>
                HtmlToPdfConverter.ConvertAsync(Html, new RemoteImageOptions(), destination!, ct),
            "HtmlToPdfConverter.ConvertAsync(PageSetup)" =>
                HtmlToPdfConverter.ConvertAsync(Html, PageSetup.Letter, destination!, ct),
            "HtmlToPdfConverter.ConvertAsync(PageSetup, RemoteImageOptions)" =>
                HtmlToPdfConverter.ConvertAsync(
                    Html, PageSetup.Letter, new RemoteImageOptions(), destination!, ct),
            "HtmlToPdfConverter.ConvertAsync(HtmlToPdfOptions)" =>
                HtmlToPdfConverter.ConvertAsync(
                    Html, new HtmlToPdfOptions { Page = PageSetup.Letter }, destination!, ct),
            "DocxToPdfConverter.ConvertAsync" =>
                DocxToPdfConverter.ConvertAsync(source!, destination!, ct),
            "DocxToPdfConverter.ConvertAsync(PdfFontOptions)" =>
                DocxToPdfConverter.ConvertAsync(source!, destination!, SampleFonts, ct),
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
            "DocxEditor.AddFootnoteAsync" =>
                DocxEditor.AddFootnoteAsync(source!, "{{logo}}", "A footnote.", destination!, ct),
            "DocxEditor.AddEndnoteAsync" =>
                DocxEditor.AddEndnoteAsync(source!, "{{logo}}", "An endnote.", destination!, ct),
            "DocxEditor.AddTableOfContentsAsync" =>
                DocxEditor.AddTableOfContentsAsync(source!, "{{toc}}", destination!, ct: ct),
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
            "WorkbookEditor.AddChartAsync" =>
                WorkbookEditor.AddChartAsync(
                    source!, "Sales", "B2", ChartType.Line,
                    new ChartData(new[] { "A" }, new[] { new ChartSeries("S", new double[] { 1 }) }),
                    destination!, ct: ct),
            "PresentationEditor.SlideCountAsync" =>
                PresentationEditor.SlideCountAsync(source!, ct),
            "PresentationEditor.ReadSlideAsync" =>
                PresentationEditor.ReadSlideAsync(source!, 1, ct),
            "PresentationEditor.ReadSmartArtAsync" =>
                PresentationEditor.ReadSmartArtAsync(source!, 1, ct),
            "PresentationEditor.ExtractTextAsync" =>
                PresentationEditor.ExtractTextAsync(source!, ct),
            "PresentationEditor.ReplaceTextAsync" =>
                PresentationEditor.ReplaceTextAsync(source!, Replacements, destination!, ct),
            "PresentationEditor.InsertSlidesAsync" =>
                PresentationEditor.InsertSlidesAsync(
                    source!, 1, new[] { PptxSlide.Titled("New") }, destination!, ct),
            "PresentationEditor.ReorderSlidesAsync" =>
                PresentationEditor.ReorderSlidesAsync(source!, new[] { 2, 1 }, destination!, ct),
            "PresentationEditor.RemoveSlidesAsync" =>
                PresentationEditor.RemoveSlidesAsync(source!, new[] { 1 }, destination!, ct),
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
        // fixture the other DocxEditor overloads use. MergeTableRowsAsync and
        // MergeTableRowGroupsAsync have the same requirement.
        "DocxEditor.FillRowsAsync" => TableDocx,
        "DocxEditor.ReadTableAsync" => TableDocx,
        "DocxMailMerge.MergeTableRowsAsync" => TableDocx,
        "DocxMailMerge.MergeTableRowGroupsAsync" => TableDocx,
        "DocxEditor.ReplaceImageAsync" => ImageDocx,
        "DocxEditor.AddFootnoteAsync" => ImageDocx,
        "DocxEditor.AddEndnoteAsync" => ImageDocx,
        "DocxEditor.AddTableOfContentsAsync" => TocDocx,
        "PresentationEditor.ReplaceImageAsync" => ImagePptx,
        "XlsxToPdfConverter.ConvertAsync" => Xlsx,
        "PptxToPdfConverter.ConvertAsync" => Pptx,
        "PdfEditor.RemovePagesAsync" => TwoPagePdf,
        "PdfEditor.UnprotectAsync" => ProtectedPdf,
        "DocxEditor.UnprotectAsync" => ProtectedDocx,
        "WorkbookEditor.UnprotectAsync" => ProtectedXlsx,
        "PresentationEditor.UnprotectAsync" => ProtectedPptx,
        "PresentationEditor.ReorderSlidesAsync" => MultiSlidePptx,
        "PresentationEditor.RemoveSlidesAsync" => MultiSlidePptx,
        _ when api.StartsWith("DocToDocxConverter", StringComparison.Ordinal) => LegacyDoc,
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

        // An ENCRYPTED Office document is not a package at all: the ZIP is sealed inside a
        // compound file (D0 CF 11 E0). Asserting "PK" for these would be asserting that the
        // encryption did not happen - so this checks the compound-file signature instead, which is
        // still a real shape assertion rather than an exemption.
        if (api.StartsWith("DocxEditor.Protect", StringComparison.Ordinal)
            || api.StartsWith("WorkbookEditor.Protect", StringComparison.Ordinal)
            || api.StartsWith("PresentationEditor.Protect", StringComparison.Ordinal))
        {
            Assert.Equal(new byte[] { 0xD0, 0xCF, 0x11, 0xE0 }, written.Take(4).ToArray());
            return;
        }

        // An OOXML package is a ZIP: local file header magic "PK\x03\x04". The Unprotect overloads
        // land here deliberately - taking encryption off must give a real package back.
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

    /// <summary>
    /// A source that dies <b>mid-read</b> surfaces as <see cref="DocumentConversionException"/>
    /// with the original failure preserved as <c>InnerException</c>.
    ///
    /// This is <c>StreamPipeline.DrainAsync</c>'s generic <c>catch (Exception)</c> arm, and before
    /// this test <b>no test in the suite reached it</b> — B14's widened mutation scope reported its
    /// mutants as <c>NoCoverage</c> rather than merely surviving, which is the stronger signal of
    /// the two. Every other stream double refuses up front (<c>CanRead</c> false), which the
    /// <c>RequireReadable</c> guard catches before a transfer begins; this one fails part-way
    /// through a copy that has already moved bytes.
    ///
    /// The contract being pinned is the public one every reader documents: failures arrive as one
    /// exception type, with the cause intact for a log.
    /// </summary>
    [Fact]
    public async Task ASourceThatDiesMidReadSurfacesAsDocumentConversionException()
    {
        using var source = new FailsPartWayStream(Docx, failAfter: 16);

        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => DocxEditor.ExtractTextAsync(source));

        Assert.IsType<IOException>(ex.InnerException);
    }

    /// <summary>
    /// The same arm on the way out — <c>StreamPipeline.EmitAsync</c>, whose two mutants were also
    /// <c>NoCoverage</c>. A destination that fails part-way is an HTTP response body whose client
    /// hung up, which is the ordinary case this path exists for rather than an exotic one.
    /// </summary>
    [Fact]
    public async Task ADestinationThatDiesMidWriteSurfacesAsDocumentConversionException()
    {
        using var destination = new FailsPartWayStream(bytes: null, failAfter: 16);

        var ex = await Assert.ThrowsAsync<DocumentConversionException>(
            () => HtmlToDocxConverter.ConvertAsync(Html, destination));

        Assert.IsType<IOException>(ex.InnerException);
    }

    /// <summary>
    /// Cancellation arriving <b>during</b> the write surfaces as
    /// <see cref="OperationCanceledException"/>, not wrapped in
    /// <see cref="DocumentConversionException"/>.
    ///
    /// That distinction is the contract: a cancelled operation is the caller getting what they
    /// asked for, and a caller who wrote <c>catch (DocumentConversionException)</c> around a
    /// conversion must not have their own <c>CancellationToken</c> come back to them disguised as
    /// a document failure.
    ///
    /// <c>EveryOverload_ThrowsForAnAlreadyCancelledToken</c> does not reach this arm — an
    /// already-cancelled token is refused by the guard at the top of the overload, before any
    /// transfer starts. Only cancellation mid-write gets here, and B14's widened mutation scope
    /// reported the rethrow as <c>NoCoverage</c>.
    /// </summary>
    [Fact]
    public async Task CancellationDuringTheWriteStaysACancellation()
    {
        using var cts = new CancellationTokenSource();
        using var destination = new CancelsOnFirstWriteSink(cts);

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => HtmlToDocxConverter.ConvertAsync(Html, destination, cts.Token));
    }
}
