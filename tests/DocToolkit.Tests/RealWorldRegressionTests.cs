using System.Text;

namespace DocToolkit.Tests;

/// <summary>
/// Conversion of real pages this project did not write.
///
/// <b>Every other HTML, DOCX, XLSX and PPTX fixture in this suite is built by the code under
/// test</b>, so the suite could only ever find defects in constructs it already knew how to produce.
/// That is not a hypothetical limitation. Measured 2026-08-17: <b>1,284 tests were green while HTML
/// to PDF succeeded on 58.6% of real `.gov` pages</b>, and every defect behind that gap - a rowspan
/// past the last row, obsolete <c>&lt;a name&gt;</c> anchors, spacer cells, image-only links - was
/// invisible to all of them.
///
/// <b>These files are kept whole, not reduced.</b> A reduced page is a page this project wrote,
/// which is the property being avoided. Each is the smallest real page in the corpus that exposed
/// one specific defect; <c>assets/realworld/README.txt</c> carries provenance and licence.
///
/// <b>The set is globbed, not listed</b>, so adding a file is the whole of adding a case - and the
/// count assertion below is what stops the glob silently matching nothing, which would turn this
/// entire class green and empty.
/// </summary>
public class RealWorldRegressionTests
{
    private static string Dir => Path.Join(AppContext.BaseDirectory, "assets", "realworld");

    public static TheoryData<string> EveryRealPage()
    {
        var data = new TheoryData<string>();
        foreach (var path in Directory.GetFiles(Dir, "*.html").OrderBy(p => p))
            data.Add(Path.GetFileName(path));
        return data;
    }

    /// <summary>
    /// Reads a page as the encoding it is actually in.
    /// </summary>
    /// <remarks>
    /// <b>Reading this corpus as UTF-8 unconditionally is how the first measurement of it went
    /// wrong.</b> Most of it is windows-1252, so invalid bytes became <c>U+FFFD</c> and the renderer
    /// then correctly refused to encode them - reported at the time as nine library failures that
    /// did not exist. Strict decoding distinguishes "is UTF-8" from "decodes to something", which
    /// the lenient default cannot: it substitutes silently.
    /// </remarks>
    private static string Read(string name)
    {
        var bytes = File.ReadAllBytes(Path.Join(Dir, name));
        try
        {
            return new UTF8Encoding(false, throwOnInvalidBytes: true).GetString(bytes);
        }
        catch (DecoderFallbackException)
        {
            return Encoding.Latin1.GetString(bytes);
        }
    }

    [Fact]
    public void TheCorpusIsActuallyPresent()
    {
        // Without this the glob could match nothing and every theory below would pass by having no
        // cases at all - the exact vacuous-green failure this project has had to correct elsewhere.
        Assert.True(Directory.Exists(Dir), $"{Dir} is missing - is the csproj still copying it?");
        Assert.True(Directory.GetFiles(Dir, "*.html").Length >= 4,
            "fewer real pages than expected: they are the only fixtures here the library did not write.");
    }

    [Theory]
    [MemberData(nameof(EveryRealPage))]
    public async Task EveryRealPageConvertsToDocx(string name)
    {
        var docx = await HtmlToDocxConverter.ConvertAsync(Read(name));

        Assert.NotEmpty(docx);
    }

    [Theory]
    [MemberData(nameof(EveryRealPage))]
    public async Task EveryRealPageConvertsToPdf(string name)
    {
        // The stricter of the two, and where all four of these defects surfaced.
        var pdf = await HtmlToPdfConverter.ConvertAsync(Read(name));

        Assert.True(PdfProbe.IsPdf(pdf), $"{name} no longer renders to PDF");
    }

    [Theory]
    [MemberData(nameof(EveryRealPage))]
    public async Task EveryRealPageKeepsItsText(string name)
    {
        // A conversion that succeeded by dropping the document's content would satisfy both theories
        // above. This does not check fidelity - only that something recognisable survived.
        var pdf = await HtmlToPdfConverter.ConvertAsync(Read(name));

        var text = PdfProbe.ExtractText(pdf);
        Assert.False(string.IsNullOrWhiteSpace(text), $"{name} rendered to a PDF with no text at all");
    }

    /// <summary>
    /// The pages are byte-for-byte what the crawl served.
    /// </summary>
    /// <remarks>
    /// <b>README.txt says these must not be edited, and this is what makes that enforceable rather
    /// than a request.</b> A hand-maintained list of hashes is normally the thing this repository
    /// avoids - but here immutability IS the property, exactly as with the approved API files, so
    /// pinning them is the point rather than a maintenance cost. A hash changing means somebody
    /// edited evidence.
    ///
    /// <b>Git nearly rewrote them on the way in.</b> <c>* text=auto</c> normalised two of the four
    /// from CRLF to LF - they came from .gov servers of the 2000s - so the committed blobs were not
    /// what the crawl served until <c>.gitattributes</c> marked this directory binary. That was
    /// invisible until the bytes were compared against the originals.
    /// </remarks>
    [Theory]
    [InlineData("image-only-link.html", "cd0acb6fe6e384e6d406cae07998e0af877598523c0763658d890ba00e3f1e0b")]
    [InlineData("old-style-name-anchor.html", "fe8d83e8e6b32091dc79badc8cb56e96857028147d99280d65d60f3f72a80bf6")]
    [InlineData("rowspan-past-last-row.html", "e6903ce52264e194a22466e75e8ffd2443be0e89e2c38c95104f72d34555ee1b")]
    [InlineData("spacer-cell.html", "d587ad6570ea15ef4339b272321701535822e694a63e3b13a6bd6928f4368187")]
    public void EachPageIsUnmodified(string name, string sha256)
    {
        var actual = Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(File.ReadAllBytes(Path.Join(Dir, name))));

        Assert.Equal(sha256, actual.ToLowerInvariant());
    }

    /// <summary>
    /// Each page still contains the construct it was kept for.
    /// </summary>
    /// <remarks>
    /// <b>A fixture that stops exposing its defect stops being a regression test</b>, silently, and
    /// would leave this class passing while guarding nothing. Editing these files is what
    /// README.txt forbids; this is what notices.
    /// </remarks>
    [Theory]
    [InlineData("rowspan-past-last-row.html", "rowspan")]
    [InlineData("old-style-name-anchor.html", "name=")]
    [InlineData("spacer-cell.html", "<td")]
    [InlineData("image-only-link.html", "<img")]
    public void EachPageStillCarriesTheConstructItWasKeptFor(string name, string marker)
    {
        Assert.Contains(marker, Read(name), StringComparison.OrdinalIgnoreCase);
    }
}
