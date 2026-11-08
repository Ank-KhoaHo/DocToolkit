using System.Buffers.Binary;
using Xunit.Abstractions;

namespace DocToolkit.Tests;

/// <summary>
/// Finds a font that covers the Arabic presentation forms the joining tests need, and records which
/// fonts on the host do.
/// </summary>
/// <remarks>
/// <b>Why the joining tests cannot use any font.</b> CI run 42513139477 (macOS): ten joining tests
/// failed because the shaped and unshaped first words started at the SAME X (192.88 both ways). A
/// render happened - nothing was refused - yet it drew the same glyphs for the letters and for their
/// presentation forms. The likely cause is that the macOS Arial lacks the forms; the hypothesis is
/// unproven, and <see cref="EveryCandidateFont_ReportsWhetherItCoversThePresentationForms"/> prints
/// the evidence into the log of each host.
/// </remarks>
public class ArabicFontTests(ITestOutputHelper output)
{
    /// <summary>Candidate fonts, in order of preference.</summary>
    internal static readonly string[] Candidates =
    [
        @"C:\Windows\Fonts\arial.ttf",
        @"C:\Windows\Fonts\tahoma.ttf",
        @"C:\Windows\Fonts\segoeui.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/Supplemental/Tahoma.ttf",
        "/Library/Fonts/Arial Unicode.ttf",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "/Library/Fonts/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
    ];

    /// <summary>Sample text the joining tests render.</summary>
    private static readonly string[] Samples = [PdfFontOptionsTests.ArabicLine, "سلام", "مدفوع", "بب"];

    /// <summary>
    /// Every presentation form the joining tests' input produces, plus the three the brief names.
    /// </summary>
    internal static IReadOnlySet<int> RequiredForms()
    {
        var forms = new SortedSet<int> { 0xFE91, 0xFE90, 0xFEFB };
        foreach (var sample in Samples)
            foreach (var piece in ArabicPreShaper.Shape(sample))
                foreach (var c in piece)
                    if (c is >= '\uFB50' and <= '\uFEFC')
                        forms.Add(c);
        return forms;
    }

    /// <summary>The required forms the font's cmap does not map; empty means it covers them all.</summary>
    internal static IReadOnlyList<int> MissingForms(byte[] font)
    {
        var mapped = FontCmap.Covers(font);
        return RequiredForms().Where(cp => !mapped(cp)).ToList();
    }

    /// <summary>
    /// The first candidate font whose cmap covers every presentation form the joining tests use; null
    /// on a host with none, which is a FAILURE instead when <c>DOCTOOLKIT_REQUIRE_FONT=1</c>.
    /// </summary>
    internal static byte[]? FindArabicFont()
    {
        var tried = new List<string>();
        foreach (var path in Candidates.Where(File.Exists))
        {
            var bytes = File.ReadAllBytes(path);
            var missing = MissingForms(bytes);
            if (missing.Count == 0)
                return bytes;
            tried.Add($"{path} lacks {string.Join(" ", missing.Select(cp => $"U+{cp:X4}"))}");
        }

        if (Environment.GetEnvironmentVariable("DOCTOOLKIT_REQUIRE_FONT") == "1")
            Assert.Fail(
                "DOCTOOLKIT_REQUIRE_FONT=1, but no font on this host covers the Arabic presentation forms. "
                + (tried.Count == 0
                    ? "None of the candidates exists: " + string.Join(", ", Candidates)
                    : "Tried: " + string.Join("; ", tried)));
        return null;
    }

    /// <summary>
    /// Records, for the log of whichever host runs this, which candidates exist and whether each covers
    /// the forms. Asserts nothing about any particular font - the point is the evidence.
    /// </summary>
    [Fact]
    public void EveryCandidateFont_ReportsWhetherItCoversThePresentationForms()
    {
        var required = RequiredForms();
        output.WriteLine($"Required forms ({required.Count}): {string.Join(" ", required.Select(cp => $"U+{cp:X4}"))}");
        foreach (var path in Candidates)
        {
            if (!File.Exists(path))
            {
                output.WriteLine($"absent   {path}");
                continue;
            }
            var missing = MissingForms(File.ReadAllBytes(path));
            output.WriteLine(missing.Count == 0
                ? $"covers   {path}"
                : $"lacks {missing.Count}/{required.Count}   {path}: {string.Join(" ", missing.Select(cp => $"U+{cp:X4}"))}");
        }

        Assert.NotEmpty(required);
    }

    [Fact]
    public void RequiredForms_AreTheFormsTheShaperProduces()
    {
        var forms = RequiredForms();

        Assert.Contains(0xFE91, forms);
        Assert.Contains(0xFEFB, forms);
        Assert.All(forms, cp => Assert.InRange(cp, 0xFB50, 0xFEFC));
    }

    // ---- the cmap reader itself, on a font built in memory ---------------------------------------

    [Fact]
    public void Cmap_Format4_ReportsMappedAndUnmappedCodePoints()
    {
        var covers = FontCmap.Covers(FontCmap.BuildFormat4(0xFE90, 0xFE91));

        Assert.True(covers(0xFE90));
        Assert.True(covers(0xFE91));
        Assert.False(covers(0xFE92));
        Assert.False(covers(0x0041));
    }

    [Fact]
    public void Cmap_Format12_ReportsMappedAndUnmappedCodePoints()
    {
        var covers = FontCmap.Covers(FontCmap.BuildFormat12(0xFEFB, 0xFEFB));

        Assert.True(covers(0xFEFB));
        Assert.False(covers(0xFEFA));
        Assert.False(covers(0xFEFC));
    }

    [Fact]
    public void Cmap_OfSomethingThatIsNotAFont_CoversNothing()
    {
        Assert.False(FontCmap.Covers([1, 2, 3, 4])(0x41));
        Assert.False(FontCmap.Covers([])(0x41));
    }

    [Fact]
    public void Cmap_Format4_HonoursIdRangeOffsetAndZeroGlyph()
    {
        // Segment 0xFE90..0xFE91 reached through a glyph-id array: first entry glyph 5, second glyph 0 (unmapped).
        var covers = FontCmap.Covers(FontCmap.BuildFormat4WithGlyphArray(0xFE90, 5, 0));

        Assert.True(covers(0xFE90));
        Assert.False(covers(0xFE91));
    }
}

/// <summary>A minimal TrueType <c>cmap</c> reader (formats 4 and 12), enough to ask "is this code point mapped".</summary>
internal static class FontCmap
{
    private static ushort U16(byte[] b, int o) => BinaryPrimitives.ReadUInt16BigEndian(b.AsSpan(o, 2));
    private static short S16(byte[] b, int o) => BinaryPrimitives.ReadInt16BigEndian(b.AsSpan(o, 2));
    private static uint U32(byte[] b, int o) => BinaryPrimitives.ReadUInt32BigEndian(b.AsSpan(o, 4));

    /// <summary>A predicate over code points; false for everything when the data is not a readable font.</summary>
    internal static Func<int, bool> Covers(byte[] font)
    {
        try
        {
            var tables = Subtables(font);
            return cp => tables.Any(t => t(cp));
        }
        catch (Exception ex) when (ex is IndexOutOfRangeException or ArgumentException)
        {
            return _ => false;
        }
    }

    private static List<Func<int, bool>> Subtables(byte[] font)
    {
        var result = new List<Func<int, bool>>();
        if (font.Length < 12)
            return result;
        var start = 0;
        if (U32(font, 0) == 0x74746366) // 'ttcf': a collection; read its first font
            start = (int)U32(font, 12);
        var tableCount = U16(font, start + 4);
        for (var i = 0; i < tableCount; i++)
        {
            var rec = start + 12 + 16 * i;
            if (U32(font, rec) != 0x636D6170) // 'cmap'
                continue;
            var cmap = (int)U32(font, rec + 8);
            var count = U16(font, cmap + 2);
            for (var e = 0; e < count; e++)
            {
                var table = cmap + (int)U32(font, cmap + 4 + 8 * e + 4);
                var format = U16(font, table);
                if (format == 4)
                    result.Add(Format4(font, table));
                else if (format == 12)
                    result.Add(Format12(font, table));
            }
        }
        return result;
    }

    private static Func<int, bool> Format4(byte[] f, int t)
    {
        var segCount = U16(f, t + 6) / 2;
        var endCodes = t + 14;
        var startCodes = endCodes + 2 * segCount + 2;
        var deltas = startCodes + 2 * segCount;
        var rangeOffsets = deltas + 2 * segCount;
        return cp =>
        {
            if (cp > 0xFFFF)
                return false;
            for (var s = 0; s < segCount; s++)
            {
                if (cp > U16(f, endCodes + 2 * s))
                    continue;
                if (cp < U16(f, startCodes + 2 * s))
                    return false;
                var ro = U16(f, rangeOffsets + 2 * s);
                if (ro == 0)
                    return ((cp + S16(f, deltas + 2 * s)) & 0xFFFF) != 0;
                var glyphAt = rangeOffsets + 2 * s + ro + 2 * (cp - U16(f, startCodes + 2 * s));
                var glyph = U16(f, glyphAt);
                return glyph != 0 && ((glyph + S16(f, deltas + 2 * s)) & 0xFFFF) != 0;
            }
            return false;
        };
    }

    private static Func<int, bool> Format12(byte[] f, int t)
    {
        var groups = (int)U32(f, t + 12);
        return cp =>
        {
            for (var g = 0; g < groups; g++)
            {
                var o = t + 16 + 12 * g;
                if (cp >= U32(f, o) && cp <= U32(f, o + 4))
                    return U32(f, o + 8) + (uint)cp - U32(f, o) != 0;
            }
            return false;
        };
    }

    // ---- builders for the unit tests -----------------------------------------------------------

    private static byte[] Wrap(byte[] subtable, ushort platform, ushort encoding)
    {
        var cmap = new List<byte>();
        P16(cmap, 0); P16(cmap, 1); P16(cmap, platform); P16(cmap, encoding); P32(cmap, 12);
        cmap.AddRange(subtable);

        var font = new List<byte>();
        P32(font, 0x00010000); P16(font, 1); P16(font, 0); P16(font, 0); P16(font, 0);
        P32(font, 0x636D6170); P32(font, 0); P32(font, 28); P32(font, (uint)cmap.Count);
        font.AddRange(cmap);
        return font.ToArray();
    }

    private static void P16(List<byte> l, int v) { l.Add((byte)(v >> 8)); l.Add((byte)v); }
    private static void P32(List<byte> l, uint v) { P16(l, (int)(v >> 16)); P16(l, (int)(v & 0xFFFF)); }

    /// <summary>A font mapping first..last to glyphs 1.. through a delta (idRangeOffset 0).</summary>
    internal static byte[] BuildFormat4(int first, int last)
    {
        var s = new List<byte>();
        P16(s, 4); P16(s, 16 + 8 * 2); P16(s, 0); P16(s, 4); P16(s, 4); P16(s, 1); P16(s, 0);
        P16(s, last); P16(s, 0xFFFF); P16(s, 0);
        P16(s, first); P16(s, 0xFFFF);
        P16(s, 1 - first & 0xFFFF); P16(s, 1);
        P16(s, 0); P16(s, 0);
        return Wrap(s.ToArray(), 3, 1);
    }

    /// <summary>A font mapping first.. to the given glyph ids through a glyph-id array.</summary>
    internal static byte[] BuildFormat4WithGlyphArray(int first, params int[] glyphs)
    {
        var s = new List<byte>();
        P16(s, 4); P16(s, 16 + 8 * 2 + 2 * glyphs.Length); P16(s, 0); P16(s, 4); P16(s, 4); P16(s, 1); P16(s, 0);
        P16(s, first + glyphs.Length - 1); P16(s, 0xFFFF); P16(s, 0);
        P16(s, first); P16(s, 0xFFFF);
        P16(s, 0); P16(s, 1);
        P16(s, 4); P16(s, 0);   // idRangeOffset of segment 0 points past the two offsets to the array
        foreach (var g in glyphs)
            P16(s, g);
        return Wrap(s.ToArray(), 3, 1);
    }

    /// <summary>A font with one format 12 group.</summary>
    internal static byte[] BuildFormat12(int first, int last)
    {
        var s = new List<byte>();
        P16(s, 12); P16(s, 0); P32(s, 28); P32(s, 0); P32(s, 1);
        P32(s, (uint)first); P32(s, (uint)last); P32(s, 1);
        return Wrap(s.ToArray(), 3, 10);
    }
}
