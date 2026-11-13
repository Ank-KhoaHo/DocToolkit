using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace DocToolkit;

// The forms table and the ArabicForms struct below are copied from OfficeIMO.Drawing/Shaper.cs,
// OfficeIMO version 3.5.3, repository github.com/EvotecIT/OfficeIMO.
// MIT License, Copyright (c) 2022 Evotec. See THIRD-PARTY-NOTICES.txt.

/// <summary>
/// Swaps Arabic-script letters (Arabic, Persian, Urdu and their relatives) for their contextual
/// presentation forms, in logical order, before OfficeIMO renders. OfficeIMO 3.5 lays the letters
/// out in the right order but does not join them; once they are already in their joined forms it
/// has nothing left to join.
/// </summary>
internal static class ArabicPreShaper
{
    private static readonly IReadOnlyDictionary<char, ArabicForms> Forms = CreateForms();

    /// <summary>
    /// Shapes a whole container's text. Returns one string per input char, so the result splits back
    /// into the same runs: usually one char, and empty for the alef of a lam-alef ligature, which
    /// takes the lam's slot. The output stays in logical order; it is never reversed.
    /// </summary>
    internal static string[] Shape(string logical)
    {
        var result = new string[logical.Length];
        for (var i = 0; i < logical.Length; i++)
        {
            var c = logical[i];
            if (!Forms.TryGetValue(c, out var forms)) { result[i] = c.ToString(); continue; }
            var joinsPrevious = forms.CanJoinPrevious && TryFind(logical, i, -1, out var prev) && prev.CanJoinNext;
            if (c == '\u0644' && NextIndex(logical, i) is int a && LamAlef(logical[a]) is (char iso, char fin))
            {
                result[i] = (joinsPrevious ? fin : iso).ToString();
                for (var k = i + 1; k <= a; k++) result[k] = k == a ? string.Empty : logical[k].ToString();
                i = a;
                continue;
            }
            var joinsNext = forms.CanJoinNext && TryFind(logical, i, +1, out var next) && next.CanJoinPrevious;
            result[i] = forms.Resolve(joinsPrevious, joinsNext).ToString();
        }
        return result;
    }

    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>
    /// <see cref="Apply(byte[])"/> for a package already sitting in an expandable stream: the
    /// stream is rewritten in place, and only when the package held Arabic. Position is left at 0.
    /// </summary>
    /// <remarks>
    /// For the stream paths, which read a package into a scratch <see cref="MemoryStream"/> and
    /// render it from there rather than through the <c>byte[]</c> overloads.
    /// </remarks>
    internal static void ApplyInPlace(MemoryStream package)
    {
        var bytes = package.ToArray();
        var shaped = Apply(bytes);
        if (!ReferenceEquals(shaped, bytes))
        {
            package.SetLength(0);
            package.Write(shaped, 0, shaped.Length);
        }

        package.Position = 0;
    }

    /// <summary>
    /// Rewrites the Arabic-script text of a Word, Excel or PowerPoint package into presentation
    /// forms, and returns <paramref name="package"/> itself when no text changes.
    /// </summary>
    /// <remarks>
    /// <b>The unit of shaping is the container, not the run.</b> A word is often split over runs
    /// (a font change, a spell-check mark), and a letter's form depends on its neighbours across
    /// the split. So a paragraph or string item is shaped as one string and each text node then
    /// gets back exactly its own slice; a lam-alef ligature may shorten the node holding the alef.
    /// An Excel shared string is one <c>si</c>, so it is shaped once however many cells use it.
    /// </remarks>
    internal static byte[] Apply(byte[] package)
    {
        // A zip starts with "PK"; anything else is not a package this can help.
        if (package.Length < 4 || package[0] != (byte)'P' || package[1] != (byte)'K') return package;
        return DocxPartRewrite.Apply(package, IsTargetPart, ShapeDocument, requireWordMain: false);
    }

    /// <summary>
    /// Whether the text of a workbook's shared strings or sheets (formula text included) holds an
    /// Arabic-script letter. Reads the package; never modifies it.
    /// </summary>
    internal static bool HasArabicLetter(byte[] package)
    {
        if (package.Length < 4 || package[0] != (byte)'P' || package[1] != (byte)'K') return false;
        var found = false;
        DocxPartRewrite.Apply(
            package, IsTargetPart, doc => { found |= ContainsArabicLetter(doc.Root!.Value); return false; }, requireWordMain: false);
        return found;
    }

    private static bool IsTargetPart(string name) =>
        name is "word/document.xml" or "word/footnotes.xml" or "word/endnotes.xml" or "xl/sharedStrings.xml"
        || IsNumbered(name, "word/header") || IsNumbered(name, "word/footer")
        || IsNumbered(name, "xl/worksheets/sheet") || IsNumbered(name, "ppt/slides/slide");

    // "<prefix>*.xml" in that folder only: header1.xml yes, _rels/header1.xml.rels and slideLayouts no.
    private static bool IsNumbered(string name, string prefix) =>
        name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) && name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
        && name.IndexOf('/', prefix.Length) < 0;

    private static bool ShapeDocument(XDocument doc)
    {
        var changed = false;
        var logical = new StringBuilder();
        foreach (var container in doc.Descendants().Where(IsContainer).ToList())
        {
            // Text nodes in document order, with the logical string they are shaped as. A line
            // break, tab or inline object between two text nodes becomes one space in that string,
            // which does not join, so letters never join across it. Separators belong to no node.
            var nodes = new List<(XElement Node, string Original, int Start)>();
            logical.Clear();
            var separate = false;
            foreach (var e in container.Descendants())
            {
                if (IsTextNode(e, container))
                {
                    if (separate && nodes.Count > 0) logical.Append(' ');
                    separate = false;
                    nodes.Add((e, e.Value, logical.Length));
                    logical.Append(e.Value);
                }
                else if (IsNonTextContent(e) && e.Ancestors().First(IsContainer) == container)
                {
                    separate = true;
                }
            }

            var text = logical.ToString();
            if (!ContainsArabicLetter(text)) continue;

            var shaped = Shape(text);
            foreach (var (node, original, start) in nodes)
            {
                var replacement = string.Concat(shaped.Skip(start).Take(original.Length));
                if (replacement == original) continue;
                node.Value = replacement;
                changed = true;
            }
        }
        return changed;
    }

    private static bool ContainsArabicLetter(string s)
    {
        foreach (var c in s) if (c is >= 'ؠ' and <= 'ۿ') return true;
        return false;
    }

    // A formula cell's text result (t="str") is one string in its v, so that is a container too.
    private static bool IsContainer(XElement e) =>
        (e.Name.LocalName == "p" && (e.Name.Namespace == W || e.Name.Namespace == A))
        || (e.Name.Namespace == S && e.Name.LocalName is "si" or "is")
        || (e.Name == S + "c" && (string?)e.Attribute("t") == "str");

    // Content that sits between text nodes and breaks a word visually: line and page breaks, tabs,
    // hyphens, fields and inline objects. Their text, if any, is not this container's text.
    private static bool IsNonTextContent(XElement e) =>
        (e.Name.Namespace == W && e.Name.LocalName is "br" or "cr" or "tab" or "ptab" or "noBreakHyphen"
            or "softHyphen" or "drawing" or "object" or "pict" or "fldChar" or "sym")
        || (e.Name.Namespace == A && e.Name.LocalName == "br");

    // A text node belongs to its NEAREST container: a text box's paragraph inside a run of another
    // paragraph is shaped on its own. In si/is only the item's own text counts (directly, or in an
    // r run), so phonetic text (rPh) is skipped.
    private static bool IsTextNode(XElement e, XElement container)
    {
        if (container.Name == S + "c") return e.Name == S + "v" && e.Parent == container;
        if (e.Name.LocalName != "t" || e.Name.Namespace != container.Name.Namespace) return false;
        if (container.Name.Namespace == S)
            return e.Parent == container || (e.Parent!.Name == S + "r" && e.Parent.Parent == container);
        return e.Ancestors().First(IsContainer) == container;
    }

    // Isolated and final ligature for lam followed by each alef (the final form is the isolated + 1).
    private static (char, char)? LamAlef(char alef) => alef switch
    {
        '\u0622' => ('\uFEF5', '\uFEF6'),
        '\u0623' => ('\uFEF7', '\uFEF8'),
        '\u0625' => ('\uFEF9', '\uFEFA'),
        '\u0627' => ('\uFEFB', '\uFEFC'),
        _ => null,
    };

    private static int? NextIndex(string s, int i)
    {
        for (var k = i + 1; k < s.Length; k++)
        {
            if (s[k] == '\u200C') return null;
            if (IsTransparent(s[k])) continue;
            return k;
        }
        return null;
    }

    private static bool TryFind(string s, int index, int step, out ArabicForms forms)
    {
        for (var k = index + step; k >= 0 && k < s.Length; k += step)
        {
            var c = s[k];
            if (c == '\u200C') break;
            if (IsTransparent(c)) continue;
            if (c == '\u0640') { forms = ArabicForms.JoiningCarrier; return true; }
            if (Forms.TryGetValue(c, out forms)) return true;
            break;
        }
        forms = default;
        return false;
    }

    // Harakat (U+064B-U+065F, U+0670) are non-spacing marks; ZWJ joins through. ZWNJ is a Format
    // char too, but TryFind stops at it before this test.
    private static bool IsTransparent(char value) =>
        value == '\u200D' || CharUnicodeInfo.GetUnicodeCategory(value) is UnicodeCategory.NonSpacingMark or UnicodeCategory.EnclosingMark or UnicodeCategory.Format;

    private static IReadOnlyDictionary<char, ArabicForms> CreateForms() => new Dictionary<char, ArabicForms>
    {
        ['\u0621'] = new('\uFE80'),
        ['\u0622'] = new('\uFE81', '\uFE82'),
        ['\u0623'] = new('\uFE83', '\uFE84'),
        ['\u0624'] = new('\uFE85', '\uFE86'),
        ['\u0625'] = new('\uFE87', '\uFE88'),
        ['\u0626'] = new('\uFE89', '\uFE8A', '\uFE8B', '\uFE8C'),
        ['\u0627'] = new('\uFE8D', '\uFE8E'),
        ['\u0628'] = new('\uFE8F', '\uFE90', '\uFE91', '\uFE92'),
        ['\u0629'] = new('\uFE93', '\uFE94'),
        ['\u062A'] = new('\uFE95', '\uFE96', '\uFE97', '\uFE98'),
        ['\u062B'] = new('\uFE99', '\uFE9A', '\uFE9B', '\uFE9C'),
        ['\u062C'] = new('\uFE9D', '\uFE9E', '\uFE9F', '\uFEA0'),
        ['\u062D'] = new('\uFEA1', '\uFEA2', '\uFEA3', '\uFEA4'),
        ['\u062E'] = new('\uFEA5', '\uFEA6', '\uFEA7', '\uFEA8'),
        ['\u062F'] = new('\uFEA9', '\uFEAA'),
        ['\u0630'] = new('\uFEAB', '\uFEAC'),
        ['\u0631'] = new('\uFEAD', '\uFEAE'),
        ['\u0632'] = new('\uFEAF', '\uFEB0'),
        ['\u0633'] = new('\uFEB1', '\uFEB2', '\uFEB3', '\uFEB4'),
        ['\u0634'] = new('\uFEB5', '\uFEB6', '\uFEB7', '\uFEB8'),
        ['\u0635'] = new('\uFEB9', '\uFEBA', '\uFEBB', '\uFEBC'),
        ['\u0636'] = new('\uFEBD', '\uFEBE', '\uFEBF', '\uFEC0'),
        ['\u0637'] = new('\uFEC1', '\uFEC2', '\uFEC3', '\uFEC4'),
        ['\u0638'] = new('\uFEC5', '\uFEC6', '\uFEC7', '\uFEC8'),
        ['\u0639'] = new('\uFEC9', '\uFECA', '\uFECB', '\uFECC'),
        ['\u063A'] = new('\uFECD', '\uFECE', '\uFECF', '\uFED0'),
        ['\u0641'] = new('\uFED1', '\uFED2', '\uFED3', '\uFED4'),
        ['\u0642'] = new('\uFED5', '\uFED6', '\uFED7', '\uFED8'),
        ['\u0643'] = new('\uFED9', '\uFEDA', '\uFEDB', '\uFEDC'),
        ['\u0644'] = new('\uFEDD', '\uFEDE', '\uFEDF', '\uFEE0'),
        ['\u0645'] = new('\uFEE1', '\uFEE2', '\uFEE3', '\uFEE4'),
        ['\u0646'] = new('\uFEE5', '\uFEE6', '\uFEE7', '\uFEE8'),
        ['\u0647'] = new('\uFEE9', '\uFEEA', '\uFEEB', '\uFEEC'),
        ['\u0648'] = new('\uFEED', '\uFEEE'),
        ['\u0649'] = new('\uFEEF', '\uFEF0'),
        ['\u064A'] = new('\uFEF1', '\uFEF2', '\uFEF3', '\uFEF4'),
        // Extended letters for Persian, Urdu, Sindhi, Pashto, Kurdish and Uyghur (Presentation Forms-A).
        ['\u0671'] = new('\uFB50', '\uFB51'),
        ['\u0679'] = new('\uFB66', '\uFB67', '\uFB68', '\uFB69'),
        ['\u067A'] = new('\uFB5E', '\uFB5F', '\uFB60', '\uFB61'),
        ['\u067B'] = new('\uFB52', '\uFB53', '\uFB54', '\uFB55'),
        ['\u067E'] = new('\uFB56', '\uFB57', '\uFB58', '\uFB59'),
        ['\u067F'] = new('\uFB62', '\uFB63', '\uFB64', '\uFB65'),
        ['\u0680'] = new('\uFB5A', '\uFB5B', '\uFB5C', '\uFB5D'),
        ['\u0683'] = new('\uFB76', '\uFB77', '\uFB78', '\uFB79'),
        ['\u0684'] = new('\uFB72', '\uFB73', '\uFB74', '\uFB75'),
        ['\u0686'] = new('\uFB7A', '\uFB7B', '\uFB7C', '\uFB7D'),
        ['\u0687'] = new('\uFB7E', '\uFB7F', '\uFB80', '\uFB81'),
        ['\u0688'] = new('\uFB88', '\uFB89'),
        ['\u068C'] = new('\uFB84', '\uFB85'),
        ['\u068D'] = new('\uFB82', '\uFB83'),
        ['\u068E'] = new('\uFB86', '\uFB87'),
        ['\u0691'] = new('\uFB8C', '\uFB8D'),
        ['\u0698'] = new('\uFB8A', '\uFB8B'),
        ['\u06A4'] = new('\uFB6A', '\uFB6B', '\uFB6C', '\uFB6D'),
        ['\u06A6'] = new('\uFB6E', '\uFB6F', '\uFB70', '\uFB71'),
        ['\u06A9'] = new('\uFB8E', '\uFB8F', '\uFB90', '\uFB91'),
        ['\u06AD'] = new('\uFBD3', '\uFBD4', '\uFBD5', '\uFBD6'),
        ['\u06AF'] = new('\uFB92', '\uFB93', '\uFB94', '\uFB95'),
        ['\u06B1'] = new('\uFB9A', '\uFB9B', '\uFB9C', '\uFB9D'),
        ['\u06B3'] = new('\uFB96', '\uFB97', '\uFB98', '\uFB99'),
        ['\u06BA'] = new('\uFB9E', '\uFB9F'),
        ['\u06BB'] = new('\uFBA0', '\uFBA1', '\uFBA2', '\uFBA3'),
        ['\u06BE'] = new('\uFBAA', '\uFBAB', '\uFBAC', '\uFBAD'),
        ['\u06C0'] = new('\uFBA4', '\uFBA5'),
        ['\u06C1'] = new('\uFBA6', '\uFBA7', '\uFBA8', '\uFBA9'),
        ['\u06C5'] = new('\uFBE0', '\uFBE1'),
        ['\u06C6'] = new('\uFBD9', '\uFBDA'),
        ['\u06C7'] = new('\uFBD7', '\uFBD8'),
        ['\u06C8'] = new('\uFBDB', '\uFBDC'),
        ['\u06C9'] = new('\uFBE2', '\uFBE3'),
        ['\u06CB'] = new('\uFBDE', '\uFBDF'),
        ['\u06CC'] = new('\uFBFC', '\uFBFD', '\uFBFE', '\uFBFF'),
        ['\u06D0'] = new('\uFBE4', '\uFBE5', '\uFBE6', '\uFBE7'),
        ['\u06D2'] = new('\uFBAE', '\uFBAF'),
        ['\u06D3'] = new('\uFBB0', '\uFBB1')
    };

    private readonly struct ArabicForms
    {
        internal static ArabicForms JoiningCarrier { get; } = new('\0', '\u0640', '\u0640', '\u0640');

        internal ArabicForms(char isolated, char final = '\0', char initial = '\0', char medial = '\0')
        {
            Isolated = isolated;
            Final = final;
            Initial = initial;
            Medial = medial;
        }

        internal char Isolated { get; }
        internal char Final { get; }
        internal char Initial { get; }
        internal char Medial { get; }
        internal bool CanJoinPrevious => Final != '\0' || Medial != '\0';
        internal bool CanJoinNext => Initial != '\0' || Medial != '\0';

        internal char Resolve(bool joinsPrevious, bool joinsNext)
        {
            if (joinsPrevious && joinsNext && Medial != '\0') return Medial;
            if (joinsPrevious && Final != '\0') return Final;
            if (joinsNext && Initial != '\0') return Initial;
            return Isolated;
        }
    }
}
