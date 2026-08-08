using System.Text;
using Xunit;
using Xunit.Abstractions;

namespace DocToolkit.Tests;

/// <summary>
/// Direct tests for <see cref="RunTextSplicer"/>, the shared placeholder-substitution core behind
/// <see cref="DocxEditor"/>'s and <see cref="PresentationEditor"/>'s ReplaceText.
///
/// <para>It had none until now — it was exercised only end to end, through documents. That is the
/// wrong shape of coverage for this particular code, for the same reason `TableRowFinderTests`
/// exists: when it regresses, nothing throws. The document is still schema-valid, it just quietly
/// lost a run's formatting or emptied a hyperlink. An end-to-end test that reads text back sees
/// the right text and passes.</para>
///
/// <para>The interesting cases are all about run <b>boundaries</b>, which is why the centrepiece
/// here is a property test rather than a list of examples: a placeholder can straddle any number
/// of runs at any offset, and hand-picked splits systematically miss the arrangement nobody
/// thought of.</para>
///
/// <para><b>One expression here is deliberately not covered, because it cannot be.</b> In Apply,
/// <c>pos = Math.Min(match.End, nodeEnd)</c> is equivalent to <c>pos = match.End</c> for every
/// input: <c>pos</c> is re-initialised to the node's start on each iteration of the outer loop, and
/// the inner loop's condition is <c>pos &lt; nodeEnd</c>, so overshooting the node end and landing
/// exactly on it both exit immediately, and <c>pos</c> is never read afterwards. Replacing the
/// <c>Math.Min</c> leaves all of these tests green — not a coverage gap but a redundant guard. Do
/// not spend an afternoon trying to write the test that kills it; write the simplification
/// instead, in a change that is allowed to touch production logic.</para>
/// </summary>
public class RunTextSplicerTests
{
    private readonly ITestOutputHelper _output;

    public RunTextSplicerTests(ITestOutputHelper output) => _output = output;

    /// <summary>One text run. Records whether it was written to, which is the whole point.</summary>
    private sealed class Node(string text)
    {
        public string Text { get; private set; } = text;
        public int Writes { get; private set; }

        public void Set(string value)
        {
            Text = value;
            Writes++;
        }
    }

    private static bool Apply(IReadOnlyList<Node> nodes, IReadOnlyDictionary<string, string> replacements)
        => RunTextSplicer.Apply(nodes, n => n.Text, (n, v) => n.Set(v), replacements);

    private static List<Node> Split(params string[] runs) => runs.Select(r => new Node(r)).ToList();

    private static string Merged(IEnumerable<Node> nodes) => string.Concat(nodes.Select(n => n.Text));

    // =============================================================================================
    // The property: how the text is split across runs must not change the result.
    // =============================================================================================

    /// <summary>
    /// The one invariant that matters. A placeholder split across runs must substitute exactly as
    /// it would in a single run — that is the entire reason this class exists over a per-run
    /// string.Replace.
    ///
    /// <para><b>The oracle is the splicer itself, on the single-run case.</b> That is deliberate:
    /// reimplementing the longest-match-wins scan here would only prove the test agrees with the
    /// test. The single-run case exercises none of the boundary arithmetic — there is one node, so
    /// every match starts and ends inside it — while the split cases exercise nothing but. The
    /// simple case is anchored by the explicit examples further down.</para>
    ///
    /// <para>The seed is fixed. A random seed would make this the fourth wall-clock-or-random flake
    /// in this repository's history, and a test that fails once a fortnight on a case nobody can
    /// reproduce is worse than no test.</para>
    /// </summary>
    [Fact]
    public void HowTheTextIsSplitAcrossRunsNeverChangesTheResult()
    {
        var random = new Random(20260808);
        var replacements = new Dictionary<string, string>
        {
            ["{{a}}"] = "ALPHA",
            ["{{b}}"] = "B",
            ["{{ab}}"] = "",          // empty replacement: the match is deleted
            ["x"] = "{{a}}",          // value contains a key, to catch rescanning
        };

        // Fragments chosen so placeholders form, half-form and abut each other by accident.
        string[] fragments = ["{{a}}", "{{b}}", "{{ab}}", "{{", "}}", "a", "b", "x", "", " ", "ab"];

        var cases = 0;
        for (var iteration = 0; iteration < 2000; iteration++)
        {
            var text = new StringBuilder();
            for (var f = random.Next(0, 6); f > 0; f--)
                text.Append(fragments[random.Next(fragments.Length)]);

            var merged = text.ToString();
            if (merged.Length == 0) continue;

            // The oracle: one run holding the whole text.
            var whole = Split(merged);
            Apply(whole, replacements);
            var expected = Merged(whole);

            // The same text, cut at random offsets - including empty runs, which are legal.
            var cuts = new SortedSet<int> { 0, merged.Length };
            for (var c = random.Next(1, 5); c > 0; c--) cuts.Add(random.Next(0, merged.Length + 1));

            var ordered = cuts.ToArray();
            var pieces = new List<string>();
            for (var i = 1; i < ordered.Length; i++)
                pieces.Add(merged[ordered[i - 1]..ordered[i]]);

            var split = Split([.. pieces]);
            Apply(split, replacements);

            Assert.Equal(expected, Merged(split));
            cases++;
        }

        _output.WriteLine($"{cases} split arrangements agreed with the single-run result");
        Assert.True(cases > 1500, $"only {cases} cases were generated; the generator is not producing text");
    }

    /// <summary>
    /// The formatting guarantee, stated as a property: a run the match does not overlap must not be
    /// written at all. Writing it back unchanged would still pass a text round-trip while, in a real
    /// document, replacing the run's element and discarding its formatting.
    /// </summary>
    [Fact]
    public void ARunNoMatchOverlapsIsNeverWritten()
    {
        var random = new Random(20260809);
        var replacements = new Dictionary<string, string> { ["{{x}}"] = "VALUE" };

        for (var iteration = 0; iteration < 500; iteration++)
        {
            // A placeholder somewhere in the middle, with untouched runs either side.
            var before = new string('a', random.Next(1, 6));
            var after = new string('z', random.Next(1, 6));
            var nodes = Split(before, "{{x", "}}", after);

            Assert.True(Apply(nodes, replacements));

            Assert.Equal(0, nodes[0].Writes);
            Assert.Equal(0, nodes[3].Writes);
            Assert.Equal(before, nodes[0].Text);
            Assert.Equal(after, nodes[3].Text);
        }
    }

    // =============================================================================================
    // The documented behaviours, as examples. These anchor the oracle the property test relies on.
    // =============================================================================================

    [Fact]
    public void SubstitutesAPlaceholderContainedInASingleRun()
    {
        var nodes = Split("Hello {{who}}!");

        Assert.True(Apply(nodes, new Dictionary<string, string> { ["{{who}}"] = "World" }));

        Assert.Equal("Hello World!", Merged(nodes));
    }

    [Fact]
    public void SubstitutesAPlaceholderSplitAcrossThreeRuns()
    {
        var nodes = Split("Hello {{w", "h", "o}}!");

        Assert.True(Apply(nodes, new Dictionary<string, string> { ["{{who}}"] = "World" }));

        Assert.Equal("Hello World!", Merged(nodes));
    }

    /// <summary>
    /// Where the value lands is not arbitrary: it goes into the run owning the match's <b>start</b>,
    /// so it inherits that run's formatting, and the runs the match merely spans lose their share of
    /// the matched characters. Pinned because it is a deliberate choice a refactor could silently
    /// reverse — putting the value on the last run instead would still round-trip the text.
    /// </summary>
    [Fact]
    public void TheValueLandsInTheRunThatOwnsTheStartOfTheMatch()
    {
        var nodes = Split("a{{w", "ho", "}}z");

        Assert.True(Apply(nodes, new Dictionary<string, string> { ["{{who}}"] = "World" }));

        Assert.Equal("aWorld", nodes[0].Text);
        Assert.Equal("", nodes[1].Text);
        Assert.Equal("z", nodes[2].Text);
    }

    [Fact]
    public void AtOneOffsetTheLongestKeyWins()
    {
        var nodes = Split("{{ab}}");

        Assert.True(Apply(nodes, new Dictionary<string, string>
        {
            ["{{a"] = "SHORT",
            ["{{ab}}"] = "LONG",
        }));

        Assert.Equal("LONG", Merged(nodes));
    }

    /// <summary>
    /// The outcome must not depend on dictionary enumeration order — otherwise the same template
    /// and the same replacements could produce different documents on different runs.
    /// </summary>
    [Fact]
    public void TheOutcomeDoesNotDependOnDictionaryOrder()
    {
        var oneWay = Split("{{ab}}");
        Apply(oneWay, new Dictionary<string, string> { ["{{a"] = "SHORT", ["{{ab}}"] = "LONG" });

        var otherWay = Split("{{ab}}");
        Apply(otherWay, new Dictionary<string, string> { ["{{ab}}"] = "LONG", ["{{a"] = "SHORT" });

        Assert.Equal(Merged(oneWay), Merged(otherWay));
    }

    /// <summary>
    /// A single left-to-right pass: a value that happens to contain a placeholder is not rescanned.
    /// Without this, a replacement whose value mentions another key would expand recursively, and a
    /// value containing its own key would not terminate.
    /// </summary>
    [Fact]
    public void ASubstitutedValueIsNotRescanned()
    {
        var nodes = Split("{{outer}}");

        Assert.True(Apply(nodes, new Dictionary<string, string>
        {
            ["{{outer}}"] = "{{inner}}",
            ["{{inner}}"] = "SHOULD NOT APPEAR",
        }));

        Assert.Equal("{{inner}}", Merged(nodes));
    }

    /// <summary>
    /// Scanning resumes <b>after</b> a match, never inside it, so matches can never overlap.
    ///
    /// <para>Added because mutation testing found the gap: changing the scan to advance one
    /// character after a match instead of past it left every other test here green.</para>
    ///
    /// <para>Getting a mutation-killing case took two attempts, and the reason is worth recording.
    /// Overlapping matches are mostly <i>harmless</i>: Apply advances a cursor monotonically and
    /// ignores any match starting behind it, so a nested key like <c>x}}</c> inside <c>{{x}}</c>
    /// is silently dropped and the output is unchanged. The damage needs keys that <b>chain</b> —
    /// <c>ab</c> and <c>bc</c> over <c>abc</c>. A one-character advance records <c>ab</c> at 0 and
    /// then <c>bc</c> at 1; Apply skips the second as starting behind the cursor, but has already
    /// moved the cursor to its <i>end</i> — swallowing the <c>c</c> entirely and emitting
    /// <c>"X"</c> where <c>"Xc"</c> is correct. Silent character loss, no exception.</para>
    /// </summary>
    [Fact]
    public void ScanningResumesAfterAMatchSoMatchesNeverOverlap()
    {
        var nodes = Split("abc");

        Assert.True(Apply(nodes, new Dictionary<string, string>
        {
            ["ab"] = "X",
            ["bc"] = "Y",
        }));

        // "ab" matches at 0 and the scan resumes at 2, where "c" matches nothing and is copied
        // through. "bc" never matches, because its only occurrence starts inside a consumed match.
        Assert.Equal("Xc", Merged(nodes));
    }

    [Fact]
    public void AKeyWhoseValueContainsItselfTerminates()
    {
        var nodes = Split("{{x}}");

        Assert.True(Apply(nodes, new Dictionary<string, string> { ["{{x}}"] = "a{{x}}b" }));

        Assert.Equal("a{{x}}b", Merged(nodes));
    }

    [Fact]
    public void AnEmptyValueDeletesTheMatch()
    {
        var nodes = Split("a{{gone}}b");

        Assert.True(Apply(nodes, new Dictionary<string, string> { ["{{gone}}"] = "" }));

        Assert.Equal("ab", Merged(nodes));
    }

    [Fact]
    public void ReportsNoChangeWhenNothingMatches()
    {
        var nodes = Split("nothing here");

        Assert.False(Apply(nodes, new Dictionary<string, string> { ["{{x}}"] = "V" }));

        Assert.Equal(0, nodes[0].Writes);
    }

    /// <summary>
    /// A match whose replacement equals the matched text changes nothing, so no run is written —
    /// which keeps a no-op ReplaceText from rewriting every run's element for nothing.
    /// </summary>
    [Fact]
    public void ReportsNoChangeWhenTheValueEqualsTheKey()
    {
        var nodes = Split("a{{x}}b");

        Assert.False(Apply(nodes, new Dictionary<string, string> { ["{{x}}"] = "{{x}}" }));

        Assert.Equal(0, nodes[0].Writes);
    }

    [Fact]
    public void HandlesNoNodesAndEmptyText()
    {
        var replacements = new Dictionary<string, string> { ["{{x}}"] = "V" };

        Assert.False(Apply([], replacements));
        Assert.False(Apply(Split("", "", ""), replacements));
    }

    [Fact]
    public void EmptyKeysAreIgnoredRatherThanMatchingEverywhere()
    {
        var nodes = Split("abc");

        // An empty key matches at every offset; taking it literally would insert the value between
        // every character, or loop forever.
        Assert.False(Apply(nodes, new Dictionary<string, string> { [""] = "X" }));

        Assert.Equal("abc", Merged(nodes));
    }

    /// <summary>
    /// A run whose text reads back as null is treated as empty. Both OOXML element wrappers can
    /// return null for an element that exists with no text child.
    /// </summary>
    [Fact]
    public void ARunWhoseTextIsNullIsTreatedAsEmpty()
    {
        var nodes = new List<Node> { new(null!), new("{{x}}") };

        Assert.True(RunTextSplicer.Apply(
            nodes, n => n.Text, (n, v) => n.Set(v),
            new Dictionary<string, string> { ["{{x}}"] = "V" }));

        Assert.Equal("V", Merged(nodes));
    }

    [Fact]
    public void SubstitutesEveryOccurrenceNotJustTheFirst()
    {
        var nodes = Split("{{x}} and {{x}} and ", "{{x}}");

        Assert.True(Apply(nodes, new Dictionary<string, string> { ["{{x}}"] = "V" }));

        Assert.Equal("V and V and V", Merged(nodes));
    }

    [Fact]
    public void SubstitutesTwoDifferentPlaceholdersInOnePass()
    {
        var nodes = Split("{{a}}-", "{{b}}");

        Assert.True(Apply(nodes, new Dictionary<string, string>
        {
            ["{{a}}"] = "1",
            ["{{b}}"] = "2",
        }));

        Assert.Equal("1-2", Merged(nodes));
    }

    /// <summary>A match starting exactly on a run boundary, which is the off-by-one's home.</summary>
    [Fact]
    public void HandlesAMatchThatStartsExactlyAtARunBoundary()
    {
        var nodes = Split("prefix", "{{x}}", "suffix");

        Assert.True(Apply(nodes, new Dictionary<string, string> { ["{{x}}"] = "V" }));

        Assert.Equal("prefixVsuffix", Merged(nodes));
        Assert.Equal(0, nodes[0].Writes);
        Assert.Equal(0, nodes[2].Writes);
    }

    [Fact]
    public void HandlesEmptyRunsInsideAMatch()
    {
        var nodes = Split("{{", "", "x", "", "}}");

        Assert.True(Apply(nodes, new Dictionary<string, string> { ["{{x}}"] = "V" }));

        Assert.Equal("V", Merged(nodes));
    }
}
