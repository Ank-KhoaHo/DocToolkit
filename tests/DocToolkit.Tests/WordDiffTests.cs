using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// The word-sequence diff behind document comparison (A118), tested with no OOXML in sight.
///
/// <b>This is the half most likely to be wrong in ways a document-level test cannot localise</b>,
/// so it is exercised directly against string sequences where a failure names the input rather than
/// a .docx. The property test below is the real guarantee; the worked examples exist so a failure
/// is readable.
/// </summary>
public class WordDiffTests
{
    private static List<WordDiffSpan> Diff(string original, string revised) =>
        WordDiff.Diff(WordDiff.Split(original), WordDiff.Split(revised));

    private static string Rebuild(IEnumerable<WordDiffSpan> spans, WordDiffKind skip) =>
        string.Concat(spans.Where(s => s.Kind != skip).SelectMany(s => s.Words));

    // ---------- the property that matters ----------

    /// <summary>
    /// THE test. Whatever the diff does, dropping the insertions must rebuild the original exactly
    /// and dropping the deletions must rebuild the revision exactly. A diff that loses, duplicates
    /// or reorders a word fails here however plausible its span list looks.
    /// </summary>
    [Theory]
    [InlineData("", "")]
    [InlineData("one two three", "one two three")]
    [InlineData("one two three", "one three")]
    [InlineData("one three", "one two three")]
    [InlineData("", "all new words here")]
    [InlineData("everything goes away", "")]
    [InlineData("the quick brown fox", "the slow brown dog")]
    [InlineData("a b c d e f g", "g f e d c b a")]
    [InlineData("repeated repeated repeated", "repeated repeated")]
    [InlineData("Acme Corporation invoice 42", "Acme Corp invoice 43 final")]
    public void DroppingEachSideRebuildsTheOtherDocumentExactly(string original, string revised)
    {
        var spans = Diff(original, revised);

        Assert.Equal(original, Rebuild(spans, WordDiffKind.Inserted));
        Assert.Equal(revised, Rebuild(spans, WordDiffKind.Deleted));
    }

    /// <summary>
    /// The same property over generated input, because hand-picked cases systematically miss the
    /// arrangement nobody thought of — the reasoning <c>RunTextSplicerTests</c> already applies to
    /// run splits.
    /// </summary>
    [Fact]
    public void TheRebuildPropertyHoldsOverGeneratedPairs()
    {
        var rng = new Random(20260903);
        var vocabulary = new[] { "alpha ", "beta ", "gamma ", "delta ", "epsilon " };

        for (var iteration = 0; iteration < 400; iteration++)
        {
            var original = string.Concat(Enumerable.Range(0, rng.Next(0, 12))
                .Select(_ => vocabulary[rng.Next(vocabulary.Length)]));
            var revised = string.Concat(Enumerable.Range(0, rng.Next(0, 12))
                .Select(_ => vocabulary[rng.Next(vocabulary.Length)]));

            var spans = WordDiff.Diff(WordDiff.Split(original), WordDiff.Split(revised));

            Assert.Equal(original, Rebuild(spans, WordDiffKind.Inserted));
            Assert.Equal(revised, Rebuild(spans, WordDiffKind.Deleted));
        }
    }

    // ---------- the negative control ----------

    /// <summary>
    /// Comparing something with itself must report NO change. Nothing else proves the diff can
    /// return nothing — every assertion above is satisfied by a diff that deletes everything and
    /// re-inserts it.
    /// </summary>
    [Theory]
    [InlineData("")]
    [InlineData("one")]
    [InlineData("a longer sentence with several words in it")]
    [InlineData("repeated repeated repeated")]
    public void ADocumentComparedWithItselfHasNoInsertionsOrDeletions(string text)
    {
        var spans = Diff(text, text);

        Assert.DoesNotContain(spans, s => s.Kind == WordDiffKind.Inserted);
        Assert.DoesNotContain(spans, s => s.Kind == WordDiffKind.Deleted);
    }

    // ---------- worked examples, so a failure reads ----------

    [Fact]
    public void ReportsOnlyTheWordThatChanged()
    {
        // The discriminating case for "marks everything changed", which passes any test asserting
        // merely that a difference was found.
        var spans = Diff("the quick brown fox", "the quick red fox");

        Assert.Equal(["red "], Assert.Single(spans, s => s.Kind == WordDiffKind.Inserted).Words);
        Assert.Equal(["brown "], Assert.Single(spans, s => s.Kind == WordDiffKind.Deleted).Words);
    }

    [Fact]
    public void MergesAdjacentWordsIntoOneSpan()
    {
        // So a caller emits one w:ins per inserted phrase rather than one per word.
        var spans = Diff("start end", "start one two three end");

        Assert.Equal(["one ", "two ", "three "], Assert.Single(spans, s => s.Kind == WordDiffKind.Inserted).Words);
    }

    [Fact]
    public void KeepsTheWhitespaceThatFollowsEachWord()
    {
        // A split that dropped trailing whitespace would make every rebuild lossy, and the loss
        // would surface as spurious differences rather than as an error.
        Assert.Equal(["one ", "two  ", "three"], WordDiff.Split("one two  three"));
    }

    [Fact]
    public void SplitsNothingIntoNoWords()
    {
        Assert.Empty(WordDiff.Split(string.Empty));
    }

    /// <summary>
    /// Two texts with nothing in common past the size bound are reported as a wholesale
    /// replacement rather than building a table too large to be worth it. That is a correct answer,
    /// just coarse - and it is the only way the bound is reachable, since the shared prefix and
    /// suffix are removed first.
    /// </summary>
    [Fact]
    public void TwoEntirelyDifferentLongTextsFallBackToAWholesaleReplacement()
    {
        var original = string.Concat(Enumerable.Range(0, 2100).Select(i => $"a{i} "));
        var revised = string.Concat(Enumerable.Range(0, 2100).Select(i => $"b{i} "));

        var spans = Diff(original, revised);

        // Everything deleted then everything inserted, and nothing marked Same.
        Assert.DoesNotContain(spans, s => s.Kind == WordDiffKind.Same);
        Assert.Equal(original, Rebuild(spans, WordDiffKind.Inserted));
        Assert.Equal(revised, Rebuild(spans, WordDiffKind.Deleted));
    }

    /// <summary>
    /// The positive control for the bound: two texts of the same size that DO share content must
    /// still get a real diff rather than the fallback.
    /// </summary>
    [Fact]
    public void PositiveControl_ALongPairThatSharesContentStillGetsARealDiff()
    {
        var shared = string.Concat(Enumerable.Range(0, 2100).Select(i => $"w{i} "));

        var spans = Diff(shared + "tail", shared + "different");

        Assert.Contains(spans, s => s.Kind == WordDiffKind.Same);
    }
}
