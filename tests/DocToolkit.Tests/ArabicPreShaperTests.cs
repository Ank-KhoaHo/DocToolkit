namespace DocToolkit.Tests;

/// <summary>
/// The Arabic pre-shaper (A124) turns logical Arabic-script letters into their contextual
/// presentation forms, still in logical order, one output string per input char.
/// </summary>
public class ArabicPreShaperTests
{
    private const string Ba = "\u0628";
    private const string Lam = "\u0644";
    private const string Alef = "\u0627";

    private static string[] Shape(string s) => ArabicPreShaper.Shape(s);

    [Fact]
    public void Lone_letter_takes_its_isolated_form()
        => Assert.Equal(new[] { "\uFE8F" }, Shape(Ba));

    [Fact]
    public void Two_joining_letters_take_initial_then_final()
        => Assert.Equal(new[] { "\uFE91", "\uFE90" }, Shape(Ba + Ba));

    [Fact]
    public void Three_joining_letters_take_initial_medial_final()
        => Assert.Equal(new[] { "\uFE91", "\uFE92", "\uFE90" }, Shape(Ba + Ba + Ba));

    [Fact]
    public void Lam_alef_is_one_ligature_and_leaves_the_alef_slot_empty()
        => Assert.Equal(new[] { "\uFEFB", "" }, Shape(Lam + Alef));

    [Fact]
    public void Lam_alef_after_a_joining_letter_takes_the_final_ligature()
        => Assert.Equal(new[] { "\uFE91", "\uFEFC", "" }, Shape(Ba + Lam + Alef));

    [Theory]
    [InlineData("\u0622", "\uFEF5")]
    [InlineData("\u0623", "\uFEF7")]
    [InlineData("\u0625", "\uFEF9")]
    [InlineData("\u0627", "\uFEFB")]
    public void Lam_alef_ligature_follows_the_alef_variant(string alef, string isolated)
    {
        var shaped = Shape(Lam + alef);
        Assert.Equal(isolated, shaped[0]);
        Assert.Equal("", shaped[1]);
        Assert.Equal(((char)(isolated[0] + 1)).ToString(), Shape(Ba + Lam + alef)[1]);
    }

    [Fact]
    public void A_right_joining_letter_breaks_the_chain()
        => Assert.Equal(new[] { "\uFEA9", "\uFE8F" }, Shape("\u062F" + Ba));

    [Fact]
    public void Persian_letters_take_their_presentation_forms()
    {
        Assert.Equal(new[] { "\uFB58", "\uFB57" }, Shape("\u067E\u067E"));   // pe
        Assert.Equal(new[] { "\uFB7C", "\uFB7B" }, Shape("\u0686\u0686"));   // che
        Assert.Equal(new[] { "\uFB8A" }, Shape("\u0698"));                    // zhe
        Assert.Equal(new[] { "\uFB94", "\uFB93" }, Shape("\u06AF\u06AF"));   // gaf
        Assert.Equal(new[] { "\uFBFE", "\uFBFD" }, Shape("\u06CC\u06CC"));   // farsi yeh
        Assert.Equal(new[] { "\uFB56" }, Shape("\u067E"));                    // lone pe is isolated
        Assert.Equal(new[] { "\uFBFC" }, Shape("\u06CC"));                    // lone yeh is isolated
        Assert.Equal(new[] { "\uFE91", "\uFB8B" }, Shape(Ba + "\u0698"));
    }

    [Fact]
    public void Harakat_are_transparent_to_joining()
    {
        // ba + fatha + ba: the first ba is initial and the second final, the fatha stays as is.
        Assert.Equal(new[] { "\uFE91", "\u064E", "\uFE90" }, Shape(Ba + "\u064E" + Ba));
        // dagger alef U+0670 is transparent too.
        Assert.Equal(new[] { "\uFE91", "\u0670", "\uFE90" }, Shape(Ba + "\u0670" + Ba));
    }

    [Fact]
    public void Lam_alef_joins_across_a_fatha()
        => Assert.Equal(new[] { "\uFEFB", "\u064E", "" }, Shape(Lam + "\u064E" + Alef));

    [Fact]
    public void Zero_width_non_joiner_breaks_joining()
        => Assert.Equal(new[] { "\uFE8F", "\u200C", "\uFE8F" }, Shape(Ba + "\u200C" + Ba));

    [Fact]
    public void Tatweel_joins_and_stays_in_the_text()
        => Assert.Equal(new[] { "\uFE91", "\u0640", "\uFE90" }, Shape(Ba + "\u0640" + Ba));

    [Fact]
    public void Latin_and_digits_are_untouched_beside_Arabic()
    {
        const string text = "Invoice \u0641\u0627\u062A\u0648\u0631\u0629 12";
        var shaped = Shape(text);
        Assert.Equal("Invoice ", string.Concat(shaped.Take(8)));
        Assert.Equal(" 12", string.Concat(shaped.Skip(shaped.Length - 3)));
        // fa initial, alef final (breaks the chain), teh initial, waw final, reh isolated, teh marbuta isolated.
        Assert.Equal("\uFED3\uFE8E\uFE97\uFEEE\uFEAD\uFE93", string.Concat(shaped.Skip(8).Take(6)));
    }

    [Fact]
    public void Output_has_one_entry_per_input_char()
    {
        foreach (var s in new[] { "", "abc", Lam + Alef, Ba + Lam + Alef + "\u064E", "x \u0641\u0627\u062A\u0648\u0631\u0629 1" })
            Assert.Equal(s.Length, Shape(s).Length);
    }

    [Fact]
    public void A_word_joins_across_what_were_run_boundaries()
    {
        // kaf-alef split across two runs would be shaped as one container string by the caller.
        var shaped = Shape("\u06A9\u0627" + "\u0644\u0627");
        Assert.Equal(new[] { "\uFB90", "\uFE8E", "\uFEFB", "" }, shaped);
        // Same chars, ba + ba as two "runs": the join crosses the boundary.
        Assert.Equal(new[] { "\uFE91", "\uFE90" }, Shape(Ba + Ba));
    }
}
