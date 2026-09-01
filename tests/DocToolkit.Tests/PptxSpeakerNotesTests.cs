using System;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace DocToolkit.Tests;

/// <summary>
/// A109: <see cref="PresentationEditor.ReadNotes(byte[], int)"/> and
/// <see cref="PresentationEditor.SetNotes(byte[], int, string)"/>.
/// </summary>
public class PptxSpeakerNotesTests
{
    private static byte[] TwoSlideDeck() => PresentationEditor.Create(new[]
    {
        PptxSlide.Titled("One", "first bullet"),
        PptxSlide.Titled("Two", "second bullet"),
    });

    [Fact]
    public void SetNotes_ThenReadNotes_RoundTrips()
    {
        var deck = PresentationEditor.SetNotes(TwoSlideDeck(), 1, "Mention the migration risk.");

        Assert.Equal("Mention the migration risk.", PresentationEditor.ReadNotes(deck, 1));
    }

    [Fact]
    public void ReadNotes_OnASlideWithNone_ReturnsEmptyRatherThanNull()
    {
        // Measured before the API was written: every slide carries a notes object whose text is
        // empty until something writes to it, so there is no "has notes" state to distinguish
        // from "notes are blank". The API does not invent one, and this pins that.
        var notes = PresentationEditor.ReadNotes(TwoSlideDeck(), 2);

        Assert.NotNull(notes);
        Assert.Equal(string.Empty, notes);
    }

    [Fact]
    public void SetNotes_LeavesEveryOtherSlideAlone()
    {
        // The silent-loss shape: editing one slide's notes and disturbing another's would pass
        // any test that only asserts the slide it just wrote.
        var deck = PresentationEditor.SetNotes(TwoSlideDeck(), 1, "FIRST-NOTES");
        deck = PresentationEditor.SetNotes(deck, 2, "SECOND-NOTES");

        Assert.Equal("FIRST-NOTES", PresentationEditor.ReadNotes(deck, 1));
        Assert.Equal("SECOND-NOTES", PresentationEditor.ReadNotes(deck, 2));
    }

    [Fact]
    public void SetNotes_ReplacesRatherThanAppends_AndAnEmptyStringClearsThem()
    {
        var deck = PresentationEditor.SetNotes(TwoSlideDeck(), 1, "ORIGINAL");
        deck = PresentationEditor.SetNotes(deck, 1, "REPLACEMENT");

        var replaced = PresentationEditor.ReadNotes(deck, 1);
        Assert.Equal("REPLACEMENT", replaced);
        Assert.DoesNotContain("ORIGINAL", replaced, StringComparison.Ordinal);

        deck = PresentationEditor.SetNotes(deck, 1, string.Empty);
        Assert.Equal(string.Empty, PresentationEditor.ReadNotes(deck, 1));
    }

    [Fact]
    public void SetNotes_DoesNotDisturbTheSlideBodies()
    {
        // The WHOLE extracted text, compared as a sequence, rather than probing one entry. Two
        // reasons: any disturbance anywhere in the deck fails this, and an earlier version of this
        // test asserted against `text[0]` believing it held slide 1's whole text. It does not -
        // ExtractText returns one entry per text BODY, so index 0 is the first title. That version
        // failed for a reason that had nothing to do with SetNotes, which preserves the bodies
        // exactly: measured identical before and after.
        var before = PresentationEditor.ExtractText(TwoSlideDeck());
        var after = PresentationEditor.ExtractText(PresentationEditor.SetNotes(TwoSlideDeck(), 1, "Notes only."));

        Assert.Equal(before, after);
    }

    [Fact]
    public void ExtractText_DoesNotReturnTheNotes()
    {
        // Stated in ReadNotes' own remarks, so it is pinned rather than left as prose: notes are
        // a separate surface a reader has to ask for, and a change that folded them into
        // ExtractText would alter every existing caller's output.
        var deck = PresentationEditor.SetNotes(TwoSlideDeck(), 1, "NOTES-MARKER");

        foreach (var slide in PresentationEditor.ExtractText(deck))
        {
            Assert.DoesNotContain("NOTES-MARKER", slide, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void ReadNotes_AndSetNotes_RejectABadIndexAndBadArguments()
    {
        var deck = TwoSlideDeck();

        Assert.Throws<ArgumentOutOfRangeException>(() => PresentationEditor.ReadNotes(deck, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => PresentationEditor.ReadNotes(deck, 3));
        Assert.Throws<ArgumentOutOfRangeException>(() => PresentationEditor.SetNotes(deck, 3, "x"));

        Assert.Throws<ArgumentNullException>(() => PresentationEditor.ReadNotes(null!, 1));
        Assert.Throws<ArgumentNullException>(() => PresentationEditor.SetNotes(deck, 1, null!));

        Assert.Equal("pptx", Assert.Throws<ArgumentException>(
            () => PresentationEditor.ReadNotes(Array.Empty<byte>(), 1)).ParamName);
        Assert.Equal("pptx", Assert.Throws<ArgumentException>(
            () => PresentationEditor.SetNotes(Array.Empty<byte>(), 1, "x")).ParamName);
    }

    [Fact]
    public async Task ReadNotesAsync_MatchesTheByteArrayOverload()
    {
        var deck = PresentationEditor.SetNotes(TwoSlideDeck(), 1, "ASYNC-NOTES");

        using var source = new MemoryStream(deck, writable: false);
        Assert.Equal("ASYNC-NOTES", await PresentationEditor.ReadNotesAsync(source, 1));
    }

    [Fact]
    public async Task SetNotesAsync_MatchesTheByteArrayOverload()
    {
        using var source = new MemoryStream(TwoSlideDeck(), writable: false);
        using var destination = new MemoryStream();

        await PresentationEditor.SetNotesAsync(source, 2, "WRITTEN-ASYNC", destination);

        Assert.Equal("WRITTEN-ASYNC", PresentationEditor.ReadNotes(destination.ToArray(), 2));
    }

}
