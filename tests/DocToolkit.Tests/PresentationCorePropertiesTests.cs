using System.IO.Packaging;

namespace DocToolkit.Tests;

/// <summary>
/// <see cref="PresentationEditor.WithMetadata"/> against the number of core-properties
/// relationships a deck carries (A123).
///
/// <b>The retry in <c>WithMetadata</c> was written for a deck with NO such relationship</b>, which
/// OfficeIMO 3.2.6 turned into two on save and then refused. Measured on 3.5.3: that deck saves on the
/// first attempt with exactly one, so the retry is no longer needed for it. The retry is still entered
/// for a deck that already has TWO, which is the refusal these tests keep alive.
/// </summary>
public class PresentationCorePropertiesTests
{
    private const string CoreType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

    private static int CoreRelationships(byte[] pptx)
    {
        using var ms = new MemoryStream(pptx, writable: false);
        using var package = Package.Open(ms, FileMode.Open, FileAccess.Read);
        return package.GetRelationshipsByType(CoreType).Count();
    }

    private static byte[] WithASecondCoreRelationship(byte[] pptx)
    {
        var ms = new MemoryStream();
        ms.Write(pptx);
        ms.Position = 0;
        using (var package = Package.Open(ms, FileMode.Open, FileAccess.ReadWrite))
        {
            var existing = package.GetRelationshipsByType(CoreType).Single();
            package.CreateRelationship(existing.TargetUri, TargetMode.Internal, CoreType, "rIdSecondCore");
        }

        return ms.ToArray();
    }

    [Fact]
    public void WithMetadata_ADeckBuiltByCreate_HasNoCoreRelationship_AndIsStampedWithExactlyOne()
    {
        var fresh = PresentationEditor.Create([PptxSlide.Titled("Slide")]);
        Assert.Equal(0, CoreRelationships(fresh));

        var stamped = PresentationEditor.WithMetadata(fresh, new DocumentMetadata { Title = "Stamped" });

        Assert.Equal(1, CoreRelationships(stamped));
        Assert.Equal("Stamped", PresentationEditor.ReadMetadata(stamped).Title);
    }

    [Fact]
    public void WithMetadata_ADeckWithTwoCoreRelationships_IsRefused_AndTheRetryDoesNotRepairIt()
    {
        var stamped = PresentationEditor.WithMetadata(
            PresentationEditor.Create([PptxSlide.Titled("Slide")]), new DocumentMetadata { Title = "First" });
        var two = WithASecondCoreRelationship(stamped);
        Assert.Equal(2, CoreRelationships(two));

        // The seed step only adds a relationship when there is none, so a deck with two is passed to
        // OfficeIMO unchanged on the retry and refused again. The refusal of the retry is what
        // reaches the caller: the exception thrown inside a catch block is not wrapped by the
        // sibling catch of the same try.
        var refusal = Assert.Throws<FileFormatException>(
            () => PresentationEditor.WithMetadata(two, new DocumentMetadata { Title = "Second" }));

        Assert.Contains("Core Properties", refusal.Message, StringComparison.Ordinal);
        Assert.Equal(2, CoreRelationships(two));
    }
}
