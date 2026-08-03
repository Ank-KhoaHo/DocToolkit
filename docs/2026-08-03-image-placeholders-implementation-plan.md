# Image placeholders — implementation plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development
> (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use
> checkbox (`- [ ]`) syntax for tracking.

**Goal:** `DocxEditor.ReplaceImage` swaps a text placeholder for an inline image — logo, signature,
QR code — anywhere `ReplaceText` reaches, sized from the image's own header.

**Architecture:** A pure `ImageInspector` detects format and pixel size from magic bytes and headers,
with no image-decoding dependency. `DocxEditor` then adds an `ImagePart` to the part that owns the
paragraph, builds a DrawingML inline element, and splices it in where the matched text was.

**Tech Stack:** `DocumentFormat.OpenXml` 3.5.1, xunit.

Design doc: `docs/2026-08-03-image-placeholders-design.md`. Read it first — why there is no image
library, and why `RunTextSplicer` cannot be reused, are both there.

## Global Constraints

- **Branch from `main`, PR back into it.** `main` cannot be pushed directly.
- **Merging no longer publishes** — release-please keeps a Release PR open and a human merges it. So
  small PRs are cheap again; this does not have to be one giant branch.
- **Conventional Commits**; merge commits are exempt, everything else is checked.
- **Never add a `Co-Authored-By` trailer.**
- **Add no new package reference.** The entire design exists to avoid one. If you find yourself
  wanting an image library, stop and re-read the design doc's constraint section.
- **Never use `Descendants<Paragraph>()` for discovery** — use the `ReplaceIn`/`ReplaceInParagraph`
  pattern already in `DocxEditor`, which filters text to the paragraph that directly owns it.
- Build runs at **0 warnings** under `-warnaserror`. **`dotnet test --filter` does not pass that
  flag**, so run a full `dotnet build ... -warnaserror` before believing a task is done — filtered
  runs hid four nullable warnings across three tasks during A4.
- Targets `net8.0;net10.0`. Currently 252 tests → 504 results; this plan adds to that.

---

### Task 1: Read format and pixel size from the image header

**Files:**
- Create: `src/DocToolkit/ImageInspector.cs`
- Test: `tests/DocToolkit.Tests/ImageInspectorTests.cs`
- Test helper: `tests/DocToolkit.Tests/ImageFixtures.cs`

**Interfaces:**
- Consumes: nothing.
- Produces: `internal enum ImageFormat { Png, Jpeg }`,
  `internal readonly record struct ImageInfo(ImageFormat Format, int WidthPx, int HeightPx)`,
  and `internal static ImageInfo Inspect(byte[] image)`.

Pure functions over bytes — no OpenXml, no I/O. That is why this is its own task: the parsers are
the part most likely to be subtly wrong, and they can be pinned down without a document in sight.

- [ ] **Step 1: Build the fixtures**

Create `tests/DocToolkit.Tests/ImageFixtures.cs`:

```csharp
using System.IO.Compression;

namespace DocToolkit.Tests;

/// <summary>
/// Real image bytes, without an image library — the whole point of the design is that this
/// repo takes no decoder dependency, so the tests cannot use one either.
/// </summary>
internal static class ImageFixtures
{
    /// <summary>
    /// A valid PNG of <paramref name="width"/> x <paramref name="height"/>, built by hand.
    /// Deliberately used at odd sizes so an intrinsic-sizing assertion cannot pass by coincidence.
    /// </summary>
    public static byte[] Png(int width = 2, int height = 3)
    {
        var ihdr = new byte[13];
        WriteBigEndian(ihdr, 0, (uint)width);
        WriteBigEndian(ihdr, 4, (uint)height);
        ihdr[8] = 8;    // bit depth
        ihdr[9] = 6;    // colour type: RGBA
        // 10..12 stay zero: deflate, adaptive filtering, no interlace.

        // One zero byte per row for the filter, then RGBA per pixel.
        var raw = new byte[height * (1 + width * 4)];
        using var deflated = new MemoryStream();
        using (var zlib = new ZLibStream(deflated, CompressionLevel.Fastest, leaveOpen: true))
            zlib.Write(raw, 0, raw.Length);

        using var png = new MemoryStream();
        png.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });
        WriteChunk(png, "IHDR", ihdr);
        WriteChunk(png, "IDAT", deflated.ToArray());
        WriteChunk(png, "IEND", Array.Empty<byte>());
        return png.ToArray();
    }

    /// <summary>
    /// Real JPEG bytes — 256 x 144 — lifted from the PPTX fixture already in this repo, so no new
    /// binary file has to be committed and no .gitattributes question arises.
    /// </summary>
    public static byte[] Jpeg()
    {
        var pptx = Path.Combine(AppContext.BaseDirectory, "assets", "sample.pptx");
        using var zip = System.IO.Compression.ZipFile.OpenRead(pptx);
        var entry = zip.GetEntry("docProps/thumbnail.jpeg")
                    ?? throw new InvalidOperationException(
                        "sample.pptx no longer contains docProps/thumbnail.jpeg - the JPEG fixture needs a new source.");
        using var stream = entry.Open();
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }

    /// <summary>A GIF header - enough to be detected as an unsupported format.</summary>
    public static byte[] Gif() =>
        new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00, 0x01, 0x00 };

    private static void WriteChunk(Stream target, string type, byte[] data)
    {
        var length = new byte[4];
        WriteBigEndian(length, 0, (uint)data.Length);
        target.Write(length);

        var typed = new byte[4 + data.Length];
        for (var i = 0; i < 4; i++) typed[i] = (byte)type[i];
        data.CopyTo(typed, 4);
        target.Write(typed);

        var crc = new byte[4];
        WriteBigEndian(crc, 0, Crc32(typed));
        target.Write(crc);
    }

    private static void WriteBigEndian(byte[] buffer, int offset, uint value)
    {
        buffer[offset] = (byte)(value >> 24);
        buffer[offset + 1] = (byte)(value >> 16);
        buffer[offset + 2] = (byte)(value >> 8);
        buffer[offset + 3] = (byte)value;
    }

    private static uint Crc32(byte[] data)
    {
        var crc = 0xFFFFFFFFu;
        foreach (var b in data)
        {
            crc ^= b;
            for (var i = 0; i < 8; i++)
                crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320u : crc >> 1;
        }
        return crc ^ 0xFFFFFFFFu;
    }
}
```

- [ ] **Step 2: Write the failing tests**

Create `tests/DocToolkit.Tests/ImageInspectorTests.cs`:

```csharp
namespace DocToolkit.Tests;

public class ImageInspectorTests
{
    [Fact]
    public void Inspect_ReadsPngDimensionsFromTheIhdrChunk()
    {
        var info = ImageInspector.Inspect(ImageFixtures.Png(width: 2, height: 3));

        Assert.Equal(ImageFormat.Png, info.Format);
        Assert.Equal(2, info.WidthPx);
        Assert.Equal(3, info.HeightPx);
    }

    [Fact]
    public void Inspect_ReadsPngDimensionsThatNeedMoreThanOneByte()
    {
        // 300 x 260 exercises the big-endian assembly; a byte-order bug survives 2 x 3.
        var info = ImageInspector.Inspect(ImageFixtures.Png(width: 300, height: 260));

        Assert.Equal(300, info.WidthPx);
        Assert.Equal(260, info.HeightPx);
    }

    [Fact]
    public void Inspect_ReadsJpegDimensionsByWalkingToTheStartOfFrame()
    {
        var info = ImageInspector.Inspect(ImageFixtures.Jpeg());

        Assert.Equal(ImageFormat.Jpeg, info.Format);
        Assert.Equal(256, info.WidthPx);
        Assert.Equal(144, info.HeightPx);
    }

    [Fact]
    public void Inspect_RejectsAnUnsupportedFormatByName()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => ImageInspector.Inspect(ImageFixtures.Gif()));

        Assert.Contains("GIF", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Inspect_RejectsBytesThatAreNoImageAtAll()
    {
        var ex = Assert.Throws<DocumentConversionException>(
            () => ImageInspector.Inspect(new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 }));

        Assert.Contains("PNG", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("JPEG", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Inspect_RejectsATruncatedPng()
    {
        var truncated = ImageFixtures.Png()[..12];

        Assert.Throws<DocumentConversionException>(() => ImageInspector.Inspect(truncated));
    }
}
```

- [ ] **Step 3: Run them and watch them fail**

Run: `dotnet test tests/DocToolkit.Tests -c Release --filter FullyQualifiedName~ImageInspectorTests`
Expected: **compile failure** — `ImageInspector` does not exist.

- [ ] **Step 4: Implement the inspector**

Create `src/DocToolkit/ImageInspector.cs`:

```csharp
using System.Buffers.Binary;

namespace DocToolkit;

/// <summary>Image formats this library can embed.</summary>
internal enum ImageFormat
{
    Png,
    Jpeg,
}

/// <summary>What a supported image is, and how big it is in pixels.</summary>
internal readonly record struct ImageInfo(ImageFormat Format, int WidthPx, int HeightPx)
{
    /// <summary>The MIME type the corresponding OpenXml image part must declare.</summary>
    public string ContentType => Format == ImageFormat.Png ? "image/png" : "image/jpeg";
}

/// <summary>
/// Reads an image's format and pixel dimensions straight from its header.
///
/// Every image in a .docx carries an explicit size, so something has to supply the dimensions —
/// and the obvious candidate, SixLabors.ImageSharp, moved its later majors onto the same
/// revenue-gated licence SixLabors.Fonts is pinned at [1.0.x] to avoid. Adding it would mean a
/// second permanent CI guard against a licence this package exists to stay clear of. PNG and JPEG
/// headers are cheap to read directly, so they are read directly.
///
/// The format is decided by MAGIC BYTES, never by a filename or a caller's assertion: an image part
/// declaring image/png while holding JPEG bytes produces a blank frame in Word with no error.
/// </summary>
internal static class ImageInspector
{
    private static ReadOnlySpan<byte> PngSignature => new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

    public static ImageInfo Inspect(byte[] image)
    {
        ArgumentNullException.ThrowIfNull(image);

        if (StartsWith(image, PngSignature)) return InspectPng(image);
        if (image.Length >= 3 && image[0] == 0xFF && image[1] == 0xD8 && image[2] == 0xFF)
            return InspectJpeg(image);

        throw new DocumentConversionException(
            $"Unsupported image format ({DescribeFormat(image)}). Only PNG and JPEG can be embedded.");
    }

    /// <summary>
    /// PNG is fixed-layout: the IHDR chunk is always first, and its width and height are big-endian
    /// uint32 at bytes 16 and 20.
    /// </summary>
    private static ImageInfo InspectPng(byte[] image)
    {
        if (image.Length < 24)
            throw new DocumentConversionException("PNG image is truncated: no IHDR chunk.");

        var width = (int)BinaryPrimitives.ReadUInt32BigEndian(image.AsSpan(16, 4));
        var height = (int)BinaryPrimitives.ReadUInt32BigEndian(image.AsSpan(20, 4));

        if (width <= 0 || height <= 0)
            throw new DocumentConversionException($"PNG image reports a nonsensical size ({width}x{height}).");

        return new ImageInfo(ImageFormat.Png, width, height);
    }

    /// <summary>
    /// JPEG has no fixed offset: walk the segment markers to a Start-Of-Frame, where height then
    /// width follow the one-byte sample precision. SOF0/1/2/3/5/6/7/9/10/11/13/14/15 all carry the
    /// size; C4 (Huffman table), C8 (reserved) and CC (arithmetic conditioning) do not.
    /// </summary>
    private static ImageInfo InspectJpeg(byte[] image)
    {
        var i = 2;
        while (i < image.Length - 1)
        {
            if (image[i] != 0xFF) { i++; continue; }

            var marker = image[i + 1];

            if (marker is >= 0xC0 and <= 0xCF && marker is not (0xC4 or 0xC8 or 0xCC))
            {
                if (i + 9 > image.Length)
                    throw new DocumentConversionException("JPEG image is truncated inside its frame header.");

                var height = BinaryPrimitives.ReadUInt16BigEndian(image.AsSpan(i + 5, 2));
                var width = BinaryPrimitives.ReadUInt16BigEndian(image.AsSpan(i + 7, 2));

                if (width == 0 || height == 0)
                    throw new DocumentConversionException($"JPEG image reports a nonsensical size ({width}x{height}).");

                return new ImageInfo(ImageFormat.Jpeg, width, height);
            }

            // Markers without a length payload.
            if (marker is 0xD8 or 0xD9 || marker is >= 0xD0 and <= 0xD7) { i += 2; continue; }

            if (i + 4 > image.Length)
                throw new DocumentConversionException("JPEG image is truncated inside a segment header.");

            i += 2 + BinaryPrimitives.ReadUInt16BigEndian(image.AsSpan(i + 2, 2));
        }

        throw new DocumentConversionException("JPEG image has no Start-Of-Frame segment, so its size is unknown.");
    }

    /// <summary>Names what the bytes look like, so the error says more than "invalid image".</summary>
    private static string DescribeFormat(byte[] image)
    {
        if (StartsWith(image, "GIF8"u8)) return "GIF";
        if (StartsWith(image, "BM"u8)) return "BMP";
        if (StartsWith(image, "RIFF"u8)) return "WebP or another RIFF container";
        if (StartsWith(image, "<svg"u8) || StartsWith(image, "<?xml"u8)) return "SVG or XML";
        if (image.Length < 8) return $"{image.Length} bytes, too short to identify";
        return "unrecognised";
    }

    private static bool StartsWith(byte[] image, ReadOnlySpan<byte> prefix) =>
        image.Length >= prefix.Length && image.AsSpan(0, prefix.Length).SequenceEqual(prefix);
}
```

- [ ] **Step 5: Run the tests**

Run: `dotnet test tests/DocToolkit.Tests -c Release --filter FullyQualifiedName~ImageInspectorTests`
Expected: PASS, 12 results (6 tests × 2 TFMs).

**If the JPEG test fails on the dimensions**, do not adjust the expected numbers — the fixture was
verified at 256 × 144 before this plan was written. A mismatch means the marker walk is wrong.

- [ ] **Step 6: Full build, then commit**

```bash
dotnet build DocToolkit.sln -c Release -warnaserror   # the filter above does NOT do this
dotnet test  DocToolkit.sln -c Release --no-build

git add src/DocToolkit/ImageInspector.cs tests/DocToolkit.Tests/ImageInspectorTests.cs tests/DocToolkit.Tests/ImageFixtures.cs
git commit -m "feat(core): read image format and size from the header

PNG and JPEG dimensions come from their own headers rather than from an
image library. SixLabors.ImageSharp's later majors moved onto the same
revenue-gated licence SixLabors.Fonts is pinned at [1.0.x] to avoid, so
taking it would have meant a second permanent CI guard against a licence
this package exists to stay clear of.

Format is decided by magic bytes, never a filename: an image part
declaring image/png while holding JPEG bytes renders as a blank frame in
Word with no error at all.

The JPEG fixture is lifted from the PPTX already in the repo, so no new
binary is committed; the PNG is hand-built at an odd size so an
intrinsic-sizing assertion cannot pass by coincidence."
```

---

### Task 2: Size resolution in EMUs

**Files:**
- Modify: `src/DocToolkit/ImageInspector.cs` (add `Sizing`)
- Test: `tests/DocToolkit.Tests/ImageInspectorTests.cs`

**Interfaces:**
- Consumes: `ImageInfo` from Task 1.
- Produces: `internal static (long WidthEmu, long HeightEmu) Resolve(ImageInfo info, double? widthPoints, double? heightPoints)`.

Its own task because the arithmetic has **no schema check behind it** — a factor-of-ten error
produces a valid document containing a wrongly sized image, which no validator and no text
assertion will ever catch.

- [ ] **Step 1: Write the failing tests**

```csharp
    [Theory]
    // 96 DPI: one pixel is 9,525 EMU. 2 x 3 px -> 19,050 x 28,575.
    [InlineData(2, 3, null, null, 19050L, 28575L)]
    // widthPoints only: 1 pt = 12,700 EMU, height scales to keep 2:3.
    [InlineData(2, 3, 10.0, null, 127000L, 190500L)]
    // heightPoints only: width scales to keep 2:3.
    [InlineData(2, 3, null, 30.0, 254000L, 381000L)]
    // both: exactly what was asked for, aspect ratio ignored.
    [InlineData(2, 3, 10.0, 10.0, 127000L, 127000L)]
    public void Resolve_ConvertsToEmusAndPreservesAspectWhenOnlyOneSideIsGiven(
        int widthPx, int heightPx, double? widthPoints, double? heightPoints,
        long expectedWidthEmu, long expectedHeightEmu)
    {
        var info = new ImageInfo(ImageFormat.Png, widthPx, heightPx);

        var (width, height) = ImageInspector.Resolve(info, widthPoints, heightPoints);

        Assert.Equal(expectedWidthEmu, width);
        Assert.Equal(expectedHeightEmu, height);
    }

    [Theory]
    [InlineData(0.0, null)]
    [InlineData(-5.0, null)]
    [InlineData(null, 0.0)]
    [InlineData(null, -5.0)]
    public void Resolve_RejectsANonPositiveSize(double? widthPoints, double? heightPoints)
    {
        var info = new ImageInfo(ImageFormat.Png, 2, 3);

        Assert.Throws<ArgumentOutOfRangeException>(
            () => ImageInspector.Resolve(info, widthPoints, heightPoints));
    }
```

- [ ] **Step 2: Run and watch fail**

Expected: compile failure — `Resolve` does not exist.

- [ ] **Step 3: Implement**

Add to `ImageInspector`:

```csharp
    private const long EmuPerPoint = 12700;
    private const long EmuPerPixelAt96Dpi = 9525;

    /// <summary>
    /// The size to write into the drawing, in EMUs.
    ///
    /// Neither dimension given: the image's intrinsic pixel size at 96 DPI, which is what Word
    /// assumes for an image with no DPI metadata. One given: the other scales to preserve the
    /// aspect ratio. Both given: exactly those, distortion accepted as the caller's choice.
    ///
    /// Nothing downstream validates these numbers — a factor-of-ten slip yields a perfectly valid
    /// document containing a wrongly sized image.
    /// </summary>
    public static (long WidthEmu, long HeightEmu) Resolve(
        ImageInfo info, double? widthPoints, double? heightPoints)
    {
        if (widthPoints is <= 0)
            throw new ArgumentOutOfRangeException(nameof(widthPoints), widthPoints, "Width must be positive.");
        if (heightPoints is <= 0)
            throw new ArgumentOutOfRangeException(nameof(heightPoints), heightPoints, "Height must be positive.");

        var intrinsicWidth = info.WidthPx * EmuPerPixelAt96Dpi;
        var intrinsicHeight = info.HeightPx * EmuPerPixelAt96Dpi;

        return (widthPoints, heightPoints) switch
        {
            (null, null) => (intrinsicWidth, intrinsicHeight),
            (not null, null) => Scale((long)(widthPoints.Value * EmuPerPoint), intrinsicWidth, intrinsicHeight),
            (null, not null) => Flip(Scale((long)(heightPoints.Value * EmuPerPoint), intrinsicHeight, intrinsicWidth)),
            _ => ((long)(widthPoints!.Value * EmuPerPoint), (long)(heightPoints!.Value * EmuPerPoint)),
        };

        static (long, long) Scale(long given, long intrinsicGiven, long intrinsicOther) =>
            (given, (long)(given * (double)intrinsicOther / intrinsicGiven));

        static (long, long) Flip((long First, long Second) pair) => (pair.Second, pair.First);
    }
```

- [ ] **Step 4: Run, then full build, then commit**

```bash
dotnet test tests/DocToolkit.Tests -c Release --filter FullyQualifiedName~ImageInspectorTests
dotnet build DocToolkit.sln -c Release -warnaserror
git add -A src/DocToolkit/ImageInspector.cs tests/DocToolkit.Tests/ImageInspectorTests.cs
git commit -m "feat(core): resolve image size to EMUs, preserving aspect ratio

Intrinsic size at 96 DPI when neither dimension is given, proportional
scaling when one is, exact when both. Tested by table rather than by
example because nothing downstream validates the arithmetic: a
factor-of-ten slip produces a valid document with a wrongly sized image,
which no schema check and no text assertion would catch."
```

---

### Task 3: Replace the placeholder with a drawing, in the body

**Files:**
- Create: `src/DocToolkit/DrawingFactory.cs`
- Modify: `src/DocToolkit/DocxEditor.cs`
- Test: `tests/DocToolkit.Tests/DocxEditorReplaceImageTests.cs`

**Interfaces:**
- Consumes: `ImageInspector.Inspect`/`Resolve`.
- Produces: `public static byte[] ReplaceImage(byte[] docx, string placeholder, byte[] image,
  double? widthPoints = null, double? heightPoints = null)`, plus private `ReplaceImageCore`.

- [ ] **Step 1: Write the failing test**

Create `tests/DocToolkit.Tests/DocxEditorReplaceImageTests.cs`:

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocToolkit.Tests;

public class DocxEditorReplaceImageTests
{
    private static void AssertValid(byte[] docx)
    {
        using var ms = new MemoryStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var errors = new OpenXmlValidator().Validate(doc).ToList();
        Assert.True(errors.Count == 0,
            "expected a schema-valid package, got:\n" +
            string.Join("\n", errors.Take(3).Select(e => "  " + e.Description)));
    }

    [Fact]
    public void ReplaceImage_SwapsThePlaceholderForADrawing()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("Logo: {{logo}} end")));

        var filled = DocxEditor.ReplaceImage(docx, "{{logo}}", ImageFixtures.Png());

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var body = doc.MainDocumentPart!.Document!.Body!;

        Assert.Single(body.Descendants<Drawing>());
        Assert.Single(doc.MainDocumentPart.ImageParts);

        // Only the matched span goes; the text around it stays in place.
        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("Logo: ", text);
        Assert.Contains(" end", text);
        Assert.DoesNotContain("{{logo}}", text);

        AssertValid(filled);
    }

    [Fact]
    public void ReplaceImage_SizesFromTheImageHeaderByDefault()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("{{logo}}")));

        var filled = DocxEditor.ReplaceImage(docx, "{{logo}}", ImageFixtures.Png(width: 2, height: 3));

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var extent = doc.MainDocumentPart!.Document!.Body!.Descendants<DW.Extent>().Single();

        Assert.Equal(2L * 9525, extent.Cx!.Value);
        Assert.Equal(3L * 9525, extent.Cy!.Value);
    }

    [Fact]
    public void ReplaceImage_ThrowsWhenThePlaceholderIsAbsent()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("nothing to replace")));

        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxEditor.ReplaceImage(docx, "{{logo}}", ImageFixtures.Png()));

        Assert.Contains("{{logo}}", ex.Message);
    }
}
```

- [ ] **Step 2: Run and watch fail**

Expected: compile failure — `ReplaceImage` does not exist.

- [ ] **Step 3: Build the drawing factory**

Create `src/DocToolkit/DrawingFactory.cs`:

```csharp
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace DocToolkit;

/// <summary>
/// Builds the DrawingML for an inline image.
///
/// Verbose, and deliberately so: this is the markup Word itself emits. The legacy VML alternative
/// (<c>w:pict</c>) is a fraction of the size and this repo even has VML fixtures already, but it is
/// deprecated and an image written that way looks subtly unlike one a human inserted.
/// </summary>
internal static class DrawingFactory
{
    /// <param name="relationshipId">The image part's relationship id, resolved in the OWNING part.</param>
    /// <param name="name">Shown in Word's selection and accessibility panes.</param>
    /// <param name="id">Must be unique across the whole document, or Word offers to repair the file.</param>
    public static Drawing InlineImage(string relationshipId, string name, uint id, long widthEmu, long heightEmu) =>
        new(new DW.Inline(
            new DW.Extent { Cx = widthEmu, Cy = heightEmu },
            new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
            new DW.DocProperties { Id = id, Name = name, Description = name },
            new DW.NonVisualGraphicFrameDrawingProperties(
                new A.GraphicFrameLocks { NoChangeAspect = true }),
            new A.Graphic(
                new A.GraphicData(
                    new PIC.Picture(
                        new PIC.NonVisualPictureProperties(
                            new PIC.NonVisualDrawingProperties { Id = 0U, Name = name },
                            new PIC.NonVisualPictureDrawingProperties()),
                        new PIC.BlipFill(
                            new A.Blip { Embed = relationshipId },
                            new A.Stretch(new A.FillRectangle())),
                        new PIC.ShapeProperties(
                            new A.Transform2D(
                                new A.Offset { X = 0L, Y = 0L },
                                new A.Extents { Cx = widthEmu, Cy = heightEmu }),
                            new A.PresetGeometry(new A.AdjustValueList())
                            {
                                Preset = A.ShapeTypeValues.Rectangle,
                            })))
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture",
                }))
        {
            DistanceFromTop = 0U,
            DistanceFromBottom = 0U,
            DistanceFromLeft = 0U,
            DistanceFromRight = 0U,
        });
}
```

- [ ] **Step 4: Implement `ReplaceImage` in `DocxEditor`**

Add, editing the file in place — never replacing it wholesale, it carries the package metadata's
sibling comments and every other public method:

```csharp
    /// <summary>
    /// Replaces every occurrence of <paramref name="placeholder"/> with <paramref name="image"/>,
    /// inline, across the body, headers, footers, footnotes and endnotes.
    ///
    /// Only the matched text is removed: text sharing a run with the placeholder keeps its place and
    /// its formatting, so <c>Signed: {{sig}} (authorised)</c> becomes <c>Signed: </c>, the image,
    /// then <c> (authorised)</c>.
    ///
    /// <paramref name="placeholder"/> is the literal text including braces, like
    /// <see cref="ReplaceText(byte[], IReadOnlyDictionary{string, string})"/> and unlike
    /// <see cref="FillRows"/>, whose keys are bare field names only because the collection name is
    /// already an argument there.
    ///
    /// Size is in points. Omit both and the image's intrinsic size is used, read from its own header
    /// at 96 DPI. Give one and the other scales to preserve the aspect ratio. Give both and the
    /// image is stretched to fit.
    ///
    /// PNG and JPEG only, detected from the image's magic bytes rather than any filename.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any of the three required arguments is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="docx"/> or <paramref name="image"/> is empty, or <paramref name="placeholder"/> is blank.</exception>
    /// <exception cref="ArgumentOutOfRangeException">A supplied size is zero or negative.</exception>
    /// <exception cref="DocumentConversionException">
    /// The image is neither PNG nor JPEG, the package could not be edited, or
    /// <paramref name="placeholder"/> does not appear anywhere — a call matching nothing is a bug,
    /// not a no-op.
    /// </exception>
    public static byte[] ReplaceImage(
        byte[] docx, string placeholder, byte[] image,
        double? widthPoints = null, double? heightPoints = null)
    {
        ArgumentNullException.ThrowIfNull(docx);
        ArgumentNullException.ThrowIfNull(placeholder);
        ArgumentNullException.ThrowIfNull(image);
        if (docx.Length == 0) throw new ArgumentException("DOCX content was empty.", nameof(docx));
        if (image.Length == 0) throw new ArgumentException("Image content was empty.", nameof(image));
        if (string.IsNullOrWhiteSpace(placeholder))
            throw new ArgumentException("Placeholder was blank.", nameof(placeholder));

        using var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        ms.Position = 0;

        ReplaceImageCore(ms, placeholder, image, widthPoints, heightPoints);
        return ms.ToArray();
    }
```

The core, and the splice. Add these privates too:

```csharp
    private static void ReplaceImageCore(
        MemoryStream ms, string placeholder, byte[] image, double? widthPoints, double? heightPoints)
    {
        var info = ImageInspector.Inspect(image);
        var (widthEmu, heightEmu) = ImageInspector.Resolve(info, widthPoints, heightPoints);
        var name = placeholder.Trim().Trim('{', '}').Trim();

        try
        {
            using (var doc = WordprocessingDocument.Open(ms, true))
            {
                var main = doc.MainDocumentPart
                           ?? throw new DocumentConversionException("Document has no main part.");
                var body = main.Document?.Body
                           ?? throw new DocumentConversionException("Document has no body.");

                // Unique across the WHOLE document: a duplicate wp:docPr id makes Word declare the
                // file corrupt and offer to repair it, so start above whatever is already there.
                var nextId = NextDrawingId(doc);
                var replaced = 0;

                replaced += InsertImagesIn(main, body, placeholder, image, info, widthEmu, heightEmu, name, ref nextId);
                main.Document!.Save();

                foreach (var part in main.HeaderParts)
                {
                    if (part.Header is null) continue;
                    replaced += InsertImagesIn(part, part.Header, placeholder, image, info, widthEmu, heightEmu, name, ref nextId);
                    part.Header.Save();
                }

                foreach (var part in main.FooterParts)
                {
                    if (part.Footer is null) continue;
                    replaced += InsertImagesIn(part, part.Footer, placeholder, image, info, widthEmu, heightEmu, name, ref nextId);
                    part.Footer.Save();
                }

                if (main.FootnotesPart?.Footnotes is { } footnotes)
                {
                    replaced += InsertImagesIn(main.FootnotesPart, footnotes, placeholder, image, info, widthEmu, heightEmu, name, ref nextId);
                    footnotes.Save();
                }

                if (main.EndnotesPart?.Endnotes is { } endnotes)
                {
                    replaced += InsertImagesIn(main.EndnotesPart, endnotes, placeholder, image, info, widthEmu, heightEmu, name, ref nextId);
                    endnotes.Save();
                }

                if (replaced == 0)
                {
                    throw new DocumentConversionException(
                        $"The placeholder '{placeholder}' was not found, so there was nothing to replace.");
                }
            }

            ms.Position = 0;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to insert an image into the DOCX package.", ex);
        }
    }

    /// <summary>Highest existing wp:docPr id anywhere in the package, plus one.</summary>
    private static uint NextDrawingId(WordprocessingDocument doc)
    {
        var highest = 0U;
        foreach (var part in doc.MainDocumentPart!.Parts.Select(p => p.OpenXmlPart).Prepend(doc.MainDocumentPart))
        {
            if (part.RootElement is null) continue;
            foreach (var properties in part.RootElement.Descendants<DW.DocProperties>())
                if (properties.Id?.Value is { } value && value > highest) highest = value;
        }

        return highest + 1;
    }
```

- [ ] **Step 5: Implement the splice**

This is the part that cannot reuse `RunTextSplicer` — that writes text, and this inserts an element:

```csharp
    private static int InsertImagesIn(
        OpenXmlPartContainer owner, OpenXmlElement root, string placeholder, byte[] image,
        ImageInfo info, long widthEmu, long heightEmu, string name, ref uint nextId)
    {
        var inserted = 0;

        foreach (var paragraph in root.Descendants<Paragraph>().ToList())
        {
            // Same scoping as ReplaceInParagraph: only the text this paragraph directly owns, so a
            // text box nested in one of its runs is visited on its own rather than folded in.
            var texts = paragraph.Descendants<Text>()
                                 .Where(t => t.Ancestors<Paragraph>().FirstOrDefault() == paragraph)
                                 .ToList();
            if (texts.Count == 0) continue;

            var merged = string.Concat(texts.Select(t => t.Text));
            if (!merged.Contains(placeholder, StringComparison.Ordinal)) continue;

            // Right to left, so earlier offsets stay valid as later matches are spliced out.
            var offsets = new List<int>();
            for (var at = merged.IndexOf(placeholder, StringComparison.Ordinal);
                 at >= 0;
                 at = merged.IndexOf(placeholder, at + placeholder.Length, StringComparison.Ordinal))
            {
                offsets.Add(at);
            }

            for (var i = offsets.Count - 1; i >= 0; i--)
            {
                var relationshipId = AddImagePart(owner, image, info);
                var drawing = DrawingFactory.InlineImage(relationshipId, name, nextId++, widthEmu, heightEmu);
                SpliceDrawingIn(texts, offsets[i], placeholder.Length, drawing);
                inserted++;
            }
        }

        return inserted;
    }

    private static string AddImagePart(OpenXmlPartContainer owner, byte[] image, ImageInfo info)
    {
        // The part must belong to the container that owns the paragraph. A header's image added to
        // the main document part yields a relationship id that resolves in the wrong scope: Word
        // opens the file and simply shows nothing where the image should be.
        var part = owner.AddNewPart<ImagePart>(info.ContentType);
        using (var stream = part.GetStream(FileMode.Create))
            stream.Write(image, 0, image.Length);

        return owner.GetIdOfPart(part);
    }

    /// <summary>
    /// Removes <paramref name="length"/> characters at <paramref name="start"/> from the
    /// concatenation of <paramref name="texts"/>, and puts <paramref name="drawing"/> there instead.
    /// Text outside the match keeps its run and its formatting.
    /// </summary>
    private static void SpliceDrawingIn(List<Text> texts, int start, int length, Drawing drawing)
    {
        var end = start + length;
        var position = 0;
        Run? anchor = null;
        string suffix = string.Empty;

        foreach (var node in texts)
        {
            var nodeStart = position;
            var nodeEnd = position + node.Text.Length;
            position = nodeEnd;

            if (nodeEnd <= start || nodeStart >= end) continue;   // untouched by the match

            var keepBefore = start > nodeStart ? node.Text[..(start - nodeStart)] : string.Empty;
            var keepAfter = end < nodeEnd ? node.Text[(end - nodeStart)..] : string.Empty;

            if (anchor is null)
            {
                node.Text = keepBefore;
                anchor = node.Ancestors<Run>().First();
                suffix = keepAfter;
            }
            else
            {
                node.Text = keepAfter;
                if (keepAfter.Length == 0) node.Ancestors<Run>().First().Remove();
            }
        }

        if (anchor is null) return;

        var imageRun = new Run(drawing);
        anchor.InsertAfterSelf(imageRun);

        // A match wholly inside one run leaves a tail that needs its own run after the image.
        if (suffix.Length > 0)
        {
            imageRun.InsertAfterSelf(new Run(
                new Text(suffix) { Space = SpaceProcessingModeValues.Preserve }));
        }
    }
```

- [ ] **Step 6: Run the tests**

Run: `dotnet test tests/DocToolkit.Tests -c Release --filter FullyQualifiedName~DocxEditorReplaceImageTests`
Expected: PASS.

**If `AssertValid` reports schema errors**, the DrawingML shape is wrong — fix the factory, do not
weaken the assertion. A4 showed how easily tests pass against invalid documents.

- [ ] **Step 7: Full build and commit**

```bash
dotnet build DocToolkit.sln -c Release -warnaserror
dotnet test  DocToolkit.sln -c Release --no-build

git add src/DocToolkit/DrawingFactory.cs src/DocToolkit/DocxEditor.cs tests/DocToolkit.Tests/DocxEditorReplaceImageTests.cs
git commit -m "feat(core): add DocxEditor.ReplaceImage

Swaps a text placeholder for an inline image across the body, headers,
footers, footnotes and endnotes. Only the matched span is removed, so text
sharing a run with the placeholder keeps its place and formatting.

The splice is new code rather than a RunTextSplicer reuse: that helper maps
match offsets back onto runs and writes TEXT, and this has to remove a span
and insert an ELEMENT at that position. Same principle - never touch a run
the match does not overlap - different mechanism.

Image parts are added to the container that owns the paragraph. A header's
image added to the main document part gives a relationship id that resolves
in the wrong scope: Word opens the file and shows nothing."
```

---

### Task 4: Prove the traps are handled

**Files:**
- Modify: `tests/DocToolkit.Tests/DocxEditorReplaceImageTests.cs`
- Modify: `src/DocToolkit/DocxEditor.cs` (only if a test exposes a defect)

Each of these produces a document that **opens**, so none of them is caught by anything already in
the suite.

- [ ] **Step 1: A header image must be owned by the header part**

```csharp
    [Fact]
    public void ReplaceImage_PutsAHeaderImageInTheHeaderPart()
    {
        var docx = DocxFixtures.Build(
            headerText: "Company {{logo}}",
            footerText: null,
            DocxFixtures.P(DocxFixtures.R("body text")));

        var filled = DocxEditor.ReplaceImage(docx, "{{logo}}", ImageFixtures.Png());

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        var header = main.HeaderParts.Single();

        // The image belongs to the header, NOT the main document part. Getting this wrong yields a
        // relationship id resolved in the wrong scope - the file opens, showing nothing.
        Assert.Single(header.ImageParts);
        Assert.Empty(main.ImageParts);
        Assert.Single(header.Header!.Descendants<Drawing>());

        AssertValid(filled);
    }
```

- [ ] **Step 2: Two occurrences get distinct ids**

```csharp
    [Fact]
    public void ReplaceImage_GivesEveryOccurrenceItsOwnIdAndImagePart()
    {
        var docx = DocxFixtures.Build(
            DocxFixtures.P(DocxFixtures.R("{{logo}} first")),
            DocxFixtures.P(DocxFixtures.R("{{logo}} second")));

        var filled = DocxEditor.ReplaceImage(docx, "{{logo}}", ImageFixtures.Png());

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var ids = doc.MainDocumentPart!.Document!.Body!
            .Descendants<DW.DocProperties>().Select(p => p.Id!.Value).ToList();

        Assert.Equal(2, ids.Count);
        // Duplicate ids are what makes Word declare the file corrupt and offer to repair it.
        Assert.Equal(ids.Count, ids.Distinct().Count());

        AssertValid(filled);
    }
```

- [ ] **Step 3: A placeholder split across runs is still matched**

```csharp
    [Fact]
    public void ReplaceImage_MatchesAPlaceholderSplitAcrossRuns()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(
            DocxFixtures.R("start {{lo"),
            DocxFixtures.R("go}} end")));

        var filled = DocxEditor.ReplaceImage(docx, "{{logo}}", ImageFixtures.Png());

        var text = DocxEditor.ExtractText(filled);
        Assert.Contains("start ", text);
        Assert.Contains(" end", text);
        Assert.DoesNotContain("{{lo", text);

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Single(doc.MainDocumentPart!.Document!.Body!.Descendants<Drawing>());

        AssertValid(filled);
    }
```

- [ ] **Step 4: Content type follows the bytes, and sizing options behave**

```csharp
    [Fact]
    public void ReplaceImage_DeclaresTheContentTypeTheBytesActuallyAre()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("{{photo}}")));

        var filled = DocxEditor.ReplaceImage(docx, "{{photo}}", ImageFixtures.Jpeg());

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        Assert.Equal("image/jpeg", doc.MainDocumentPart!.ImageParts.Single().ContentType);
    }

    [Fact]
    public void ReplaceImage_ScalesTheOtherSideWhenOnlyOneIsGiven()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("{{logo}}")));

        var filled = DocxEditor.ReplaceImage(
            docx, "{{logo}}", ImageFixtures.Png(width: 2, height: 3), widthPoints: 10);

        using var ms = new MemoryStream(filled);
        using var doc = WordprocessingDocument.Open(ms, false);
        var extent = doc.MainDocumentPart!.Document!.Body!.Descendants<DW.Extent>().Single();

        Assert.Equal(127000L, extent.Cx!.Value);           // 10pt
        Assert.Equal(190500L, extent.Cy!.Value);           // 15pt, keeping 2:3
    }

    [Fact]
    public void ReplaceImage_RejectsAnUnsupportedFormatByName()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("{{logo}}")));

        var ex = Assert.Throws<DocumentConversionException>(
            () => DocxEditor.ReplaceImage(docx, "{{logo}}", ImageFixtures.Gif()));

        Assert.Contains("GIF", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ReplaceImage_RejectsBadArguments()
    {
        var docx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("{{logo}}")));
        var png = ImageFixtures.Png();

        Assert.Throws<ArgumentNullException>(() => DocxEditor.ReplaceImage(null!, "{{logo}}", png));
        Assert.Throws<ArgumentNullException>(() => DocxEditor.ReplaceImage(docx, null!, png));
        Assert.Throws<ArgumentNullException>(() => DocxEditor.ReplaceImage(docx, "{{logo}}", null!));
        Assert.Throws<ArgumentException>(() => DocxEditor.ReplaceImage(Array.Empty<byte>(), "{{logo}}", png));
        Assert.Throws<ArgumentException>(() => DocxEditor.ReplaceImage(docx, " ", png));
        Assert.Throws<ArgumentException>(() => DocxEditor.ReplaceImage(docx, "{{logo}}", Array.Empty<byte>()));
        Assert.Throws<ArgumentOutOfRangeException>(
            () => DocxEditor.ReplaceImage(docx, "{{logo}}", png, widthPoints: 0));
    }
```

- [ ] **Step 5: Run, full build, commit**

```bash
dotnet test tests/DocToolkit.Tests -c Release --filter FullyQualifiedName~DocxEditorReplaceImageTests
dotnet build DocToolkit.sln -c Release -warnaserror
dotnet test  DocToolkit.sln -c Release --no-build

git add tests/DocToolkit.Tests/DocxEditorReplaceImageTests.cs src/DocToolkit/DocxEditor.cs
git commit -m "test(core): prove the image-insertion traps are handled

Each of these produces a document that OPENS, so nothing already in the
suite would catch them: a header image must be owned by the header part or
its relationship id resolves in the wrong scope, two occurrences must get
distinct docPr ids or Word offers to repair the file, and the content type
must follow the magic bytes rather than any assertion.

Also covers a placeholder split across runs, proportional scaling from one
given dimension, an unsupported format named in the error, and argument
validation."
```

---

### Task 5: The `Stream` overload, registration, and docs

**Files:**
- Modify: `src/DocToolkit/DocxEditor.cs`
- Modify: `tests/DocToolkit.Tests/StreamOverloadTests.cs` — **the three name lists, the dispatcher, and `SourceBytesFor`**
- Modify: `tests/DocToolkit.Tests/AirGapGuardTests.cs`
- Modify: `src/DocToolkit/README.md`, `README.md`, `CLAUDE.md`

- [ ] **Step 1: Add `ReplaceImageAsync`**

Mirror `FillRowsAsync` exactly — note the real `StreamPipeline` signatures, which take message
strings:

```csharp
    /// <summary>
    /// Reads a .docx from <paramref name="source"/>, replaces every occurrence of
    /// <paramref name="placeholder"/> with <paramref name="image"/>, and writes the result to
    /// <paramref name="destination"/>. See <see cref="ReplaceImage"/> for what is matched and how it
    /// is sized.
    ///
    /// <paramref name="source"/> is read to its end and <paramref name="destination"/> is written;
    /// neither is disposed, closed or sought, and neither has to be seekable.
    /// </summary>
    public static async Task ReplaceImageAsync(
        Stream source, string placeholder, byte[] image, Stream destination,
        double? widthPoints = null, double? heightPoints = null, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(placeholder);
        ArgumentNullException.ThrowIfNull(image);
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        if (image.Length == 0) throw new ArgumentException("Image content was empty.", nameof(image));
        if (string.IsNullOrWhiteSpace(placeholder))
            throw new ArgumentException("Placeholder was blank.", nameof(placeholder));

        using var buffer = await StreamPipeline
            .DrainAsync(source, "DOCX content was empty.", nameof(source), "Failed to insert an image into the DOCX package.", ct)
            .ConfigureAwait(false);

        ReplaceImageCore(buffer, placeholder, image, widthPoints, heightPoints);

        await StreamPipeline
            .EmitAsync(buffer, destination, "Failed to insert an image into the DOCX package.", ct)
            .ConfigureAwait(false);
    }
```

- [ ] **Step 2: Register it in `StreamOverloadTests` — four edits, all required**

1. Add `"DocxEditor.ReplaceImageAsync"` to `DestinationWriterNames`, `SourceReaderNames` **and**
   `BufferedDestinationWriterNames`.
2. Add a dispatcher case:
   `"DocxEditor.ReplaceImageAsync" => DocxEditor.ReplaceImageAsync(source!, "{{logo}}", ImageFixtures.Png(), destination!, ct: ct),`
3. Add a source fixture, because the generic `Docx` has no `{{logo}}` and `ReplaceImage` throws
   without a match:
   `"DocxEditor.ReplaceImageAsync" => ImageDocx,` in `SourceBytesFor`.
4. Add that fixture:
   `private static readonly byte[] ImageDocx = DocxFixtures.Build(DocxFixtures.P(DocxFixtures.R("Logo: {{logo}}")));`

- [ ] **Step 3: Prove the registration took**

```bash
# before
git stash push tests/DocToolkit.Tests/StreamOverloadTests.cs
dotnet test tests/DocToolkit.Tests -c Release --filter FullyQualifiedName~StreamOverloadTests | grep -E 'Passed!'
git stash pop
# after
dotnet test tests/DocToolkit.Tests -c Release --filter FullyQualifiedName~StreamOverloadTests | grep -E 'Passed!'
```

Expected: the second count is **higher**. `FillRowsAsync` moved it 83 → 90; this should move it
similarly. An unchanged count means the name was not picked up and the overload is escaping the
entire suite — which `CLAUDE.md` calls out as the only way that happens.

- [ ] **Step 4: Add the air-gap case**

In `AirGapGuardTests`, alongside `DocxEditorFillRows_ContactsNothing`, add a case that puts the
loopback URL in the surrounding text and calls both `ReplaceImage` and `ReplaceImageAsync`, then
`await probe.AssertSilentAsync("DocxEditor.ReplaceImage / ReplaceImageAsync")`.

- [ ] **Step 5: Documentation**

- `src/DocToolkit/README.md` — a **Images** section after *Repeating table rows*, showing the
  default-size call and the `widthPoints` call, and stating PNG/JPEG only and that size comes from
  the header at 96 DPI.
- `README.md` — one line in the usage block.
- `CLAUDE.md` — under *Traps*, record that image parts must be added to the part owning the
  paragraph and that `wp:docPr` ids must be unique, both being failures that open cleanly.
- **Update the counts** in both READMEs and `CLAUDE.md` to whatever the suite actually reports.
  Measure, do not estimate: the A4 pass documented 209 core before the air-gap test pushed it to 210.

- [ ] **Step 6: Final verification and PR**

```bash
dotnet build DocToolkit.sln -c Release -warnaserror
dotnet test  DocToolkit.sln -c Release --no-build
dotnet restore src/DocToolkit/DocToolkit.csproj --locked-mode    # csproj untouched, but cheap to confirm

git add -A
git commit -m "feat(core): add DocxEditor.ReplaceImageAsync, and document images"
git push -u origin feat/image-placeholders
gh pr create --base main --title "feat(core): image placeholders for DOCX templates" --body "Implements backlog A5, per docs/2026-08-03-image-placeholders-design.md. Completes the template story: scalars, repeating rows, and now images.

Dimensions come from PNG and JPEG headers in managed code - no image-decoding dependency, because SixLabors.ImageSharp's later majors moved onto the same revenue-gated licence SixLabors.Fonts is pinned to avoid.

Covers the traps that produce a document which OPENS: image parts owned by the wrong container, duplicate wp:docPr ids, and a content type that disagrees with the magic bytes."
```

---

## Notes for the reviewer

- **Tasks 1 and 2 are pure functions** and deliberately separate from any document. The parsers and
  the EMU arithmetic are the parts most likely to be subtly wrong, and the arithmetic in particular
  has no schema check behind it — a wrong size is a valid document.
- **Task 4 is the one that matters most.** Every case in it produces a file that opens; none would
  be caught by an existing test.
- **Task 5 Step 3 is the step most likely to be skipped**, and the quietest if it is.
- **Merging no longer publishes**, so this can land as several small PRs if that reads better than
  one branch. The commits are already split along those lines.
