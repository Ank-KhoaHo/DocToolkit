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
/// Reads an image's format and pixel dimensions straight from its header, and converts a requested
/// size into EMUs.
///
/// Every image in a .docx carries an explicit size, so something must supply the dimensions. The
/// obvious candidate — SixLabors.ImageSharp — moved its later majors onto the same revenue-gated
/// Six Labors Split License that <c>SixLabors.Fonts</c> is pinned at <c>[1.0.x]</c> to avoid, so
/// taking it would mean a second permanent CI guard against a licence this package exists to stay
/// clear of. PNG and JPEG headers are cheap to read directly, so they are read directly.
///
/// The format is decided by MAGIC BYTES, never by a filename or a caller's assertion: an image part
/// declaring <c>image/png</c> while holding JPEG bytes renders as a blank frame in Word, with no
/// error anywhere.
/// </summary>
internal static class ImageInspector
{
    private const long EmuPerPoint = 12700;

    /// <summary>Word assumes 96 DPI for an image carrying no DPI metadata, so this does too.</summary>
    private const long EmuPerPixelAt96Dpi = 9525;

    /// <summary>
    /// The largest extent a drawing may carry, in EMU — about 2,348 inches per side.
    ///
    /// ECMA-376 types an extent as <c>ST_PositiveCoordinate</c>, whose documented maximum is
    /// 27,273,042,316,900 EMU (exactly <see cref="int.MaxValue"/> POINTS). <b>DocumentFormat.OpenXml
    /// does not enforce that value.</b> Measured against OpenXmlValidator 3.5.1, on both
    /// <c>wp:extent</c> and <c>a:ext</c>: 2,147,483,647 validates, 2,147,483,648 fails with
    /// "The MaxInclusive constraint failed. The value must be less than or equal to 2147483647",
    /// and the spec's own maximum fails the same way. The SDK caps at <see cref="int.MaxValue"/>
    /// EMU, four orders of magnitude lower.
    ///
    /// The SDK's number wins here, deliberately: OpenXmlValidator is this library's definition of
    /// "schema-valid" — it is what every AssertValid in the test suite calls — so a bound the
    /// validator rejects would let this guard pass documents the repo considers invalid. Using the
    /// spec value first was measured to do exactly that: widthPoints of 1,000,000 returned a
    /// document with four validation errors and threw nothing.
    /// </summary>
    private const long MaxCoordinateEmu = int.MaxValue;

    private static ReadOnlySpan<byte> PngSignature =>
        new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

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
    /// The size to write into the drawing, in EMUs.
    ///
    /// Neither dimension given: the image's intrinsic pixel size at 96 DPI. One given: the other
    /// scales to preserve the aspect ratio. Both given: exactly those, distortion accepted as the
    /// caller's choice.
    ///
    /// Nothing downstream validates these numbers — a factor-of-ten slip yields a perfectly valid
    /// document containing a wrongly sized image, which no schema check would ever flag.
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

        var resolved =
            widthPoints is null && heightPoints is null
                ? (intrinsicWidth, intrinsicHeight)
            : widthPoints is not null && heightPoints is not null
                ? ((long)(widthPoints.Value * EmuPerPoint), (long)(heightPoints.Value * EmuPerPoint))
            : widthPoints is not null
                ? ScaleFromWidth(widthPoints.Value)
                : ScaleFromHeight(heightPoints!.Value);

        // Checked AFTER resolving, not on the arguments, because the out-of-range value is often one
        // the caller never passed: a legal widthPoints on a tall image derives a larger height, so
        // the derived side can exceed the limit while the argument sits comfortably inside it.
        //
        // The INTRINSIC path cannot overflow and is not why this exists: WidthPx is an int, so the
        // largest size any header can express is int.MaxValue * 9525 = 20,454,781,737,675 EMU, about
        // three quarters of the limit. Measured, after a comment here first claimed the opposite.
        //
        // Without this the (long) cast saturates to long.MinValue and the document is written with a
        // NEGATIVE extent — schema-invalid, but produced with no exception, so the caller gets a
        // corrupt file and no signal. Nothing downstream validates these numbers.
        if (resolved.Item1 is <= 0 or > MaxCoordinateEmu || resolved.Item2 is <= 0 or > MaxCoordinateEmu)
        {
            // Which argument to blame. When only one side was given, blame it even if the DERIVED
            // side is what overflowed — the caller chose the one number that caused it. When BOTH
            // were given the sides are independent, so name the one actually out of range; blaming
            // widthPoints for a bad heightPoints would report a value the caller can see is fine.
            var widthBad = resolved.Item1 is <= 0 or > MaxCoordinateEmu;
            var (name, value) =
                widthPoints is not null && (heightPoints is null || widthBad)
                    ? (nameof(widthPoints), widthPoints)
                : heightPoints is not null
                    ? (nameof(heightPoints), heightPoints)
                    : (nameof(info), null);

            throw new ArgumentOutOfRangeException(name, value,
                $"The resulting image size is outside what OOXML can represent " +
                $"({MaxCoordinateEmu} EMU, or {MaxCoordinateEmu / EmuPerPoint} points, per side).");
        }

        return resolved;

        (long, long) ScaleFromWidth(double points)
        {
            var width = (long)(points * EmuPerPoint);
            return (width, (long)(width * (double)intrinsicHeight / intrinsicWidth));
        }

        (long, long) ScaleFromHeight(double points)
        {
            var height = (long)(points * EmuPerPoint);
            return ((long)(height * (double)intrinsicWidth / intrinsicHeight), height);
        }
    }

    /// <summary>
    /// PNG is fixed-layout: the IHDR chunk is always first, and its width and height are big-endian
    /// <c>uint32</c> at bytes 16 and 20.
    /// </summary>
    private static ImageInfo InspectPng(byte[] image)
    {
        if (image.Length < 24)
        {
            throw new DocumentConversionException(
                "PNG image is truncated: it has no complete IHDR chunk. The image bytes are "
                + "incomplete — check the file was fully read or uploaded before it was passed in.");
        }

        var width = (int)BinaryPrimitives.ReadUInt32BigEndian(image.AsSpan(16, 4));
        var height = (int)BinaryPrimitives.ReadUInt32BigEndian(image.AsSpan(20, 4));

        if (width <= 0 || height <= 0)
            throw new DocumentConversionException($"PNG image reports a nonsensical size ({width}x{height}).");

        return new ImageInfo(ImageFormat.Png, width, height);
    }

    /// <summary>
    /// JPEG has no fixed offset: walk the segment markers to a Start-Of-Frame, where height then
    /// width follow the one-byte sample precision. Every SOF marker in C0..CF carries the size
    /// except C4 (Huffman table), C8 (reserved) and CC (arithmetic conditioning).
    /// </summary>
    private static ImageInfo InspectJpeg(byte[] image)
    {
        var i = 2;
        while (i < image.Length - 1)
        {
            if (image[i] != 0xFF)
            {
                i++;
                continue;
            }

            var marker = image[i + 1];

            if (marker is >= 0xC0 and <= 0xCF && marker is not (0xC4 or 0xC8 or 0xCC))
            {
                if (i + 9 > image.Length)
                {
                    throw new DocumentConversionException(
                        "JPEG image is truncated inside its frame header. The image bytes are "
                        + "incomplete — check the file was fully read or uploaded before it was passed in.");
                }

                var height = BinaryPrimitives.ReadUInt16BigEndian(image.AsSpan(i + 5, 2));
                var width = BinaryPrimitives.ReadUInt16BigEndian(image.AsSpan(i + 7, 2));

                if (width == 0 || height == 0)
                    throw new DocumentConversionException($"JPEG image reports a nonsensical size ({width}x{height}).");

                return new ImageInfo(ImageFormat.Jpeg, width, height);
            }

            // Standalone markers: start/end of image, and the restart markers, carry no length.
            if (marker is 0xD8 or 0xD9 || marker is >= 0xD0 and <= 0xD7)
            {
                i += 2;
                continue;
            }

            if (i + 4 > image.Length)
            {
                throw new DocumentConversionException(
                    "JPEG image is truncated inside a segment header. The image bytes are "
                    + "incomplete — check the file was fully read or uploaded before it was passed in.");
            }

            i += 2 + BinaryPrimitives.ReadUInt16BigEndian(image.AsSpan(i + 2, 2));
        }

        throw new DocumentConversionException(
            "JPEG image has no Start-Of-Frame segment, so its size cannot be determined. This "
            + "usually means the image bytes are truncated — check the file was fully read or "
            + "uploaded before it was passed in.");
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
