using System.IO.Compression;

namespace DocToolkit.Tests;

/// <summary>
/// Real image bytes, without an image library.
///
/// The whole point of the design is that this repo takes no image-decoding dependency — see
/// <c>docs/2026-08-03-image-placeholders-design.md</c> — so the tests cannot quietly take one
/// either. The PNG is hand-built; the JPEG is lifted from a fixture already committed here.
/// </summary>
internal static class ImageFixtures
{
    /// <summary>
    /// A valid PNG of <paramref name="width"/> x <paramref name="height"/>.
    ///
    /// Used at deliberately odd sizes so an intrinsic-sizing assertion cannot pass by coincidence:
    /// 2 x 3 would still look right if width and height were swapped, which is why one test uses
    /// 300 x 260 instead.
    /// </summary>
    public static byte[] Png(int width = 2, int height = 3)
    {
        var ihdr = new byte[13];
        WriteBigEndian(ihdr, 0, (uint)width);
        WriteBigEndian(ihdr, 4, (uint)height);
        ihdr[8] = 8;   // bit depth
        ihdr[9] = 6;   // colour type: RGBA
        // 10..12 stay zero: deflate compression, adaptive filtering, no interlacing.

        // One filter byte per row, then RGBA per pixel. All zeroes is a valid black image.
        var raw = new byte[height * (1 + (width * 4))];
        using var deflated = new MemoryStream();
        using (var zlib = new ZLibStream(deflated, CompressionLevel.Fastest, leaveOpen: true))
        {
            zlib.Write(raw, 0, raw.Length);
        }

        using var png = new MemoryStream();
        png.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });
        WriteChunk(png, "IHDR", ihdr);
        WriteChunk(png, "IDAT", deflated.ToArray());
        WriteChunk(png, "IEND", Array.Empty<byte>());
        return png.ToArray();
    }

    /// <summary>
    /// Real JPEG bytes, 256 x 144, taken from the PPTX fixture already in this repo — so no new
    /// binary file is committed and no <c>.gitattributes</c> question arises for it.
    /// </summary>
    public static byte[] Jpeg()
    {
        var pptx = Path.Combine(AppContext.BaseDirectory, "assets", "sample.pptx");
        using var zip = ZipFile.OpenRead(pptx);
        var entry = zip.GetEntry("docProps/thumbnail.jpeg")
                    ?? throw new InvalidOperationException(
                        "sample.pptx no longer contains docProps/thumbnail.jpeg - the JPEG fixture needs a new source.");

        using var stream = entry.Open();
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }

    /// <summary>A GIF header — enough to be detected, and rejected, as an unsupported format.</summary>
    public static byte[] Gif() =>
        new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00, 0x01, 0x00 };

    /// <summary>
    /// A 1x1 24-bit BMP built by hand: 14-byte file header, 40-byte BITMAPINFOHEADER, one padded
    /// pixel. BMP rather than PNG so no compressor is involved. Used by <c>LoopbackProbe</c> and by
    /// data-URI fixtures that need something a real image decoder would accept.
    /// </summary>
    public static byte[] Bmp()
    {
        var bmp = new byte[58];

        // Disposed, though BinaryWriter writes through to the stream per call rather than
        // accumulating like StreamWriter, so nothing is lost without it today. Not worth
        // depending on that: a reader should not have to know which writer buffers.
        using var w = new BinaryWriter(new MemoryStream(bmp));
        w.Write((byte)'B'); w.Write((byte)'M');
        w.Write(58);            // file size
        w.Write(0);             // reserved
        w.Write(54);            // pixel data offset
        w.Write(40);            // BITMAPINFOHEADER size
        w.Write(1);             // width
        w.Write(1);             // height
        w.Write((short)1);      // planes
        w.Write((short)24);     // bits per pixel
        w.Write(0);             // BI_RGB, uncompressed
        w.Write(4);             // image byte size
        w.Write(2835);          // horizontal pixels per metre
        w.Write(2835);          // vertical pixels per metre
        w.Write(0);             // palette colours used
        w.Write(0);             // important colours
        w.Write(new byte[] { 0x00, 0x00, 0xFF, 0x00 }); // one red pixel + row padding
        return bmp;
    }

    private static void WriteChunk(Stream target, string type, byte[] data)
    {
        var length = new byte[4];
        WriteBigEndian(length, 0, (uint)data.Length);
        target.Write(length);

        // The CRC covers the type and the data together, not the length.
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
