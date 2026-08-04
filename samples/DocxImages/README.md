# DOCX images

Replacing a placeholder with a picture — a logo, a signature, a QR code — using `ReplaceImage`.

```bash
dotnet run --project samples/DocxImages
```

Prints the image size, how much the document grew, and whether the placeholder is gone.

## The non-obvious part

**Sizing is in points, and you usually want to give only one dimension.** Supply `widthPoints` and
the height scales to preserve the aspect ratio. Supply neither and the image's own header decides,
read at 96 DPI.

**The format is decided by the image's magic bytes, never by a filename.** PNG and JPEG are
supported. A file that claims to be a PNG while holding JPEG bytes renders as a blank frame in
Word, silently — so the bytes are what count.

The logo here is a 137-byte PNG inlined as base64, so this sample carries no binary file. Any real
PNG or JPEG works identically.
