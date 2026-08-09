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

**An image the HTML only *points at* is a different problem, and the sample ends on it.** An
`<img src="https://…">` is a request that your server fetch a URL somebody else chose, so remote
downloads are off by default and the document is produced with the image absent rather than
failing.

When you opt in with `RemoteImageOptions`, the allow-list is the part that matters. A host that is
not on it is refused **before any connection is attempted** — no DNS lookup, no socket — which is
why this sample can demonstrate the guard without touching the network, and why the two documents
it prints come out at exactly the same size.
