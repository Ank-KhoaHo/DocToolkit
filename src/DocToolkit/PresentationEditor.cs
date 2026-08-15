using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace DocToolkit;

/// <summary>Opens and edits PowerPoint (.pptx) presentations.</summary>
public static class PresentationEditor
{
    /// <summary>
    /// Creates a deck from <paramref name="slides"/>, one slide each.
    ///
    /// This exists for content that comes from data rather than from an existing file: there is no
    /// template to edit, so <see cref="ReplaceText"/> cannot help, and the same slides produce the
    /// same CONTENT on every machine — nothing here consults the current culture. Not the same
    /// BYTES: the OpenXml SDK mints fresh relationship ids per package, so two calls with identical
    /// slides in the same process differ. Do not build a cache key, a content hash or a golden-file
    /// test on the bytes.
    ///
    /// An empty sequence is valid and produces a valid deck with no slides.
    /// </summary>
    /// <param name="slides">The slides, in deck order.</param>
    /// <exception cref="ArgumentNullException"><paramref name="slides"/> is null.</exception>
    /// <exception cref="ArgumentException">An element of <paramref name="slides"/> is null.</exception>
    /// <exception cref="DocumentConversionException">The deck could not be built.</exception>
    /// <example>
    /// <code source="../../tests/DocToolkit.Tests/DocumentationExamples.cs" region="PresentationCreate"/>
    /// </example>
    public static byte[] Create(IEnumerable<PptxSlide> slides)
    {
        var materialised = ValidateSlides(slides);
        using var ms = PptxDocumentWriter.Write(materialised);
        return ms.ToArray();
    }

    /// <summary>
    /// Builds a deck from <paramref name="slides"/> and writes it to
    /// <paramref name="destination"/>. See <see cref="Create"/> for the slide semantics.
    ///
    /// <paramref name="destination"/> is <b>written</b>, from its current position, and is
    /// <b>not</b> disposed, closed or sought — it belongs to the caller, and may be write-only and
    /// forward-only, such as an HTTP response body.
    /// </summary>
    /// <param name="slides">The slides, in deck order.</param>
    /// <param name="destination">The stream the deck is written to.</param>
    /// <param name="ct">Cancels the build and the write to <paramref name="destination"/>.</param>
    /// <exception cref="ArgumentNullException"><paramref name="slides"/> or <paramref name="destination"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// An element of <paramref name="slides"/> is null, or <paramref name="destination"/> is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The deck could not be built or written.</exception>
    public static async Task CreateAsync(
        IEnumerable<PptxSlide> slides, Stream destination, CancellationToken ct = default)
    {
        var materialised = ValidateSlides(slides);
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var ms = PptxDocumentWriter.Write(materialised);
        await StreamPipeline.EmitAsync(ms, destination, "Failed to create PPTX.", ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Builds a deck from <paramref name="slides"/> and writes it to <paramref name="outputPath"/>.
    /// See <see cref="Create"/> for the slide semantics.
    ///
    /// Named <c>CreateToFileAsync</c> rather than a third <c>CreateAsync</c> overload, matching
    /// <see cref="WorkbookEditor.CreateToFileAsync(string, System.Collections.Generic.IEnumerable{System.Collections.Generic.IEnumerable{object}}, string, System.Threading.CancellationToken)"/>:
    /// the distinct name keeps which kind of destination a call writes to visible at the call site,
    /// rather than resting on the argument type alone.
    ///
    /// The deck is built completely before the output is opened. That ordering is what stops a
    /// failed build truncating a file that was already there, and it is pinned by
    /// <c>FilePathOverloadTests</c> rather than left as a comment — it survives only as long as
    /// nobody rewrites this into a streaming write.
    /// </summary>
    /// <param name="slides">The slides, in deck order.</param>
    /// <param name="outputPath">Where to write the deck. Overwritten if it exists.</param>
    /// <param name="ct">Cancels the write to <paramref name="outputPath"/>.</param>
    /// <exception cref="ArgumentNullException"><paramref name="slides"/> or <paramref name="outputPath"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="outputPath"/> is blank, or an element of <paramref name="slides"/> is null.
    /// </exception>
    /// <exception cref="DirectoryNotFoundException"><paramref name="outputPath"/>'s directory does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The deck could not be built.</exception>
    public static async Task CreateToFileAsync(
        IEnumerable<PptxSlide> slides, string outputPath, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var bytes = Create(slides);
        await File.WriteAllBytesAsync(outputPath, bytes, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Materialises and null-checks up front, so a null slide surfaces as the
    /// <see cref="ArgumentException"/> it is rather than as a <see cref="NullReferenceException"/>
    /// wrapped in a conversion failure. Mirrors <c>DocxEditor.ValidateBlocks</c>.
    /// </summary>
    private static IReadOnlyList<PptxSlide> ValidateSlides(IEnumerable<PptxSlide> slides)
    {
        ArgumentNullException.ThrowIfNull(slides);

        return slides
            .Select((slide, index) => slide
                ?? throw new ArgumentException($"Slide {index + 1} was null.", nameof(slides)))
            .ToList();
    }

    /// <summary>Number of slides in the deck, as counted from the deck's slide list.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static int SlideCount(byte[] pptx)
    {
        using var ms = OpenForWrite(pptx);
        return SlideCountCore(ms);
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/> and returns its slide count, counted from the
    /// deck's slide list. <paramref name="source"/> is <b>read</b> to its end and is neither
    /// disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static async Task<int> SlideCountAsync(Stream source, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var ms = await StreamPipeline
            .DrainAsync(source, "Presentation content was empty.", nameof(source), "Failed to read PPTX.", ct)
            .ConfigureAwait(false);

        return SlideCountCore(ms);
    }

    private static int SlideCountCore(MemoryStream ms)
    {
        try
        {
            using var doc = OpenDocument(ms, false);
            return SlidesInDeckOrder(PresentationPartOf(doc)).Count();
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read PPTX.", ex);
        }
    }

    /// <summary>
    /// All text found on every slide, one entry per text-bearing body, in the order the deck is
    /// presented rather than the order the slide parts happen to be related.
    ///
    /// A "text-bearing body" is any element holding &lt;a:p&gt; paragraphs: an ordinary shape's
    /// &lt;p:txBody&gt;, a shape nested in a group, and a table cell's &lt;a:txBody&gt; alike.
    /// Paragraphs within one body are joined with newlines. This is deliberately the same walk
    /// <see cref="ReplaceText"/> performs, so anything this reports is something that can be
    /// replaced and vice versa. Speaker notes and slide masters/layouts are not included.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static IReadOnlyList<string> ExtractText(byte[] pptx)
    {
        using var ms = OpenForWrite(pptx);
        return ExtractTextCore(ms);
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/> and returns all text found on every slide, one
    /// entry per text-bearing body, in deck order — see <see cref="ExtractText(byte[])"/> for
    /// exactly what counts as a text-bearing body. <paramref name="source"/> is <b>read</b> to its
    /// end and is neither disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static async Task<IReadOnlyList<string>> ExtractTextAsync(Stream source, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var ms = await StreamPipeline
            .DrainAsync(source, "Presentation content was empty.", nameof(source), "Failed to read PPTX.", ct)
            .ConfigureAwait(false);

        return ExtractTextCore(ms);
    }

    private static IReadOnlyList<string> ExtractTextCore(MemoryStream ms)
    {
        try
        {
            using var doc = OpenDocument(ms, false);

            var results = new List<string>();
            foreach (var slide in SlidesInDeckOrder(PresentationPartOf(doc))
                         .Select(p => p.Slide).Where(s => s is not null).Select(s => s!))
            {

                // PowerPoint stores shape text as a:t runs under a:p paragraphs - Wordprocessing's
                // w:t is the DOCX equivalent, not what PPTX uses. Grouping by the paragraph's
                // parent yields one entry per text body while preserving document order.
                foreach (var body in slide.Descendants<A.Paragraph>().GroupBy(p => p.Parent))
                {
                    var bodyText = string.Join(
                        "\n",
                        body.Select(p => string.Concat(p.Descendants<A.Text>().Select(t => t.Text))));

                    if (bodyText.Length > 0) results.Add(bodyText);
                }
            }

            return results;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read PPTX.", ex);
        }
    }

    /// <summary>
    /// Replaces every key with its value across all slide text, returning updated bytes.
    ///
    /// PowerPoint routinely splits a single visible word across several &lt;a:t&gt; runs
    /// (spell-check state, formatting changes), so a naive per-run replace misses any placeholder
    /// that straddles a run boundary. Substitution therefore happens against the concatenated text
    /// of each paragraph, but the result is spliced back into only the runs the match actually
    /// overlaps: runs outside a match keep their text and their formatting untouched. When a
    /// placeholder does straddle runs, the replacement value is written into the run holding its
    /// first character and so inherits that run's formatting.
    ///
    /// Keys are matched in a single left-to-right pass and the longest key wins at any given
    /// offset, so a substituted value is never rescanned for further placeholders. Slides are
    /// visited in deck order; speaker notes and slide masters/layouts are not touched.
    /// </summary>
    /// <exception cref="ArgumentNullException">Either argument is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or edited.</exception>
    public static byte[] ReplaceText(byte[] pptx, IReadOnlyDictionary<string, string> replacements)
    {
        ArgumentNullException.ThrowIfNull(replacements);

        using var ms = OpenForWrite(pptx);
        ReplaceTextCore(ms, replacements);
        return ms.ToArray();
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/>, replaces every key with its value across all
    /// slide text, and writes the result to <paramref name="destination"/> — see
    /// <see cref="ReplaceText"/> for exactly what counts as a match and how formatting survives it.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought, and neither has to be seekable.
    /// </summary>
    /// <param name="source">The stream the .pptx package is read from.</param>
    /// <param name="replacements">Each key is replaced by its value, longest key wins per match.</param>
    /// <param name="destination">The stream the edited .pptx package is written to.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="destination"/>
    /// is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or edited.</exception>
    public static async Task ReplaceTextAsync(
        Stream source, IReadOnlyDictionary<string, string> replacements, Stream destination,
        CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(replacements);
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var ms = await StreamPipeline
            .DrainAsync(source, "Presentation content was empty.", nameof(source), "Failed to edit PPTX.", ct)
            .ConfigureAwait(false);

        ReplaceTextCore(ms, replacements);

        await StreamPipeline.EmitAsync(ms, destination, "Failed to edit PPTX.", ct).ConfigureAwait(false);
    }

    private static void ReplaceTextCore(MemoryStream ms, IReadOnlyDictionary<string, string> replacements)
    {
        try
        {
            using (var doc = OpenDocument(ms, true))
            {
                foreach (var slide in SlidesInDeckOrder(PresentationPartOf(doc))
                             .Select(p => p.Slide).Where(s => s is not null).Select(s => s!))
                {

                    foreach (var paragraph in slide.Descendants<A.Paragraph>())
                    {
                        var texts = paragraph.Descendants<A.Text>().ToList();
                        if (texts.Count == 0) continue;

                        RunTextSplicer.Apply(
                            texts, static t => t.Text, static (t, v) => t.Text = v, replacements);
                    }

                    slide.Save();
                }
            }
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to edit PPTX.", ex);
        }
    }

    /// <summary>
    /// Replaces every shape whose text is exactly <paramref name="placeholder"/> with
    /// <paramref name="image"/>, which is scaled to fit inside that shape's box and centred there.
    ///
    /// Position and size come from the template, so there is nothing to pass: a designer draws a
    /// box in PowerPoint where the image belongs and the image lands there. This deliberately does
    /// not mirror <see cref="DocxEditor.ReplaceImage"/>'s size arguments — a DOCX image is inline
    /// in the text flow and needs a size, a PPTX picture is a positioned shape and already has one.
    ///
    /// The shape's text must be nothing but the placeholder. The unit replaced is the whole shape,
    /// so a shape reading <c>Chart: {{chart}} (Q3)</c> would lose the words around the placeholder
    /// — silently, and with a schema-valid result. That is refused instead.
    ///
    /// PNG and JPEG only, detected from magic bytes rather than any filename.
    /// </summary>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="pptx"/> or <paramref name="image"/> is empty, or <paramref name="placeholder"/>
    /// is blank.
    /// </exception>
    /// <exception cref="DocumentConversionException">
    /// The placeholder appears nowhere, appears only inside a grouped shape, a matched shape holds
    /// other text, a matched shape has no explicit position, the image is neither PNG nor JPEG,
    /// or the package could not be edited.
    /// </exception>
    public static byte[] ReplaceImage(byte[] pptx, string placeholder, byte[] image)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(placeholder);
        ArgumentNullException.ThrowIfNull(image);
        if (image.Length == 0)
            throw new ArgumentException("Image content was empty.", nameof(image));

        using var ms = OpenForWrite(pptx);
        ReplaceImageCore(ms, placeholder, image);
        return ms.ToArray();
    }

    private static void ReplaceImageCore(MemoryStream ms, string placeholder, byte[] image)
    {
        // Inspect before opening the package: an unsupported format is the caller's mistake and
        // should not depend on whether the deck happens to be readable.
        var info = ImageInspector.Inspect(image);
        var (imageCx, imageCy) = ImageInspector.Resolve(info, null, null);

        var replaced = 0;
        var placeholderOnlyInsideAGroup = false;

        try
        {
            using (var doc = OpenDocument(ms, true))
            {
                foreach (var slidePart in SlidesInDeckOrder(PresentationPartOf(doc)))
                {
                    var tree = slidePart.Slide?.CommonSlideData?.ShapeTree;
                    if (tree is null) continue;

                    // Direct children only. A shape inside a group carries coordinates in the
                    // group's own space, so placing a picture there from slide-space numbers would
                    // put it somewhere unrelated. A placeholder that exists only inside a group is
                    // therefore never matched by this loop; it is detected separately below so the
                    // refusal at the end of this method can name that as the reason instead of
                    // falsely reporting the placeholder as absent.
                    foreach (var shape in tree.Elements<P.Shape>().ToList())
                    {
                        var text = string.Concat(shape.Descendants<A.Text>().Select(t => t.Text));
                        if (!text.Contains(placeholder, StringComparison.Ordinal)) continue;

                        if (text.Trim() != placeholder)
                        {
                            throw new DocumentConversionException(
                                $"The shape holding '{placeholder}' also holds other text "
                                + $"(\"{text}\"). ReplaceImage swaps the whole shape, so its text "
                                + "must be only the placeholder — anything else would be silently "
                                + "discarded. Put the placeholder in a box of its own.");
                        }

                        var xfrm = shape.ShapeProperties?.Transform2D;
                        if (xfrm?.Offset?.X is null || xfrm.Offset.Y is null
                            || xfrm.Extents?.Cx is null || xfrm.Extents.Cy is null)
                        {
                            throw new DocumentConversionException(
                                $"The shape holding '{placeholder}' has no position of its own, so "
                                + "there is nowhere to put the image. Draw a text box rather than "
                                + "using an unpositioned layout placeholder.");
                        }

                        var (x, y, cx, cy) = PptxPictureFactory.Fit(
                            xfrm.Offset.X!.Value, xfrm.Offset.Y!.Value,
                            xfrm.Extents.Cx!.Value, xfrm.Extents.Cy!.Value,
                            imageCx, imageCy);

                        // The image part belongs to the slide that owns the shape. On the
                        // presentation part the relationship resolves in the wrong scope and
                        // PowerPoint renders nothing at all.
                        //
                        // Not slidePart.AddImagePart(ImagePartType...): in this SDK version that
                        // extension takes a PartTypeInfo (or a raw content-type string) rather than
                        // ImagePartType, so it does not resolve against the enum. AddNewPart<T> with
                        // an explicit content type is the same pattern DocxEditor's own AddImagePart
                        // helper already uses.
                        var imagePart = slidePart.AddNewPart<ImagePart>(info.ContentType);
                        using (var content = new MemoryStream(image, writable: false))
                        {
                            imagePart.FeedData(content);
                        }

                        // The id is load-bearing: this is a 1:1 swap of the replaced shape, so the
                        // plan reuses its own p:cNvPr/@id rather than minting a new one — minting
                        // one here could collide with an id already in use elsewhere in the deck.
                        // A missing id means the input is malformed enough that there is no safe
                        // id to give the picture, so this refuses rather than guessing. The name is
                        // purely cosmetic (PowerPoint's selection pane label), so a missing one gets
                        // a sensible fallback instead of failing the whole replacement over it.
                        var id = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id
                                 ?? throw new DocumentConversionException(
                                     $"The shape holding '{placeholder}' has no p:cNvPr/@id, so "
                                     + "there is no id to reuse for the replacement picture and it "
                                     + "cannot be replaced.");
                        var name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name
                                   ?? new StringValue("Picture");

                        shape.Parent!.ReplaceChild(
                            PptxPictureFactory.Picture(
                                id!, name!, slidePart.GetIdOfPart(imagePart), x, y, cx, cy),
                            shape);

                        replaced++;
                    }

                    if (replaced == 0 && !placeholderOnlyInsideAGroup)
                    {
                        placeholderOnlyInsideAGroup = tree.Descendants<P.GroupShape>()
                            .SelectMany(group => group.Descendants<P.Shape>())
                            .Any(shape => string.Concat(shape.Descendants<A.Text>().Select(t => t.Text))
                                .Contains(placeholder, StringComparison.Ordinal));
                    }

                    slidePart.Slide!.Save();
                }
            }
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to edit PPTX.", ex);
        }

        if (replaced == 0)
        {
            if (placeholderOnlyInsideAGroup)
            {
                throw new DocumentConversionException(
                    $"'{placeholder}' appears only inside a grouped shape (p:grpSp). ReplaceImage "
                    + "does not look inside groups: a shape inside one carries coordinates in the "
                    + "group's own space, not the slide's, so there is no slide-space position to "
                    + "give the replacement picture. Ungroup the shape, or draw the placeholder box "
                    + "outside any group.");
            }

            throw new DocumentConversionException(
                $"'{placeholder}' does not appear in any shape, so nothing was replaced.");
        }
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/>, replaces every shape whose text is exactly
    /// <paramref name="placeholder"/> with <paramref name="image"/>, and writes the result to
    /// <paramref name="destination"/> — see <see cref="ReplaceImage"/> for exactly what counts as a
    /// match and how the image is fit into the matched shape's box.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought, and neither has to be seekable.
    /// </summary>
    /// <param name="source">The stream the .pptx package is read from.</param>
    /// <param name="placeholder">The placeholder text a shape must hold, and hold only.</param>
    /// <param name="image">PNG or JPEG bytes. The format is decided by the bytes, never a filename.</param>
    /// <param name="destination">The stream the edited .pptx package is written to.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="placeholder"/> is blank, <paramref name="image"/> is empty,
    /// <paramref name="source"/> is not readable or held no bytes, or <paramref name="destination"/>
    /// is not writable.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The placeholder appears nowhere, appears only inside a grouped shape, a matched shape holds
    /// other text, a matched shape has no explicit position, the image is neither PNG nor JPEG,
    /// or the package could not be edited.
    /// </exception>
    public static async Task ReplaceImageAsync(
        Stream source, string placeholder, byte[] image, Stream destination,
        CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(placeholder);
        ArgumentNullException.ThrowIfNull(image);
        if (image.Length == 0)
            throw new ArgumentException("Image content was empty.", nameof(image));
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var ms = await StreamPipeline
            .DrainAsync(source, "Presentation content was empty.", nameof(source), "Failed to edit PPTX.", ct)
            .ConfigureAwait(false);

        ReplaceImageCore(ms, placeholder, image);

        await StreamPipeline.EmitAsync(ms, destination, "Failed to edit PPTX.", ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Slide parts in the order the deck presents them.
    ///
    /// <c>PresentationPart.SlideParts</c> is part-relationship order, which has nothing to do with
    /// slide order - reorder a deck in PowerPoint and the two diverge completely. The authoritative
    /// order is <c>p:sldIdLst</c>, so resolve each <c>p:sldId</c> through its relationship id.
    /// </summary>
    private static IEnumerable<SlidePart> SlidesInDeckOrder(PresentationPart presentationPart)
    {
        var slideIdList = presentationPart.Presentation?.SlideIdList;
        if (slideIdList is null)
        {
            // No slide list at all: nothing better to go on than the relationship order.
            foreach (var slidePart in presentationPart.SlideParts) yield return slidePart;
            yield break;
        }

        foreach (var relationshipId in slideIdList.Elements<P.SlideId>()
                     .Select(id => id.RelationshipId?.Value)
                     .Where(id => !string.IsNullOrEmpty(id))
                     .Select(id => id!))
        {

            OpenXmlPart part;
            try { part = presentationPart.GetPartById(relationshipId); }
            catch (ArgumentOutOfRangeException) { continue; }

            if (part is SlidePart slidePart) yield return slidePart;
        }
    }

    private static PresentationPart PresentationPartOf(PresentationDocument doc)
        => doc.PresentationPart
           ?? throw new DocumentConversionException("Presentation has no presentation part.");

    private static PresentationDocument OpenDocument(MemoryStream ms, bool isEditable)
    {
        try
        {
            return PresentationDocument.Open(ms, isEditable);
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException("Failed to open PPTX.", ex);
        }
    }

    private static MemoryStream OpenForWrite(byte[] pptx)
    {
        ArgumentNullException.ThrowIfNull(pptx);
        if (pptx.Length == 0)
            throw new ArgumentException("Presentation content was empty.", nameof(pptx));

        var ms = new MemoryStream();
        ms.Write(pptx, 0, pptx.Length);
        ms.Position = 0;
        return ms;
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="path"/> and returns its slide count, as counted from
    /// the deck's slide list — see <see cref="SlideCount"/> for details.
    /// </summary>
    /// <param name="path">The .pptx to read.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <returns>The number of slides in the deck.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="path"/> is blank, or the file it names is empty.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="path"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException"><paramref name="path"/>'s directory does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static async Task<int> SlideCountAsync(string path, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        var bytes = await File.ReadAllBytesAsync(path, ct).ConfigureAwait(false);
        return SlideCount(bytes);
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="path"/> and returns all text found on every slide, one
    /// entry per text-bearing body, in deck order — see <see cref="ExtractText(byte[])"/> for
    /// exactly what counts as a text-bearing body.
    /// </summary>
    /// <param name="path">The .pptx to read.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <returns>One entry per text-bearing body, in deck order.</returns>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="path"/> is blank, or the file it names is empty.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="path"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException"><paramref name="path"/>'s directory does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static async Task<IReadOnlyList<string>> ExtractTextAsync(
        string path, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);

        var bytes = await File.ReadAllBytesAsync(path, ct).ConfigureAwait(false);
        return ExtractText(bytes);
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="inputPath"/>, replaces every key with its value across
    /// all slide text, and writes the result to <paramref name="outputPath"/> — see
    /// <see cref="ReplaceText"/> for exactly what counts as a match and how formatting survives
    /// it. The two paths may be the same file: the updated bytes are computed in full before
    /// <paramref name="outputPath"/> is opened, so a document that fails to process — cannot be
    /// read, or cannot be edited — leaves <paramref name="outputPath"/> untouched. That guarantee
    /// does not extend to a failure during the write itself: a full disk, a cancellation, or the
    /// process dying mid-write can still leave a partial file, so in-place editing of an
    /// irreplaceable document is not crash-safe.
    /// </summary>
    /// <param name="inputPath">The .pptx to read.</param>
    /// <param name="outputPath">Where to write the result. Overwritten if it exists.</param>
    /// <param name="replacements">Each key is replaced by its value, longest key wins per match.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    /// <exception cref="ArgumentNullException">A path or <paramref name="replacements"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// A path is blank, or the file at <paramref name="inputPath"/> is empty.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="inputPath"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException">
    /// <paramref name="inputPath"/>'s or <paramref name="outputPath"/>'s directory does not exist.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or edited.</exception>
    public static async Task ReplaceTextAsync(
        string inputPath, string outputPath,
        IReadOnlyDictionary<string, string> replacements, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var bytes = await File.ReadAllBytesAsync(inputPath, ct).ConfigureAwait(false);
        var result = ReplaceText(bytes, replacements);
        await File.WriteAllBytesAsync(outputPath, result, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="inputPath"/>, replaces every shape whose text is exactly
    /// <paramref name="placeholder"/> with <paramref name="image"/>, and writes the result to
    /// <paramref name="outputPath"/> — see <see cref="ReplaceImage"/> for exactly what counts as a
    /// match and how the image is fit into the matched shape's box. The two paths may be the same
    /// file: the updated bytes are computed in full before <paramref name="outputPath"/> is opened,
    /// so a document that fails to process — cannot be read, or cannot be edited — leaves
    /// <paramref name="outputPath"/> untouched. That guarantee does not extend to a failure during
    /// the write itself: a full disk, a cancellation, or the process dying mid-write can still leave
    /// a partial file, so in-place editing of an irreplaceable document is not crash-safe.
    /// </summary>
    /// <param name="inputPath">The .pptx to read.</param>
    /// <param name="outputPath">Where to write the result. Overwritten if it exists.</param>
    /// <param name="placeholder">The placeholder text a shape must hold, and hold only.</param>
    /// <param name="image">PNG or JPEG bytes. The format is decided by the bytes, never a filename.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    /// <exception cref="ArgumentNullException">
    /// A path, <paramref name="placeholder"/> or <paramref name="image"/> is null.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// A path is blank, the file at <paramref name="inputPath"/> is empty, <paramref name="placeholder"/>
    /// is blank, or <paramref name="image"/> is empty.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="inputPath"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException">
    /// <paramref name="inputPath"/>'s or <paramref name="outputPath"/>'s directory does not exist.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The placeholder appears nowhere, appears only inside a grouped shape, a matched shape holds
    /// other text, a matched shape has no explicit position, the image is neither PNG nor JPEG,
    /// or the package could not be edited.
    /// </exception>
    public static async Task ReplaceImageAsync(
        string inputPath, string outputPath, string placeholder, byte[] image,
        CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        var bytes = await File.ReadAllBytesAsync(inputPath, ct).ConfigureAwait(false);
        var result = ReplaceImage(bytes, placeholder, image);
        await File.WriteAllBytesAsync(outputPath, result, ct).ConfigureAwait(false);
    }
}
