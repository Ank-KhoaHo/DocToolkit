using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

using OfficeIMOPowerPointPowerPointPresentation = OfficeIMO.PowerPoint.PowerPointPresentation;
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
    /// <c>WorkbookEditor.CreateToFileAsync</c>:
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
            .DrainAsync(source, "Presentation content was empty.", nameof(source), "Failed to read PPTX. See the inner exception for details.", ct)
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
            throw new DocumentConversionException("Failed to read PPTX. See the inner exception for details.", ex);
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
    ///
    /// Each slide's SmartArt diagrams follow that slide's own text-bearing bodies, one entry per
    /// diagram — see <see cref="ReadSmartArt"/>. A SmartArt diagram's text lives in a diagram data
    /// part, not a &lt;p:txBody&gt;, so it is not itself a text-bearing body and was invisible here
    /// before this was added; it is not something <see cref="ReplaceText"/> can reach.
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
            .DrainAsync(source, "Presentation content was empty.", nameof(source), "Failed to read PPTX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        return ExtractTextCore(ms);
    }

    private static IReadOnlyList<string> ExtractTextCore(MemoryStream ms)
    {
        try
        {
            using var doc = OpenDocument(ms, false);
            var slides = SlidesInDeckOrder(PresentationPartOf(doc)).Select(p => p.Slide).ToList();
            var smartArtBySlide = SmartArtTextBySlide(ms.ToArray());

            var results = new List<string>();
            for (var i = 0; i < slides.Count; i++)
            {
                if (slides[i] is { } slide) results.AddRange(TextBodiesOf(slide));
                if (i < smartArtBySlide.Count) results.AddRange(smartArtBySlide[i]);
            }

            return results;
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read PPTX. See the inner exception for details.", ex);
        }
    }

    /// <summary>
    /// All text-bearing bodies on one slide, in document order — shared by <see cref="ExtractTextCore"/>
    /// (all slides) and <see cref="ReadSlideCore"/> (one slide), so the two can never disagree
    /// about what counts as a text-bearing body.
    ///
    /// PowerPoint stores shape text as a:t runs under a:p paragraphs - Wordprocessing's w:t is the
    /// DOCX equivalent, not what PPTX uses. Grouping by the paragraph's parent yields one entry per
    /// text body while preserving document order.
    /// </summary>
    private static IReadOnlyList<string> TextBodiesOf(P.Slide slide)
    {
        var results = new List<string>();

        foreach (var body in slide.Descendants<A.Paragraph>().GroupBy(p => p.Parent))
        {
            var bodyText = string.Join(
                "\n",
                body.Select(p => string.Concat(p.Descendants<A.Text>().Select(t => t.Text))));

            if (bodyText.Length > 0) results.Add(bodyText);
        }

        return results;
    }

    /// <summary>
    /// All text found on slide <paramref name="index"/> — see <see cref="ExtractText(byte[])"/>
    /// for exactly what counts as a text-bearing body. Same per-body granularity as that method,
    /// scoped to one slide.
    /// </summary>
    /// <param name="pptx">The presentation to read.</param>
    /// <param name="index">1-based, because that is how a reader numbers slides.</param>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="index"/> is below 1, or above the deck's slide count.
    /// </exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static IReadOnlyList<string> ReadSlide(byte[] pptx, int index)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        using var ms = OpenForWrite(pptx);
        return ReadSlideCore(ms, index);
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/> and returns all text found on slide
    /// <paramref name="index"/> — see <see cref="ReadSlide(byte[], int)"/> for exactly what counts
    /// as a text-bearing body. <paramref name="source"/> is <b>read</b> to its end and is neither
    /// disposed, closed nor sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="index"/> is below 1, or above the deck's slide count.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static async Task<IReadOnlyList<string>> ReadSlideAsync(
        Stream source, int index, CancellationToken ct = default)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);
        StreamPipeline.RequireReadable(source, nameof(source));
        ct.ThrowIfCancellationRequested();

        using var ms = await StreamPipeline
            .DrainAsync(source, "Presentation content was empty.", nameof(source), "Failed to read PPTX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        return ReadSlideCore(ms, index);
    }

    private static IReadOnlyList<string> ReadSlideCore(MemoryStream ms, int index)
    {
        try
        {
            using var doc = OpenDocument(ms, false);

            var slideParts = SlidesInDeckOrder(PresentationPartOf(doc)).ToList();
            if (index > slideParts.Count)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(index), index,
                    $"Slide {index} was requested from a deck with {slideParts.Count} slide(s).");
            }

            var slide = slideParts[index - 1].Slide;
            return slide is null ? Array.Empty<string>() : TextBodiesOf(slide);
        }
        catch (Exception ex) when (ex is not DocumentConversionException and not ArgumentOutOfRangeException)
        {
            throw new DocumentConversionException("Failed to read PPTX. See the inner exception for details.", ex);
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
            .DrainAsync(source, "Presentation content was empty.", nameof(source), "Failed to edit PPTX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        ReplaceTextCore(ms, replacements);

        await StreamPipeline.EmitAsync(ms, destination, "Failed to edit PPTX. See the inner exception for details.", ct).ConfigureAwait(false);
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
            throw new DocumentConversionException("Failed to edit PPTX. See the inner exception for details.", ex);
        }
    }

    /// <summary>
    /// Replaces every shape whose text is exactly <paramref name="placeholder"/> with
    /// <paramref name="image"/>, which is scaled to fit inside that shape's box and centred there.
    ///
    /// Position and size come from the template, so there is nothing to pass: a designer draws a
    /// box in PowerPoint where the image belongs and the image lands there. This deliberately does
    /// not mirror <c>DocxEditor.ReplaceImage</c>'s size arguments — a DOCX image is inline
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
            throw new DocumentConversionException("Failed to edit PPTX. See the inner exception for details.", ex);
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
                $"'{placeholder}' does not appear in any shape, so nothing was replaced. Check "
                + "the placeholder text matches a shape's text exactly.");
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
            .DrainAsync(source, "Presentation content was empty.", nameof(source), "Failed to edit PPTX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        ReplaceImageCore(ms, placeholder, image);

        await StreamPipeline.EmitAsync(ms, destination, "Failed to edit PPTX. See the inner exception for details.", ct).ConfigureAwait(false);
    }

    /// <summary>
    /// A copy of <paramref name="pptx"/> with the slides at <paramref name="indices"/> removed.
    /// </summary>
    /// <param name="pptx">The presentation to remove slides from. It is not modified.</param>
    /// <param name="indices">
    /// 1-based slide numbers to remove, each exactly once and in any order — not a contiguous
    /// range, so <c>[2, 7]</c> removes exactly those two slides in one call.
    /// </param>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> or <paramref name="indices"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="pptx"/> is empty, or <paramref name="indices"/> contains a duplicate.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// An index in <paramref name="indices"/> is outside the deck's slide range, or removing every
    /// listed index would leave a zero-slide deck.
    /// </exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or edited.</exception>
    public static byte[] RemoveSlides(byte[] pptx, IEnumerable<int> indices)
    {
        ArgumentNullException.ThrowIfNull(indices);
        var wanted = indices.ToArray();

        using var ms = OpenForWrite(pptx);
        RemoveSlidesCore(ms, wanted);
        return ms.ToArray();
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/>, removes the slides at
    /// <paramref name="indices"/>, and writes the result to <paramref name="destination"/> — see
    /// <see cref="RemoveSlides"/> for exactly what <paramref name="indices"/> accepts.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought, and neither has to be seekable.
    /// </summary>
    /// <param name="source">The stream the .pptx package is read from.</param>
    /// <param name="indices">1-based slide numbers to remove, each exactly once, any order.</param>
    /// <param name="destination">The stream the edited .pptx package is written to.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, <paramref name="destination"/>
    /// is not writable, or <paramref name="indices"/> contains a duplicate.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// An index in <paramref name="indices"/> is outside the deck's slide range, or removing every
    /// listed index would leave a zero-slide deck.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or edited.</exception>
    public static async Task RemoveSlidesAsync(
        Stream source, IEnumerable<int> indices, Stream destination, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(indices);
        var wanted = indices.ToArray();
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var ms = await StreamPipeline
            .DrainAsync(source, "Presentation content was empty.", nameof(source), "Failed to edit PPTX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        RemoveSlidesCore(ms, wanted);

        await StreamPipeline.EmitAsync(ms, destination, "Failed to edit PPTX. See the inner exception for details.", ct).ConfigureAwait(false);
    }

    private static void RemoveSlidesCore(MemoryStream ms, IReadOnlyList<int> indices)
    {
        try
        {
            using (var doc = OpenDocument(ms, true))
            {
                var presentationPart = PresentationPartOf(doc);
                var slideParts = SlidesInDeckOrder(presentationPart).ToList();

                if (indices.Distinct().Count() != indices.Count)
                {
                    throw new ArgumentException(
                        $"The indices must not repeat. Got [{string.Join(", ", indices)}].",
                        nameof(indices));
                }

                foreach (var index in indices)
                {
                    if (index < 1 || index > slideParts.Count)
                    {
                        throw new ArgumentOutOfRangeException(
                            nameof(indices), index,
                            $"Slide {index} was requested for removal from a deck with " +
                            $"{slideParts.Count} slide(s).");
                    }
                }

                if (indices.Count >= slideParts.Count)
                {
                    throw new ArgumentOutOfRangeException(
                        nameof(indices), indices.Count,
                        $"Removing {indices.Count} of {slideParts.Count} slide(s) would leave " +
                        "nothing. A zero-slide deck is not a presentation.");
                }

                var presentation = presentationPart.Presentation!;
                var slideIdList = presentation.SlideIdList!;
                var toRemove = indices.Select(i => slideParts[i - 1]).ToList();
                var idsToRemove = toRemove.Select(part => presentationPart.GetIdOfPart(part)).ToHashSet();

                foreach (var slideId in slideIdList.Elements<P.SlideId>().ToList())
                {
                    if (idsToRemove.Contains(slideId.RelationshipId!.Value!))
                    {
                        slideId.Remove();
                    }
                }

                foreach (var part in toRemove)
                {
                    presentationPart.DeletePart(part);
                }

                presentation.Save();
            }
        }
        catch (Exception ex) when (ex is not DocumentConversionException and not ArgumentOutOfRangeException and not ArgumentException)
        {
            throw new DocumentConversionException("Failed to edit PPTX. See the inner exception for details.", ex);
        }
    }

    /// <summary>
    /// A copy of <paramref name="pptx"/> with its slides in the order given by
    /// <paramref name="order"/>, which holds 1-based slide numbers.
    /// </summary>
    /// <param name="pptx">The presentation to reorder. It is not modified.</param>
    /// <param name="order">
    /// A <b>permutation of every slide</b> — the same slides, in a different order. Not a subset,
    /// and no repeats.
    /// </param>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> or <paramref name="order"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="pptx"/> is empty, or <paramref name="order"/> is not a permutation of
    /// 1..<c>SlideCount</c>.
    /// </exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or edited.</exception>
    public static byte[] ReorderSlides(byte[] pptx, IEnumerable<int> order)
    {
        ArgumentNullException.ThrowIfNull(order);
        var wanted = order.ToArray();

        using var ms = OpenForWrite(pptx);
        ReorderSlidesCore(ms, wanted);
        return ms.ToArray();
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/>, reorders its slides per
    /// <paramref name="order"/>, and writes the result to <paramref name="destination"/> — see
    /// <see cref="ReorderSlides"/> for exactly what <paramref name="order"/> must contain.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought, and neither has to be seekable.
    /// </summary>
    /// <param name="source">The stream the .pptx package is read from.</param>
    /// <param name="order">A permutation of every slide's 1-based number.</param>
    /// <param name="destination">The stream the edited .pptx package is written to.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, <paramref name="destination"/>
    /// is not writable, or <paramref name="order"/> is not a permutation of every slide.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or edited.</exception>
    public static async Task ReorderSlidesAsync(
        Stream source, IEnumerable<int> order, Stream destination, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(order);
        var wanted = order.ToArray();
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var ms = await StreamPipeline
            .DrainAsync(source, "Presentation content was empty.", nameof(source), "Failed to edit PPTX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        ReorderSlidesCore(ms, wanted);

        await StreamPipeline.EmitAsync(ms, destination, "Failed to edit PPTX. See the inner exception for details.", ct).ConfigureAwait(false);
    }

    private static void ReorderSlidesCore(MemoryStream ms, IReadOnlyList<int> order)
    {
        try
        {
            using (var doc = OpenDocument(ms, true))
            {
                var presentationPart = PresentationPartOf(doc);
                var slideParts = SlidesInDeckOrder(presentationPart).ToList();

                var expected = Enumerable.Range(1, slideParts.Count);
                if (!order.OrderBy(i => i).SequenceEqual(expected))
                {
                    throw new ArgumentException(
                        $"The order must be a permutation of slides 1-{slideParts.Count}, each " +
                        $"exactly once. Got [{string.Join(", ", order)}].",
                        nameof(order));
                }

                var presentation = presentationPart.Presentation!;
                var slideIdList = presentation.SlideIdList!;
                var existingIds = slideIdList.Elements<P.SlideId>().ToList();

                foreach (var id in existingIds) id.Remove();

                foreach (var index in order)
                {
                    var relationshipId = presentationPart.GetIdOfPart(slideParts[index - 1]);
                    var originalId = existingIds.Single(id => id.RelationshipId!.Value == relationshipId);
                    slideIdList.Append(originalId);
                }

                presentation.Save();
            }
        }
        catch (Exception ex) when (ex is not DocumentConversionException and not ArgumentException)
        {
            throw new DocumentConversionException("Failed to edit PPTX. See the inner exception for details.", ex);
        }
    }

    /// <summary>
    /// A copy of <paramref name="pptx"/> with <paramref name="slides"/> inserted so that the
    /// first of them becomes slide <paramref name="atIndex"/>. See <see cref="Create"/> for what
    /// a <see cref="PptxSlide"/> becomes.
    ///
    /// Each inserted slide attaches to the layout of the slide immediately before the insertion
    /// point — or, when inserting at position 1 of a non-empty deck, the layout of what is
    /// currently the first slide — so it renders consistently with its neighbours rather than
    /// requiring a caller to name one. A deck with no slides at all falls back to its first slide
    /// master's first layout.
    ///
    /// The inserted slide's title and body boxes inherit the target layout's own placeholder
    /// position only when that layout <b>positions that placeholder itself</b> — it has a
    /// placeholder of the matching role (same type, and for the body, the same index),
    /// <i>and</i> that placeholder carries its own position and size, <i>and</i> it is one of the
    /// layout's own TOP-LEVEL shapes: a placeholder nested inside a group (<c>p:grpSp</c>) on the
    /// layout does not count, matching role and complete geometry notwithstanding. A hand-designed
    /// deck's own title/body geometry is then honoured rather than overridden.
    ///
    /// Otherwise the shape keeps this library's own fixed coordinates, rescaled to fit the target
    /// deck's canvas size. That covers four distinct cases, and the first two are common: a layout
    /// with no placeholder of that role at all (an ordinary "Title Slide" layout uses
    /// <c>ctrTitle</c>/<c>subTitle</c>), a layout that names the role but leaves its geometry to
    /// the slide master (a stock "Title and Content" layout usually does), a layout whose
    /// placeholder carries only part of a box (a position with no size, or a size with no
    /// position), and a layout that positions the role's placeholder only inside a group.
    /// Inheriting in any of these would leave content schema-valid but with no box this library's
    /// render pipeline (which resolves a slide's inherited geometry from the layout's TOP-LEVEL
    /// shape tree only — never from inside a group, and never by resolving layout → master) could
    /// actually draw in — invisible in a render. That is why the fallback exists: it keeps content
    /// visible, positioned by this library's own choice, at the cost of not matching a
    /// hand-designed layout's own intended position in these cases.
    /// </summary>
    /// <param name="pptx">The presentation to insert into. It is not modified.</param>
    /// <param name="atIndex">
    /// 1-based position the first inserted slide will occupy. <c>1</c> puts them in front of
    /// everything; <c>SlideCount + 1</c> appends, which is deliberately allowed — it is the
    /// obvious way to say "after everything".
    /// </param>
    /// <param name="slides">The slides to insert, in order.</param>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> or <paramref name="slides"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="pptx"/> is empty, or an element of <paramref name="slides"/> is null.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="atIndex"/> is below 1 or more than one past the last slide.
    /// </exception>
    /// <exception cref="DocumentConversionException">
    /// The package could not be opened or edited, or the slide the new content would attach to has
    /// no layout of its own.
    /// </exception>
    public static byte[] InsertSlides(byte[] pptx, int atIndex, IEnumerable<PptxSlide> slides)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(atIndex, 1);
        var materialised = ValidateSlides(slides);

        using var ms = OpenForWrite(pptx);
        InsertSlidesCore(ms, atIndex, materialised);
        return ms.ToArray();
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/>, inserts <paramref name="slides"/> at
    /// <paramref name="atIndex"/>, and writes the result to <paramref name="destination"/> — see
    /// <see cref="InsertSlides"/> for the insertion and layout rules.
    ///
    /// <paramref name="source"/> is <b>read</b> to its end and <paramref name="destination"/> is
    /// <b>written</b>; neither is disposed, closed or sought, and neither has to be seekable.
    /// </summary>
    /// <param name="source">The stream the .pptx package is read from.</param>
    /// <param name="atIndex">1-based insertion position; <c>SlideCount + 1</c> appends.</param>
    /// <param name="slides">The slides to insert, in order.</param>
    /// <param name="destination">The stream the edited .pptx package is written to.</param>
    /// <param name="ct">Cancels the read, the edit and the write.</param>
    /// <exception cref="ArgumentNullException">Any argument is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="source"/> is not readable or held no bytes, <paramref name="destination"/>
    /// is not writable, or an element of <paramref name="slides"/> is null.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="atIndex"/> is below 1 or more than one past the last slide.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The package could not be opened or edited, or the slide the new content would attach to has
    /// no layout of its own.
    /// </exception>
    public static async Task InsertSlidesAsync(
        Stream source, int atIndex, IEnumerable<PptxSlide> slides, Stream destination,
        CancellationToken ct = default)
    {
        ArgumentOutOfRangeException.ThrowIfLessThan(atIndex, 1);
        var materialised = ValidateSlides(slides);
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        ct.ThrowIfCancellationRequested();

        using var ms = await StreamPipeline
            .DrainAsync(source, "Presentation content was empty.", nameof(source), "Failed to edit PPTX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        InsertSlidesCore(ms, atIndex, materialised);

        await StreamPipeline.EmitAsync(ms, destination, "Failed to edit PPTX. See the inner exception for details.", ct).ConfigureAwait(false);
    }

    private static void InsertSlidesCore(MemoryStream ms, int atIndex, IReadOnlyList<PptxSlide> slides)
    {
        try
        {
            using (var doc = OpenDocument(ms, true))
            {
                var presentationPart = PresentationPartOf(doc);
                var existingSlideParts = SlidesInDeckOrder(presentationPart).ToList();

                if (atIndex > existingSlideParts.Count + 1)
                {
                    throw new ArgumentOutOfRangeException(
                        nameof(atIndex), atIndex,
                        $"Cannot insert at slide {atIndex} of a deck with " +
                        $"{existingSlideParts.Count} slide(s); {existingSlideParts.Count + 1} " +
                        "appends and is the highest position allowed.");
                }

                var layoutPart = ResolveLayoutForInsertion(presentationPart, existingSlideParts, atIndex);
                var presentation = presentationPart.Presentation!;
                var slideIdList = presentation.SlideIdList!;

                var nextSlideId = slideIdList.Elements<P.SlideId>().Select(s => s.Id!.Value)
                    .DefaultIfEmpty(PptxDocumentWriter.FirstSlideId - 1).Max() + 1;

                // Resolved BEFORE any insertion: inserting shifts what "the element currently at
                // atIndex" would mean if looked up again mid-loop.
                var insertBeforeId = atIndex <= existingSlideParts.Count
                    ? slideIdList.Elements<P.SlideId>().Single(id =>
                        id.RelationshipId!.Value == presentationPart.GetIdOfPart(existingSlideParts[atIndex - 1]))
                    : null;

                foreach (var slide in slides)
                {
                    var slidePart = presentationPart.AddNewPart<SlidePart>();
                    slidePart.Slide = PptxDocumentWriter.BuildSlide(slide);

                    foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
                    {
                        var shapePlaceholder = shape.NonVisualShapeProperties
                            ?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>();
                        if (shapePlaceholder is not null
                            && LayoutHasMatchingPositionedPlaceholder(layoutPart, shapePlaceholder))
                        {
                            shape.ShapeProperties?.Transform2D?.Remove();
                        }
                    }

                    ScaleToFitDeck(slidePart.Slide, presentation.SlideSize);
                    slidePart.AddPart(layoutPart);
                    slidePart.Slide.Save();

                    var newSlideId = new P.SlideId
                    {
                        Id = nextSlideId++,
                        RelationshipId = presentationPart.GetIdOfPart(slidePart),
                    };

                    if (insertBeforeId is not null)
                    {
                        slideIdList.InsertBefore(newSlideId, insertBeforeId);
                    }
                    else
                    {
                        slideIdList.Append(newSlideId);
                    }
                }

                presentation.Save();
            }
        }
        catch (Exception ex) when (ex is not DocumentConversionException and not ArgumentOutOfRangeException)
        {
            throw new DocumentConversionException("Failed to edit PPTX. See the inner exception for details.", ex);
        }
    }

    /// <summary>
    /// Rescales every shape's position and size on <paramref name="slide"/> from the fixed design
    /// size <c>PptxDocumentWriter.BuildSlide</c> assumes (the same 16:9 canvas <see cref="Create"/>
    /// always writes) to <paramref name="targetSize"/>, so a slide inserted into a deck of a
    /// DIFFERENT size — a 4:3 deck, most commonly — is not left overhanging the canvas edge.
    /// <c>BuildSlide</c> itself stays unchanged and keeps producing content sized for its own
    /// writer's deck; this is the one place inserted content is adapted to a foreign one.
    /// </summary>
    private static void ScaleToFitDeck(P.Slide slide, P.SlideSize? targetSize)
    {
        if (targetSize?.Cx?.Value is not int targetCx || targetSize.Cy?.Value is not int targetCy)
            return;

        var scaleX = (double)targetCx / PptxDocumentWriter.SlideWidthEmu;
        var scaleY = (double)targetCy / PptxDocumentWriter.SlideHeightEmu;
        if (scaleX == 1.0 && scaleY == 1.0) return;

        foreach (var xfrm in slide.Descendants<A.Transform2D>())
        {
            if (xfrm.Offset is { } offset)
            {
                if (offset.X is not null) offset.X = (int)(offset.X.Value * scaleX);
                if (offset.Y is not null) offset.Y = (int)(offset.Y.Value * scaleY);
            }
            if (xfrm.Extents is { } extents)
            {
                if (extents.Cx is not null) extents.Cx = (int)(extents.Cx.Value * scaleX);
                if (extents.Cy is not null) extents.Cy = (int)(extents.Cy.Value * scaleY);
            }
        }
    }

    /// <summary>
    /// Whether <paramref name="layoutPart"/>'s own TOP-LEVEL shape tree has a placeholder with the
    /// same role as <paramref name="shapePlaceholder"/> — same <c>Type</c>, and for
    /// <see cref="P.PlaceholderValues.Body"/>, the same <c>Index</c> too — <b>which also carries a
    /// complete, usable box of its own: both an offset and an extent</b>. A role match alone is not
    /// enough, and neither is a bare <see cref="A.Transform2D"/> — only a top-level placeholder with
    /// both halves of its box gives the inserted shape something to inherit.
    /// </summary>
    /// <remarks>
    /// <b>The geometry half is not a refinement, it is the load-bearing half.</b> Many real
    /// PowerPoint layouts declare a placeholder's ROLE and leave its geometry to the slide master —
    /// measured on the eleven stock layouts in this repo's own sample deck, "Title and Content",
    /// "Two Content", "Title Only" and "Title and Vertical Text" all do, and "Title and Content" is
    /// what PowerPoint gives a new body slide by default. Dropping the inserted shape's own
    /// <c>a:xfrm</c> for one of those leaves it with no box from anywhere, because this repo's
    /// render pipeline resolves slide → layout and not layout → master: measured, the text
    /// disappeared from the render entirely. So an unpositioned match falls back to the shape's own
    /// fixed geometry, which <c>ScaleToFitDeck</c> then rescales like any other unmatched case.
    ///
    /// <b>Deliberately exact, not cross-type.</b> A layout whose title placeholder is
    /// <c>ctrTitle</c> rather than plain <c>title</c> — an ordinary "Title Slide" layout — does NOT
    /// match a <see cref="PptxDocumentWriter.BuildSlide"/>-built <c>title</c> shape. Measured: this
    /// repo's own render pipeline does not resolve that either, so treating them as compatible here
    /// would assert a guarantee nothing can verify.
    ///
    /// <b>Deliberately literal about the body's type, too.</b> ECMA-376 treats an untyped
    /// <c>&lt;p:ph idx="1"/&gt;</c> as a body placeholder by default, and this only matches an
    /// EXPLICIT <c>Type=Body</c> — so an untyped one never matches and always falls back. That is
    /// the safe direction (fixed geometry that renders, rather than inherited geometry that might
    /// not), and it is conservative rather than wrong.
    ///
    /// <b>A present <see cref="A.Transform2D"/> is not enough — it has to be complete.</b>
    /// <c>CT_Transform2D</c> declares both <c>a:off</c> and <c>a:ext</c> as optional, so a layout
    /// placeholder can carry <c>&lt;a:xfrm&gt;&lt;a:off .../&gt;&lt;/a:xfrm&gt;</c> with no
    /// <c>a:ext</c> — a non-null <c>Transform2D</c> that is still not a usable box. Measured
    /// directly: stripping <c>&lt;a:ext&gt;</c> from the "Section Header" layout's title
    /// placeholder — the positive control this method's own test suite uses — still passed a
    /// presence-only check, still stripped the inserted shape's own <c>a:xfrm</c>, and the title
    /// text still vanished from the render: the identical failure class this whole check exists to
    /// prevent. So both <c>.Offset</c> and <c>.Extents</c> must be non-null, not merely
    /// <c>Transform2D</c> itself.
    ///
    /// <b>Only the layout's TOP-LEVEL shapes are considered — a placeholder nested inside a group
    /// does not count, even with a matching role and a complete box.</b>
    /// <c>Descendants&lt;P.Shape&gt;()</c> also matches a shape nested inside a <c>p:grpSp</c>, but
    /// this repo's render pipeline (<c>PptxToPdfConverter</c>/OfficeIMO) resolves a slide's
    /// inherited geometry from the layout's TOP-LEVEL shape tree only, never from inside a group.
    /// Measured directly: wrapping "Section Header"'s title placeholder in a group — matching type,
    /// complete <c>a:off</c>/<c>a:ext</c>, schema-valid both before and after (0
    /// <c>OpenXmlValidator</c> errors) — still matched under a <c>Descendants</c>-based walk, still
    /// stripped the inserted slide's own <c>a:xfrm</c>, and the title text vanished from the render:
    /// the identical failure class every check above already guards against, one level further in.
    /// Same class of mistake as <c>DocxEditor</c> reaching into <c>w:txbxContent</c> and
    /// <c>TableRowFinder</c> reaching into nested tables — see <c>CLAUDE.md</c>.
    /// </remarks>
    private static bool LayoutHasMatchingPositionedPlaceholder(
        SlideLayoutPart layoutPart, P.PlaceholderShape shapePlaceholder)
    {
        // InsertSlides never read the layout's XML before this check existed. A null SlideLayout
        // would therefore be a NEW way for it to fail, and the outer catch would turn the resulting
        // NullReferenceException into an opaque DocumentConversionException. Degrade to the
        // fixed-geometry fallback instead — which is exactly what this method returning false means.
        if (layoutPart.SlideLayout is null) return false;

        // Direct children only, matching ReplaceImageCore's own walk of a SLIDE's shape tree. A
        // shape nested inside a p:grpSp is invisible to the render pipeline's layout-inheritance
        // resolution, so Descendants<P.Shape>() would report a match this library cannot actually
        // honour -- see the remarks above for the measured failure this replaced.
        var layoutShapes = layoutPart.SlideLayout.CommonSlideData?.ShapeTree?.Elements<P.Shape>()
            ?? Enumerable.Empty<P.Shape>();

        foreach (var layoutShape in layoutShapes)
        {
            var layoutPh = layoutShape.NonVisualShapeProperties
                ?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>();
            if (layoutPh is null) continue;

            if (layoutPh.Type?.Value != shapePlaceholder.Type?.Value) continue;

            if (shapePlaceholder.Type?.Value == P.PlaceholderValues.Body
                && (layoutPh.Index?.Value ?? 0) != (shapePlaceholder.Index?.Value ?? 0))
            {
                continue;
            }

            // The role matches, but the layout's box is missing or incomplete -- an a:xfrm with
            // only a:off or only a:ext is not enough to draw in, so treat it the same as no
            // Transform2D at all rather than inheriting half a box.
            var layoutXfrm = layoutShape.ShapeProperties?.Transform2D;
            if (layoutXfrm?.Offset is null || layoutXfrm.Extents is null) continue;

            return true;
        }

        return false;
    }

    /// <summary>
    /// The layout a slide inserted at <paramref name="atIndex"/> should attach to: the layout of
    /// the slide immediately before the insertion point, or — when inserting at position 1 of a
    /// non-empty deck — the layout of what is currently the first slide. Only an empty deck has no
    /// neighbour to borrow from, and then the first layout of the first slide master is used.
    /// </summary>
    /// <exception cref="DocumentConversionException">
    /// The relevant slide has no layout of its own, or the deck has no slide master with a layout
    /// to fall back on.
    /// </exception>
    private static SlideLayoutPart ResolveLayoutForInsertion(
        PresentationPart presentationPart, IReadOnlyList<SlidePart> slidesInDeckOrder, int atIndex)
    {
        if (slidesInDeckOrder.Count > 0)
        {
            var neighbourIndex = atIndex > 1 ? atIndex - 2 : 0;
            var neighbour = slidesInDeckOrder[neighbourIndex];

            return neighbour.SlideLayoutPart
                ?? throw new DocumentConversionException(
                    $"Slide {neighbourIndex + 1} has no layout of its own, so there is nothing " +
                    "for the inserted slide to attach to.");
        }

        var master = presentationPart.SlideMasterParts.FirstOrDefault()
            ?? throw new DocumentConversionException(
                "The deck has no slide master, so there is no layout for the inserted slide to " +
                "attach to.");

        var firstLayoutId = master.SlideMaster?.SlideLayoutIdList?.Elements<P.SlideLayoutId>().FirstOrDefault()
            ?? throw new DocumentConversionException(
                "The deck's slide master has no layout, so there is no layout for the inserted " +
                "slide to attach to.");

        return (SlideLayoutPart)master.GetPartById(firstLayoutId.RelationshipId!.Value!);
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
           ?? throw new DocumentConversionException(
               "Presentation has no presentation part. This usually means the file is not really "
               + "a .pptx (for example it was renamed from another format) or the upload is corrupt.");

    private static PresentationDocument OpenDocument(MemoryStream ms, bool isEditable)
    {
        try
        {
            return PresentationDocument.Open(ms, isEditable);
        }
        catch (Exception ex)
        {
            throw new DocumentConversionException("Failed to open PPTX. See the inner exception for details.", ex);
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

        var bytes = await FilePipeline.ReadAsync(path, nameof(path), ct).ConfigureAwait(false);
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

        var bytes = await FilePipeline.ReadAsync(path, nameof(path), ct).ConfigureAwait(false);
        return ExtractText(bytes);
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="path"/> and returns all text found on slide
    /// <paramref name="index"/> — see <see cref="ReadSlide(byte[], int)"/> for exactly what counts
    /// as a text-bearing body.
    /// </summary>
    /// <param name="path">The .pptx to read.</param>
    /// <param name="index">1-based, because that is how a reader numbers slides.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="path"/> is blank, or the file it names is empty.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="index"/> is below 1, or above the deck's slide count.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="path"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException"><paramref name="path"/>'s directory does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static async Task<IReadOnlyList<string>> ReadSlideAsync(
        string path, int index, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        var bytes = await FilePipeline.ReadAsync(path, nameof(path), ct).ConfigureAwait(false);
        return ReadSlide(bytes, index);
    }

    /// <summary>
    /// The text of every SmartArt diagram on slide <paramref name="index"/>, one entry per
    /// diagram, each diagram's nodes joined with newlines in the order OfficeIMO reports them.
    ///
    /// A SmartArt diagram's text lives in a diagram data part, not a text-bearing shape body, so
    /// it is invisible to <see cref="ReadSlide"/> — <see cref="ExtractText(byte[])"/> reports it
    /// too, alongside every ordinary text-bearing body, for exactly that reason.
    ///
    /// An empty list means the slide has no SmartArt, which is not an error — the same convention
    /// <see cref="ReadSlide"/> uses for a slide with no text-bearing shapes.
    /// </summary>
    /// <param name="pptx">The presentation to read.</param>
    /// <param name="index">1-based, because that is how a reader numbers slides.</param>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="index"/> is below 1, or above the deck's slide count.
    /// </exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static IReadOnlyList<string> ReadSmartArt(byte[] pptx, int index)
    {
        ArgumentNullException.ThrowIfNull(pptx);
        if (pptx.Length == 0)
            throw new ArgumentException("Presentation content was empty.", nameof(pptx));
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        return ReadSmartArtCore(pptx, index);
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="source"/> and returns the text of every SmartArt
    /// diagram on slide <paramref name="index"/> — see <see cref="ReadSmartArt(byte[], int)"/>.
    /// <paramref name="source"/> is <b>read</b> to its end and is neither disposed, closed nor
    /// sought.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="source"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="source"/> is not readable or held no bytes.</exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="index"/> is below 1, or above the deck's slide count.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static async Task<IReadOnlyList<string>> ReadSmartArtAsync(
        Stream source, int index, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);
        ct.ThrowIfCancellationRequested();

        using var ms = await StreamPipeline
            .DrainAsync(source, "Presentation content was empty.", nameof(source), "Failed to read PPTX. See the inner exception for details.", ct)
            .ConfigureAwait(false);

        return ReadSmartArtCore(ms.ToArray(), index);
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="path"/> and returns the text of every SmartArt diagram
    /// on slide <paramref name="index"/> — see <see cref="ReadSmartArt(byte[], int)"/>.
    /// </summary>
    /// <param name="path">The .pptx to read.</param>
    /// <param name="index">1-based, because that is how a reader numbers slides.</param>
    /// <param name="ct">Cancels the read.</param>
    /// <exception cref="ArgumentNullException"><paramref name="path"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// <paramref name="path"/> is blank, or the file it names is empty.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="index"/> is below 1, or above the deck's slide count.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="path"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException"><paramref name="path"/>'s directory does not exist.</exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or read.</exception>
    public static async Task<IReadOnlyList<string>> ReadSmartArtAsync(
        string path, int index, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentOutOfRangeException.ThrowIfLessThan(index, 1);

        var bytes = await FilePipeline.ReadAsync(path, nameof(path), ct).ConfigureAwait(false);
        return ReadSmartArt(bytes, index);
    }

    /// <summary>
    /// The text of every SmartArt diagram on every slide, in deck order — <see cref="ReadSmartArt"/>
    /// applied slide by slide, shared with <see cref="ExtractTextCore"/> so the two can never
    /// disagree about which diagrams exist or what order they come in.
    /// </summary>
    private static IReadOnlyList<IReadOnlyList<string>> SmartArtTextBySlide(byte[] pptx)
    {
        try
        {
            using var source = new MemoryStream(pptx, writable: false);
            using var document = OfficeIMOPowerPointPowerPointPresentation.Load(source);

            return document.Slides
                .Select(slide => (IReadOnlyList<string>)slide.SmartArts
                    .Select(art => string.Join("\n", art.GetNodeTexts()))
                    .ToList())
                .ToList();
        }
        catch (Exception ex) when (ex is not DocumentConversionException)
        {
            throw new DocumentConversionException("Failed to read PPTX. See the inner exception for details.", ex);
        }
    }

    private static IReadOnlyList<string> ReadSmartArtCore(byte[] pptx, int index)
    {
        var bySlide = SmartArtTextBySlide(pptx);
        if (index > bySlide.Count)
        {
            throw new ArgumentOutOfRangeException(
                nameof(index), index,
                $"Slide {index} was requested from a deck with {bySlide.Count} slide(s).");
        }

        return bySlide[index - 1];
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

        var bytes = await FilePipeline.ReadAsync(inputPath, nameof(inputPath), ct).ConfigureAwait(false);
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

        var bytes = await FilePipeline.ReadAsync(inputPath, nameof(inputPath), ct).ConfigureAwait(false);
        var result = ReplaceImage(bytes, placeholder, image);
        await File.WriteAllBytesAsync(outputPath, result, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="inputPath"/>, removes the slides at
    /// <paramref name="indices"/>, and writes the result to <paramref name="outputPath"/> — see
    /// <see cref="RemoveSlides"/> for exactly what <paramref name="indices"/> accepts. The two
    /// paths may be the same file: the updated bytes are computed in full before
    /// <paramref name="outputPath"/> is opened.
    /// </summary>
    /// <param name="inputPath">The .pptx to read.</param>
    /// <param name="outputPath">Where to write the result. Overwritten if it exists.</param>
    /// <param name="indices">1-based slide numbers to remove, each exactly once, any order.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    /// <exception cref="ArgumentNullException">A path or <paramref name="indices"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// A path is blank, the file at <paramref name="inputPath"/> is empty, or
    /// <paramref name="indices"/> contains a duplicate.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// An index in <paramref name="indices"/> is outside the deck's slide range, or removing every
    /// listed index would leave a zero-slide deck.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="inputPath"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException">
    /// <paramref name="inputPath"/>'s or <paramref name="outputPath"/>'s directory does not exist.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or edited.</exception>
    public static async Task RemoveSlidesAsync(
        string inputPath, string outputPath, IEnumerable<int> indices, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        ArgumentNullException.ThrowIfNull(indices);

        var bytes = await FilePipeline.ReadAsync(inputPath, nameof(inputPath), ct).ConfigureAwait(false);
        var result = RemoveSlides(bytes, indices);
        await File.WriteAllBytesAsync(outputPath, result, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="inputPath"/>, reorders its slides per
    /// <paramref name="order"/>, and writes the result to <paramref name="outputPath"/> — see
    /// <see cref="ReorderSlides"/> for exactly what <paramref name="order"/> must contain. The two
    /// paths may be the same file: the updated bytes are computed in full before
    /// <paramref name="outputPath"/> is opened.
    /// </summary>
    /// <param name="inputPath">The .pptx to read.</param>
    /// <param name="outputPath">Where to write the result. Overwritten if it exists.</param>
    /// <param name="order">A permutation of every slide's 1-based number.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    /// <exception cref="ArgumentNullException">A path or <paramref name="order"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// A path is blank, the file at <paramref name="inputPath"/> is empty, or
    /// <paramref name="order"/> is not a permutation of every slide.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="inputPath"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException">
    /// <paramref name="inputPath"/>'s or <paramref name="outputPath"/>'s directory does not exist.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">The package could not be opened or edited.</exception>
    public static async Task ReorderSlidesAsync(
        string inputPath, string outputPath, IEnumerable<int> order, CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        ArgumentNullException.ThrowIfNull(order);

        var bytes = await FilePipeline.ReadAsync(inputPath, nameof(inputPath), ct).ConfigureAwait(false);
        var result = ReorderSlides(bytes, order);
        await File.WriteAllBytesAsync(outputPath, result, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// Reads a .pptx from <paramref name="inputPath"/>, inserts <paramref name="slides"/> at
    /// <paramref name="atIndex"/>, and writes the result to <paramref name="outputPath"/> — see
    /// <see cref="InsertSlides"/> for the insertion and layout rules. The two paths may be the
    /// same file: the updated bytes are computed in full before <paramref name="outputPath"/> is
    /// opened.
    /// </summary>
    /// <param name="inputPath">The .pptx to read.</param>
    /// <param name="outputPath">Where to write the result. Overwritten if it exists.</param>
    /// <param name="atIndex">1-based insertion position; <c>SlideCount + 1</c> appends.</param>
    /// <param name="slides">The slides to insert, in order.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    /// <exception cref="ArgumentNullException">A path or <paramref name="slides"/> is null.</exception>
    /// <exception cref="ArgumentException">
    /// A path is blank, the file at <paramref name="inputPath"/> is empty, or an element of
    /// <paramref name="slides"/> is null.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="atIndex"/> is below 1 or more than one past the last slide.
    /// </exception>
    /// <exception cref="FileNotFoundException"><paramref name="inputPath"/> does not exist.</exception>
    /// <exception cref="DirectoryNotFoundException">
    /// <paramref name="inputPath"/>'s or <paramref name="outputPath"/>'s directory does not exist.
    /// </exception>
    /// <exception cref="OperationCanceledException"><paramref name="ct"/> was cancelled.</exception>
    /// <exception cref="DocumentConversionException">
    /// The package could not be opened or edited, or the slide the new content would attach to has
    /// no layout of its own.
    /// </exception>
    public static async Task InsertSlidesAsync(
        string inputPath, string outputPath, int atIndex, IEnumerable<PptxSlide> slides,
        CancellationToken ct = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);
        ArgumentOutOfRangeException.ThrowIfLessThan(atIndex, 1);
        ArgumentNullException.ThrowIfNull(slides);

        var bytes = await FilePipeline.ReadAsync(inputPath, nameof(inputPath), ct).ConfigureAwait(false);
        var result = InsertSlides(bytes, atIndex, slides);
        await File.WriteAllBytesAsync(outputPath, result, ct).ConfigureAwait(false);
    }

    /// <summary>
    /// A copy of <paramref name="pptx"/> encrypted with <paramref name="password"/>, so it cannot
    /// be opened without one.
    /// </summary>
    /// <remarks>
    /// <b>This is file encryption, not presentation protection.</b> Office offers both under the same
    /// menu and they are not the same thing: this scrambles the whole file, so nothing can be read
    /// without the password. The other kind - a flag asking a reader not to edit - is a request
    /// rather than a lock, and is deliberately not exposed here.
    ///
    /// <b>The result is not a PPTX package any more.</b> An encrypted Office document is a
    /// compound file with the package sealed inside it, so every other method on this class refuses
    /// it - call <see cref="Unprotect(byte[], string)"/> first. That refusal is the honest
    /// behaviour: those methods could not read the content even if they tried.
    /// </remarks>
    /// <param name="pptx">The presentation to encrypt.</param>
    /// <param name="password">The password required to open the result. May not be empty.</param>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">The presentation could not be read or encrypted.</exception>
    public static byte[] Protect(byte[] pptx, string password)
    {
        ArgumentNullException.ThrowIfNull(pptx);
        if (pptx.Length == 0)
            throw new ArgumentException("PPTX content was empty.", nameof(pptx));
        OfficeCrypto.RequirePassword(password, nameof(password));

        return OfficeCrypto.TranslateWrite(() =>
        {
            using var source = new MemoryStream(pptx, writable: false);
            using var document = OfficeIMOPowerPointPowerPointPresentation.Load(source);
            using var encrypted = new MemoryStream();
            document.SaveEncrypted(encrypted, password);
            return encrypted.ToArray();
        }, "PPTX");
    }

    /// <summary>
    /// A copy of <paramref name="pptx"/> with its encryption removed, so the rest of this class
    /// can work on it.
    /// </summary>
    /// <remarks>
    /// <b>The output is not protected in any way.</b> That is what was asked for, but the bytes
    /// this returns are readable by anyone who obtains them.
    ///
    /// A presentation that was never encrypted is reported as such rather than passed through, because
    /// silently returning the input would make a broken pipeline look like a working one.
    /// </remarks>
    /// <param name="pptx">The encrypted presentation.</param>
    /// <param name="password">The password the presentation was encrypted with.</param>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> or <paramref name="password"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="pptx"/> is empty, or <paramref name="password"/> is empty.</exception>
    /// <exception cref="DocumentConversionException">
    /// The password was wrong, the presentation was not encrypted, or it could not be read.
    /// </exception>
    public static byte[] Unprotect(byte[] pptx, string password)
    {
        ArgumentNullException.ThrowIfNull(pptx);
        if (pptx.Length == 0)
            throw new ArgumentException("PPTX content was empty.", nameof(pptx));
        OfficeCrypto.RequirePassword(password, nameof(password));

        return OfficeCrypto.Translate(() =>
        {
            using var source = new MemoryStream(pptx, writable: false);
            using var document = OfficeIMOPowerPointPowerPointPresentation.LoadEncrypted(source, password);
            using var plain = new MemoryStream();
            document.Save(plain);
            return plain.ToArray();
        }, "PPTX");
    }

    /// <summary>
    /// Whether <paramref name="pptx"/> is an ENCRYPTED Office document.
    /// </summary>
    /// <remarks>
    /// <b>This is not a validity check, and a <see langword="false"/> is not a promise that
    /// anything else will succeed.</b> It distinguishes an encrypted document from a plain one;
    /// input that is neither — an image, a PDF, a text file, random bytes — is not encrypted, so
    /// this answers <see langword="false"/> for it, while every other method on this class refuses
    /// it. Measured over real files: a JPEG and a PDF both return <see langword="false"/> here and
    /// both throw from <c>ExtractText</c>.
    ///
    /// <b>The summary used to say "that is, whether the other methods on this class will refuse
    /// it".</b> That reads as a guard — test it, and if false, proceed — and takes the wrong branch
    /// for every input that is not a document at all. The behaviour was always right and only the
    /// sentence was wrong, which is why the fix is here and not in the code.
    ///
    /// Reads the file signature; it does not try the password and does not need one. A plain PPTX
    /// is a ZIP package, an encrypted one is a compound file, and the two are distinguishable from
    /// their first eight bytes.
    /// </remarks>
    /// <param name="pptx">The bytes to inspect.</param>
    /// <exception cref="ArgumentNullException"><paramref name="pptx"/> is null.</exception>
    public static bool IsProtected(byte[] pptx)
    {
        ArgumentNullException.ThrowIfNull(pptx);

        return OfficeCrypto.IsEncrypted(pptx);
    }

    /// <summary>
    /// Reads a presentation from <paramref name="source"/> and writes the encrypted copy to
    /// <paramref name="destination"/>.
    ///
    /// Neither stream is disposed, closed or sought.
    /// </summary>
    /// <inheritdoc cref="Protect(byte[], string)" path="/remarks|/exception"/>
    /// <param name="source">The stream the presentation is read from.</param>
    /// <param name="destination">The stream the encrypted presentation is written to.</param>
    /// <param name="password">The password required to open the result. May not be empty.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    public static async Task ProtectAsync(
        Stream source, Stream destination, string password, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        OfficeCrypto.RequirePassword(password, nameof(password));
        ct.ThrowIfCancellationRequested();

        using var buffer = await StreamPipeline
            .DrainAsync(source, "PPTX content was empty.", nameof(source),
                        "Failed to encrypt the PPTX.", ct)
            .ConfigureAwait(false);

        using var result = new MemoryStream(Protect(buffer.ToArray(), password), writable: false);
        await StreamPipeline
            .EmitAsync(result, destination, "Failed to encrypt the PPTX.", ct)
            .ConfigureAwait(false);
    }

    /// <summary>
    /// Reads an encrypted presentation from <paramref name="source"/> and writes the unprotected copy to
    /// <paramref name="destination"/>.
    ///
    /// Neither stream is disposed, closed or sought.
    /// </summary>
    /// <inheritdoc cref="Unprotect(byte[], string)" path="/remarks|/exception"/>
    /// <param name="source">The stream the encrypted presentation is read from.</param>
    /// <param name="destination">The stream the unprotected presentation is written to.</param>
    /// <param name="password">The password the presentation was encrypted with.</param>
    /// <param name="ct">Cancels the read and the write.</param>
    public static async Task UnprotectAsync(
        Stream source, Stream destination, string password, CancellationToken ct = default)
    {
        StreamPipeline.RequireReadable(source, nameof(source));
        StreamPipeline.RequireWritable(destination, nameof(destination));
        OfficeCrypto.RequirePassword(password, nameof(password));
        ct.ThrowIfCancellationRequested();

        using var buffer = await StreamPipeline
            .DrainAsync(source, "PPTX content was empty.", nameof(source),
                        "Failed to read the encrypted PPTX.", ct)
            .ConfigureAwait(false);

        using var result = new MemoryStream(Unprotect(buffer.ToArray(), password), writable: false);
        await StreamPipeline
            .EmitAsync(result, destination, "Failed to read the encrypted PPTX.", ct)
            .ConfigureAwait(false);
    }
}
