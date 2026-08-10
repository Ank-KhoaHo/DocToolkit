using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit;

/// <summary>
/// Creates the header and footer parts a <see cref="PageSetup"/> asks for, and returns the
/// references that belong in <c>w:sectPr</c>.
/// </summary>
/// <remarks>
/// Separate from <c>SectionPropertiesFactory</c> because these are package PARTS, not elements: a
/// header lives in its own part with its own relationship id, and only the reference goes in
/// <c>sectPr</c>. Keeping the two apart leaves each with one job.
/// </remarks>
internal static class HeaderFooterFactory
{
    public static IReadOnlyList<OpenXmlElement> CreateReferences(MainDocumentPart main, PageSetup page)
    {
        var references = new List<OpenXmlElement>();

        if (page.Header is { } header)
            references.Add(CreateHeader(main, header, HeaderFooterValues.Default));

        if (page.Footer is { } footer)
            references.Add(CreateFooter(main, footer, HeaderFooterValues.Default));

        // Only when the caller asked for a distinct first page. A first-page reference without
        // w:titlePg is ignored by Word, and w:titlePg without one leaves page one blank - which is
        // exactly what a caller passing null asked for.
        if (page.HasDistinctFirstPage)
        {
            if (page.FirstPageHeader is { } firstHeader)
                references.Add(CreateHeader(main, firstHeader, HeaderFooterValues.First));

            if (page.FirstPageFooter is { } firstFooter)
                references.Add(CreateFooter(main, firstFooter, HeaderFooterValues.First));
        }

        return references;
    }

    private static HeaderReference CreateHeader(
        MainDocumentPart main, DocxHeader content, HeaderFooterValues type)
    {
        var part = main.AddNewPart<HeaderPart>();
        part.Header = new Header(BuildParagraph(content));
        part.Header.Save();

        return new HeaderReference { Type = type, Id = main.GetIdOfPart(part) };
    }

    private static FooterReference CreateFooter(
        MainDocumentPart main, DocxHeader content, HeaderFooterValues type)
    {
        var part = main.AddNewPart<FooterPart>();
        part.Footer = new Footer(BuildParagraph(content));
        part.Footer.Save();

        return new FooterReference { Type = type, Id = main.GetIdOfPart(part) };
    }

    private static Paragraph BuildParagraph(DocxHeader content)
    {
        var paragraph = new Paragraph(
            new ParagraphProperties(new Justification { Val = ToJustification(content.Alignment) }));

        foreach (var segment in content.Segments)
        {
            switch (segment)
            {
                case DocxHeaderSegment.LiteralSegment literal:
                    // Space preserved: a header is often "Page " + field, and XML would otherwise
                    // collapse the trailing space and print "Page3".
                    paragraph.AppendChild(new Run(
                        new Text(literal.Value) { Space = SpaceProcessingModeValues.Preserve }));
                    break;

                case DocxHeaderSegment.FieldSegment field:
                    // The run inside carries the cached result. Word replaces it on open; readers
                    // that do not evaluate fields show this text, so it is a plausible placeholder
                    // rather than empty.
                    paragraph.AppendChild(new SimpleField(new Run(new Text("1")))
                    {
                        Instruction = field.Instruction,
                    });
                    break;
            }
        }

        return paragraph;
    }

    private static JustificationValues ToJustification(HeaderAlignment alignment) => alignment switch
    {
        HeaderAlignment.Center => JustificationValues.Center,
        HeaderAlignment.Right => JustificationValues.Right,
        _ => JustificationValues.Left,
    };
}
