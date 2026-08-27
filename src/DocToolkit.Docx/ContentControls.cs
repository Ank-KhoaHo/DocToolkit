using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocToolkit;

/// <summary>
/// Looks through Word content controls (<c>w:sdt</c>) to the tables, rows and cells inside them.
/// </summary>
/// <remarks>
/// <para>
/// A content control is a wrapper Word puts around ordinary content, so <c>body.Elements&lt;Table&gt;()</c>
/// simply does not see a table inside one. Three readers hit that blind spot in three different
/// ways, and the symptoms were not equally visible: <c>TableCount</c> answered <b>0</b>,
/// <c>ReadTable(0)</c> returned the table that was physically <i>second</i>, and a wrapped row
/// vanished from an otherwise correct table — which looks exactly like data.
/// </para>
/// <para>
/// <b>This is one class rather than three copies because the alternative already has a name here.</b>
/// <c>DocxEditor</c> and <c>TableRowFinder</c> both need to know what counts as a table, and a
/// second source of truth about that is the drift <c>SetCellValue</c>, <c>SectionPropertiesFactory</c>
/// and <c>ValidateSheetName</c> each exist to prevent. Two readers disagreeing about one document
/// is the defect this closes; adding a third way to answer the question would re-create it.
/// </para>
/// <para>
/// <b>Elements, never Descendants, at every level — this is the constraint that makes it safe.</b>
/// Each method unwraps <i>one control at a time</i> and recurses only through control content, so
/// nesting is untouched: a table inside a cell stays part of that cell rather than becoming a
/// top-level entry, and a nested table's rows are never swept into its container's. A
/// <c>Descendants</c> walk would be shorter and has twice cost this repository real defects —
/// deleted text-box content relocated into an outer paragraph, and rows cloned once per record
/// because a container row was mistaken for a template.
/// </para>
/// </remarks>
internal static class ContentControls
{
    /// <summary>
    /// The tables directly under <paramref name="scope"/>, including any wrapped in one or more
    /// block-level content controls, in document order.
    /// </summary>
    /// <remarks>
    /// Recursion is through <c>SdtBlock</c> content <b>only</b>. Word nests controls, so unwrapping
    /// a single level is not enough — but descending into anything else would change what a
    /// top-level table is.
    /// </remarks>
    public static IEnumerable<Table> Tables(OpenXmlElement scope) =>
        scope.Elements().SelectMany<OpenXmlElement, Table>(child => child switch
        {
            Table table => [table],
            SdtBlock control => control.SdtContentBlock is { } content ? Tables(content) : [],
            _ => [],
        });

    /// <summary>
    /// A table's rows, including any wrapped in a row-level content control, in document order.
    /// </summary>
    public static IEnumerable<TableRow> Rows(Table table) => RowsIn(table);

    /// <summary>
    /// A row's cells, including any wrapped in a cell-level content control, in document order.
    /// </summary>
    /// <remarks>
    /// A dropped cell does not merely go missing: it shifts every cell beside it, so the columns
    /// stop lining up and a caller reading by position gets the wrong value rather than none.
    /// </remarks>
    public static IEnumerable<TableCell> Cells(TableRow row) => CellsIn(row);

    /// <summary>
    /// The paragraphs directly under <paramref name="scope"/>, including any wrapped in one or more
    /// block-level content controls, in document order.
    /// </summary>
    /// <remarks>
    /// The fourth wrapper position, and the one a first pass missed. A control can wrap a
    /// <i>paragraph</i> inside a cell — <c>w:tc &gt; w:sdt &gt; w:p</c> — which is what Word writes
    /// when a Rich Text control is inserted into an empty cell. A reader that unwrapped tables,
    /// rows and cells but not paragraphs still could not see a template marker sitting there.
    /// </remarks>
    public static IEnumerable<Paragraph> Paragraphs(OpenXmlElement scope) =>
        scope.Elements().SelectMany<OpenXmlElement, Paragraph>(child => child switch
        {
            Paragraph paragraph => [paragraph],
            SdtBlock control => control.SdtContentBlock is { } content ? Paragraphs(content) : [],
            _ => [],
        });

    private static IEnumerable<TableRow> RowsIn(OpenXmlElement scope) =>
        scope.Elements().SelectMany<OpenXmlElement, TableRow>(child => child switch
        {
            TableRow row => [row],
            SdtRow control => control.SdtContentRow is { } content ? RowsIn(content) : [],
            _ => [],
        });

    private static IEnumerable<TableCell> CellsIn(OpenXmlElement scope) =>
        scope.Elements().SelectMany<OpenXmlElement, TableCell>(child => child switch
        {
            TableCell cell => [cell],
            SdtCell control => control.SdtContentCell is { } content ? CellsIn(content) : [],
            _ => [],
        });
}
