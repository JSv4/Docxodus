using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Docxodus.Internal;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Docxodus.Tests;

// Issue #1737: range boundaries (bookmarks, tracked custom XML ranges) may sit between the cells of a
// row or between the rows of a table. HTML admits only cells in a row and rows in a table, so what such
// a boundary renders must land inside a cell, exactly where the same boundary inside that cell renders.
public class TableRangeBoundaryTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string Rev = "w:author=\"Reviewer\" w:date=\"2026-01-05T09:00:00Z\"";
    private const string Leading = "<w:bookmarkStart w:id=\"1\" w:name=\"Ledger\"/>";
    private const string BetweenCells = $"<w:customXmlInsRangeStart w:id=\"2\" {Rev}/>";
    private const string EndOfRow = "<w:customXmlInsRangeEnd w:id=\"2\"/>";
    private const string BetweenRows = $"<w:customXmlDelRangeStart w:id=\"3\" {Rev}/>";
    private const string Trailing = "<w:customXmlDelRangeEnd w:id=\"3\"/><w:bookmarkEnd w:id=\"1\"/>";

    private static string Cell(string text, string before = "", string after = "") =>
        $"<w:tc><w:tcPr><w:tcW w:w=\"2400\" w:type=\"dxa\"/></w:tcPr>{before}<w:p><w:r><w:t>{text}</w:t></w:r></w:p>{after}</w:tc>";

    // The same three-row ledger, with its boundaries between cells and rows or inside the adjacent cells.
    // trPr applies to the first and last rows, which receive the boundaries; the narrow first grid
    // column is the one a w:gridBefore row omits.
    private static byte[] Ledger(bool betweenCells, string trPr)
    {
        string Row(string rowPr, string cells, string after = "") => $"<w:tr><w:trPr>{rowPr}</w:trPr>{cells}{after}</w:tr>";
        var rows = betweenCells
            ? Leading + Row(trPr, Cell("Alder") + BetweenCells + Cell("Ash"), EndOfRow) + Row("", Cell("Birch") + Cell("Beech"))
              + BetweenRows + Row(trPr, Cell("Cedar") + Cell("Cypress")) + Trailing
            : Row(trPr, Cell("Alder", before: Leading) + Cell("Ash", before: BetweenCells, after: EndOfRow))
              + Row("", Cell("Birch") + Cell("Beech")) + Row(trPr, Cell("Cedar", before: BetweenRows) + Cell("Cypress", after: Trailing));
        var content = $"<w:document xmlns:w=\"{W}\"><w:body><w:tbl><w:tblPr><w:tblW w:w=\"5400\" w:type=\"dxa\"/></w:tblPr>"
            + "<w:tblGrid><w:gridCol w:w=\"600\"/><w:gridCol w:w=\"2400\"/><w:gridCol w:w=\"2400\"/></w:tblGrid>"
            + rows + "</w:tbl><w:p/></w:body></w:document>";
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            using var xml = new MemoryStream(Encoding.UTF8.GetBytes(content));
            document.AddMainDocumentPart().FeedData(xml);
        }
        return stream.ToArray();
    }

    private static XElement Table(bool betweenCells, string trPr) =>
        WmlToHtmlConverter.ConvertToHtml(new WmlDocument("Ledger.docx", Ledger(betweenCells, trPr)),
            new WmlToHtmlConverterSettings { RenderTrackedChanges = true, IncludeRevisionMetadata = true })
        .Descendants(Xhtml.table).Single();

    private static int Markers(XElement html) =>
        html.Descendants().Count(e => ((string?)e.Attribute("class"))?.Contains("rev-cxml-marker") == true);

    [Theory]
    [InlineData("")]
    [InlineData("<w:trHeight w:val=\"900\" w:hRule=\"exact\"/>")]
    [InlineData("<w:gridBefore w:val=\"1\"/><w:wBefore w:w=\"600\" w:type=\"dxa\"/>")]
    public void BoundaryBetweenCellsOrRowsRendersInTheAdjacentCell(string trPr)
    {
        var table = Table(betweenCells: true, trPr);

        Assert.All(table.Elements(), e => Assert.Contains(e.Name, new[] { Xhtml.colgroup, Xhtml.tr }));
        Assert.All(table.Elements(Xhtml.tr).Elements(), e => Assert.Equal(Xhtml.td, e.Name));
        Assert.Equal(4, Markers(table));
        Assert.Single(table.Descendants(Xhtml.a), a => (string?)a.Attribute("id") == "Ledger");
        Assert.Equal(Table(betweenCells: false, trPr).ToString(), table.ToString());
    }

    [Fact]
    public void RowRenderedAloneKeepsItsBoundariesInItsCells()
    {
        var bytes = Ledger(betweenCells: true, "");
        using var session = new DocxSession(bytes);
        var firstRow = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "tr").Anchor.Id;

        var row = XElement.Parse(HtmlConversionOps.RenderBlockHtml(bytes, firstRow,
            new HtmlConversionOptions { RenderTrackedChanges = true }));

        Assert.All(row.Elements(), e => Assert.Equal("td", e.Name.LocalName));
        Assert.Equal(2, Markers(row));
    }
}
