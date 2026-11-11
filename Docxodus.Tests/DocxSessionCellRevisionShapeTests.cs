using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #1404: the two table shapes bulk accept/reject refused. Expected geometry is what desktop
/// Word 2024 (16.0.20326) wrote for the same markup: an accepted deleted cell gives its grid columns
/// to the cell before it, or to the cell after it when it starts the row, and a cell's own content
/// control leaves the row with it.
/// </summary>
public class DocxSessionCellRevisionShapeTests
{
    private const string Stamp = "w:author=\"Reviewer A\" w:date=\"2026-03-02T10:00:00Z\"";

    private static string Cell(string text, string? mark = null, int span = 1, string stamp = Stamp) =>
        "<w:tc><w:tcPr><w:tcW w:w=\"" + 2400 * span + "\" w:type=\"dxa\"/>"
        + (span > 1 ? $"<w:gridSpan w:val=\"{span}\"/>" : "")
        + (mark is null ? "" : $"<w:{mark} w:id=\"{text[0] - 'A' + 1}\" {stamp}/>")
        + $"</w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";

    private static string Sdt(int id, params string[] cells) =>
        $"<w:sdt><w:sdtPr><w:id w:val=\"{id}\"/></w:sdtPr><w:sdtContent>{string.Concat(cells)}</w:sdtContent></w:sdt>";

    private static string Row(params string[] cells) => "<w:tr>" + string.Concat(cells) + "</w:tr>";

    private static byte[] Table(int columns, params string[] rows)
    {
        var xml = $"<w:document xmlns:w=\"{W.w.NamespaceName}\"><w:body><w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>"
            + "<w:tblGrid>" + string.Concat(Enumerable.Repeat("<w:gridCol w:w=\"2400\"/>", columns)) + "</w:tblGrid>"
            + string.Concat(rows) + "</w:tbl><w:p/>"
            + "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:body></w:document>";
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.PutXDocument(XDocument.Parse(xml));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(new XDocument(new XElement(W.styles)));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(new XDocument(new XElement(W.settings)));
        }
        return stream.ToArray();
    }

    /// <summary>Each row as its cells' text, grid span and wrapping content control, for example
    /// <c>sdt7:A×2 C</c>; rows are joined with <c> / </c>. Fails on a leftover revision marker, a
    /// content control left without a cell, or schema-invalid output.</summary>
    private static string Shape(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).Select(error => error.Description));
        var table = doc.MainDocumentPart!.GetXDocument().Descendants(W.tbl).First();
        Assert.Empty(table.Descendants().Where(Docxodus.Internal.RevisionOps.IsRecognizedRevisionMarker));
        Assert.DoesNotContain(table.Descendants(W.sdt), sdt => !sdt.Descendants(W.tc).Any());
        return string.Join(" / ", table.Elements(W.tr).Select(row => string.Join(" ", row.Descendants(W.tc).Select(cell =>
        {
            int span = (int?)cell.Element(W.tcPr)?.Element(W.gridSpan)?.Attribute(W.val) ?? 1;
            var control = cell.Ancestors(W.sdt).FirstOrDefault()?.Element(W.sdtPr)?.Element(W.id)?.Attribute(W.val);
            return (control is null ? "" : $"sdt{control.Value}:") + cell.Value + (span > 1 ? "×" + span : "");
        }))));
    }

    private static string Resolve(byte[] bytes, bool accept)
    {
        using var session = new DocxSession(bytes);
        Assert.All(session.ListRevisions(), entry =>
            Assert.True(entry.ResolutionStatus == RevisionResolutionStatus.Supported, entry.Diagnostic?.Code));
        var result = accept ? session.AcceptAllRevisions() : session.RejectAllRevisions();
        Assert.True(result.Success, result.Error?.Message);
        Assert.Empty(session.ListRevisions());
        return Shape(session.Save());
    }

    public static TheoryData<string, int, string, string, string> DeletedCells => new()
    {
        // shape, grid columns, row, accepted, rejected
        { "middle (control)", 3, Row(Cell("A"), Cell("B", "cellDel"), Cell("C")), "A×2 C", "A B C" },
        { "last (control)", 3, Row(Cell("A"), Cell("B"), Cell("C", "cellDel")), "A B×2", "A B C" },
        { "first", 3, Row(Cell("A", "cellDel"), Cell("B"), Cell("C")), "B×2 C", "A B C" },
        { "first two", 3, Row(Cell("A", "cellDel"), Cell("B", "cellDel"), Cell("C")), "C×3", "A B C" },
        { "first and last", 3, Row(Cell("A", "cellDel"), Cell("B"), Cell("C", "cellDel")), "B×3", "A B C" },
        { "first, before a merged cell", 4, Row(Cell("A", "cellDel"), Cell("B", span: 2), Cell("D")), "B×3 D", "A B×2 D" },
    };

    [Theory]
    [MemberData(nameof(DeletedCells))]
    public void DeletedCell_GivesItsColumnsToTheCellBeforeOrElseTheCellAfter(
        string shape, int columns, string row, string accepted, string rejected)
    {
        var bytes = Table(columns, row, Row(Enumerable.Range(0, columns).Select(i => Cell("x")).ToArray()));
        string untouched = " / " + string.Join(" ", Enumerable.Repeat("x", columns));

        Assert.True(accepted + untouched == Resolve(bytes, accept: true), shape);
        Assert.True(rejected + untouched == Resolve(bytes, accept: false), shape);
    }

    public static TheoryData<string, string, string, string> WrappedCells => new()
    {
        // shape, row, accepted, rejected
        {
            "deleted middle", Row(Sdt(7, Cell("A")), Sdt(8, Cell("B", "cellDel")), Sdt(9, Cell("C"))),
            "sdt7:A×2 sdt9:C", "sdt7:A sdt8:B sdt9:C"
        },
        {
            "deleted first", Row(Sdt(7, Cell("A", "cellDel")), Sdt(8, Cell("B")), Sdt(9, Cell("C"))),
            "sdt8:B×2 sdt9:C", "sdt7:A sdt8:B sdt9:C"
        },
        {
            "inserted middle", Row(Sdt(7, Cell("A")), Sdt(8, Cell("B", "cellIns")), Sdt(9, Cell("C"))),
            "sdt7:A sdt8:B sdt9:C", "sdt7:A sdt9:C"
        },
        {
            // A control around two cells outlives the one that goes.
            "deleted beside a cell in the same control", Row(Cell("A"), Sdt(7, Cell("B", "cellDel"), Cell("C"))),
            "A×2 sdt7:C", "A sdt7:B sdt7:C"
        },
    };

    [Theory]
    [MemberData(nameof(WrappedCells))]
    public void CellInsideAContentControl_ResolvesLikeABareCellAndTakesItsOwnControlWithIt(
        string shape, string row, string accepted, string rejected)
    {
        var plain = Row(Cell("x"), Cell("x"), Cell("x"));
        var bytes = Table(3, row, plain);

        Assert.True(accepted + " / x x x" == Resolve(bytes, accept: true), shape);
        Assert.True(rejected + " / x x x" == Resolve(bytes, accept: false), shape);

        // The content controls change nothing about the geometry: the same row without them
        // resolves to the same cells.
        var bare = Table(3, System.Text.RegularExpressions.Regex.Replace(row, "<w:sdt>.*?<w:sdtContent>|</w:sdtContent></w:sdt>", ""), plain);
        static string WithoutControls(string cells) => System.Text.RegularExpressions.Regex.Replace(cells, @"sdt\d+:", "");
        Assert.Equal(WithoutControls(accepted) + " / x x x", Resolve(bare, accept: true));
        Assert.Equal(WithoutControls(rejected) + " / x x x", Resolve(bare, accept: false));
    }

    [Fact]
    public void WrappedCellsDeletedDownAColumn_AreOneRevisionAsBareCellsAre()
    {
        static byte[] Column(bool wrapped)
        {
            string Mid(string text) => wrapped ? Sdt(text[0], Cell(text, "cellDel")) : Cell(text, "cellDel");
            return Table(3, Row(Cell("A"), Mid("B"), Cell("C")), Row(Cell("D"), Mid("E"), Cell("F")));
        }

        using var bare = new DocxSession(Column(wrapped: false));
        using var wrapped = new DocxSession(Column(wrapped: true));

        Assert.Single(bare.ListRevisions());
        Assert.Single(wrapped.ListRevisions());
        Assert.Equal("A×2 C / D×2 F", Resolve(Column(wrapped: true), accept: true));

        // The cells that went are reported removed either way.
        var removed = bare.AcceptAllRevisions().Removed;
        Assert.NotEmpty(removed);
        Assert.Equal(removed.Count, wrapped.AcceptAllRevisions().Removed.Count);
    }

    [Theory]
    [InlineData(Stamp)]
    [InlineData("w:author=\"Reviewer B\" w:date=\"2026-03-03T10:00:00Z\"")]
    public void RowWithEveryCellDeleted_IsStillRefused(string secondStamp)
    {
        // Control: no cell survives to take the columns, whether one revision deletes the row's
        // cells or two do. Word removes the row; the session does not resolve that yet, and must
        // refuse before changing anything rather than leave a row with no cells.
        var bytes = Table(2, Row(Cell("A", "cellDel"), Cell("B", "cellDel", stamp: secondStamp)), Row(Cell("x"), Cell("x")));
        using var session = new DocxSession(bytes);

        Assert.All(session.ListRevisions(), entry => Assert.Equal("unabsorbable_cell_deletion", entry.Diagnostic?.Code));
        Assert.False(session.AcceptAllRevisions().Success);
        Assert.Equal(2, MainTable(session.Save()).Descendants(W.cellDel).Count());
    }

    [Fact]
    public void AbsorbingCell_TakesTheDeletedCellsWidth()
    {
        // Word widens the absorbing cell's w:tcW by the deleted cell's. With the first column
        // deleted from every row no row uses the grid's first column on its own any more, and
        // Word lays the cells out by these widths: stale ones made both cells equally wide.
        var firstColumn = Table(3,
            Row(Cell("A", "cellDel"), Cell("B"), Cell("C")), Row(Cell("D", "cellDel"), Cell("E"), Cell("F")));
        using var session = new DocxSession(firstColumn);
        Assert.True(session.AcceptAllRevisions().Success);
        Assert.Equal(new[] { "4800", "2400", "4800", "2400" }, Widths(session.Save()));

        // RevisionProcessor gives the columns to the cell before a deleted one, and the width with them.
        var middle = new WmlDocument("middle.docx", Table(3, Row(Cell("A"), Cell("B", "cellDel"), Cell("C"))));
        Assert.Equal(new[] { "4800", "2400" }, Widths(RevisionProcessor.AcceptRevisions(middle).DocumentByteArray));

        static string?[] Widths(byte[] bytes) =>
            MainTable(bytes).Descendants(W.tcW).Select(width => (string?)width.Attribute(W._w)).ToArray();
    }

    private static XElement MainTable(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        return doc.MainDocumentPart!.GetXDocument().Descendants(W.tbl).First();
    }
}
