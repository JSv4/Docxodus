using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issues #1404 and #1688: cells that leave a row, by an accepted <c>w:cellDel</c> or a rejected
/// <c>w:cellIns</c>. Expected geometry is what desktop Word 2024 (16.0.20326) wrote for the same
/// markup: the leaving cell's grid columns and width go to the cell before it, or to the cell after
/// it when it starts the row; a cell's own content control leaves with it; a row left without a
/// cell is removed, and a table left without a row with it. <see cref="DocxSession"/> and
/// <see cref="RevisionProcessor"/> share the rule, so every case is asserted on both.
/// </summary>
public class DocxSessionCellRevisionShapeTests
{
    private const string Stamp = "w:author=\"Reviewer A\" w:date=\"2026-03-02T10:00:00Z\"";
    private const string OtherStamp = "w:author=\"Reviewer B\" w:date=\"2026-03-03T10:00:00Z\"";

    private static string Cell(string text, string? mark = null, int span = 1, string stamp = Stamp) =>
        "<w:tc><w:tcPr><w:tcW w:w=\"" + 2400 * span + "\" w:type=\"dxa\"/>"
        + (span > 1 ? $"<w:gridSpan w:val=\"{span}\"/>" : "")
        + (mark is null ? "" : $"<w:{mark} w:id=\"{text[0] - 'A' + 1}\" {stamp}/>")
        + $"</w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";

    private static string Sdt(int id, params string[] cells) =>
        $"<w:sdt><w:sdtPr><w:id w:val=\"{id}\"/></w:sdtPr><w:sdtContent>{string.Concat(cells)}</w:sdtContent></w:sdt>";

    private static string Row(params string[] cells) => "<w:tr>" + string.Concat(cells) + "</w:tr>";

    private static string Plain(int columns) => Row(Enumerable.Repeat(Cell("x"), columns).ToArray());

    private static string Tbl(int columns, params string[] rows) =>
        "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid>"
        + string.Concat(Enumerable.Repeat("<w:gridCol w:w=\"2400\"/>", columns)) + "</w:tblGrid>" + string.Concat(rows) + "</w:tbl>";

    private static byte[] Table(int columns, params string[] rows) => Body(Tbl(columns, rows) + "<w:p/>");

    private static byte[] Body(string blocks)
    {
        var xml = $"<w:document xmlns:w=\"{W.w.NamespaceName}\"><w:body>{blocks}"
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

    private static XElement MainBody(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(doc).Select(error => error.Description));
        return doc.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }

    /// <summary>The first table: each row as its cells' text, grid span and wrapping content
    /// control, for example <c>sdt7:A×2 C</c>, rows joined with <c> / </c>; <c>no table</c> when
    /// there is none. Fails on a leftover revision marker, a content control left without a cell,
    /// a cell whose width is not its columns' (every column here is 2400 wide), or schema-invalid
    /// output.</summary>
    private static string Shape(byte[] bytes)
    {
        var body = MainBody(bytes);
        Assert.Empty(body.Descendants().Where(Docxodus.Internal.RevisionOps.IsRecognizedRevisionMarker));
        if (body.Element(W.tbl) is not { } table) return "no table";
        Assert.DoesNotContain(table.Descendants(W.sdt), sdt => !sdt.Descendants(W.tc).Any());
        return string.Join(" / ", table.Elements(W.tr).Select(row => string.Join(" ", WordprocessingMLUtil.RowCells(row).Select(cell =>
        {
            int span = (int?)cell.Element(W.tcPr)?.Element(W.gridSpan)?.Attribute(W.val) ?? 1;
            Assert.Equal(2400 * span, (int?)cell.Element(W.tcPr)?.Element(W.tcW)?.Attribute(W._w));
            var control = cell.Ancestors(W.sdt).FirstOrDefault()?.Element(W.sdtPr)?.Element(W.id)?.Attribute(W.val);
            return (control is null ? "" : $"sdt{control.Value}:") + string.Concat(cell.Elements(W.p).Select(p => p.Value))
                + (span > 1 ? "×" + span : "");
        }))));
    }

    /// <summary>Resolves every revision through the session, and through
    /// <see cref="RevisionProcessor"/>, which must give the same table.</summary>
    private static string Resolve(byte[] bytes, bool accept)
    {
        using var session = new DocxSession(bytes);
        Assert.All(session.ListRevisions(), entry =>
            Assert.True(entry.ResolutionStatus == RevisionResolutionStatus.Supported, entry.Diagnostic?.Code));
        var result = accept ? session.AcceptAllRevisions() : session.RejectAllRevisions();
        Assert.True(result.Success, result.Error?.Message);
        Assert.Empty(session.ListRevisions());
        string shape = Shape(session.Save());

        var document = new WmlDocument("cells.docx", bytes);
        var processed = accept ? RevisionProcessor.AcceptRevisions(document) : RevisionProcessor.RejectRevisions(document);
        Assert.Equal(shape, Shape(processed.DocumentByteArray));
        return shape;
    }

    public static TheoryData<string, int, string, string, string> LeavingCells => new()
    {
        // shape, grid columns, row with ? for the mark, the row once the marked cells have left, the row with them
        { "middle", 3, Row(Cell("A"), Cell("B", "?"), Cell("C")), "A×2 C", "A B C" },
        { "last", 3, Row(Cell("A"), Cell("B"), Cell("C", "?")), "A B×2", "A B C" },
        { "first", 3, Row(Cell("A", "?"), Cell("B"), Cell("C")), "B×2 C", "A B C" },
        { "first two", 3, Row(Cell("A", "?"), Cell("B", "?"), Cell("C")), "C×3", "A B C" },
        { "first and last", 3, Row(Cell("A", "?"), Cell("B"), Cell("C", "?")), "B×3", "A B C" },
        { "first, before a merged cell", 4, Row(Cell("A", "?"), Cell("B", span: 2), Cell("D")), "B×3 D", "A B×2 D" },
        { "each in its own control, middle", 3, Row(Sdt(7, Cell("A")), Sdt(8, Cell("B", "?")), Sdt(9, Cell("C"))), "sdt7:A×2 sdt9:C", "sdt7:A sdt8:B sdt9:C" },
        { "each in its own control, first", 3, Row(Sdt(7, Cell("A", "?")), Sdt(8, Cell("B")), Sdt(9, Cell("C"))), "sdt8:B×2 sdt9:C", "sdt7:A sdt8:B sdt9:C" },
        // A control around two cells outlives the one that goes.
        { "beside a cell in the same control", 3, Row(Cell("A"), Sdt(7, Cell("B", "?"), Cell("C"))), "A×2 sdt7:C", "A sdt7:B sdt7:C" },
    };

    [Theory]
    [MemberData(nameof(LeavingCells))]
    public void LeavingCell_GivesItsColumnsAndWidthToTheCellBeforeOrElseTheCellAfter(
        string shape, int columns, string row, string gone, string kept)
    {
        string untouched = " / " + string.Join(" ", Enumerable.Repeat("x", columns));
        var deleted = Table(columns, row.Replace("w:?", "w:cellDel"), Plain(columns));
        var inserted = Table(columns, row.Replace("w:?", "w:cellIns"), Plain(columns));

        Assert.True(gone + untouched == Resolve(deleted, accept: true), shape + ": accepted deletion");
        Assert.True(kept + untouched == Resolve(deleted, accept: false), shape + ": rejected deletion");
        Assert.True(gone + untouched == Resolve(inserted, accept: false), shape + ": rejected insertion");
        Assert.True(kept + untouched == Resolve(inserted, accept: true), shape + ": accepted insertion");
    }

    [Theory]
    [InlineData("cellDel")]
    [InlineData("cellIns")]
    public void ColumnThatLeavesEveryRow_LeavesNoRowShort_SoNoCellIsWidened(string mark)
    {
        // The rows decide whether a row is short. With the first column gone from every row the
        // table is one column narrower and whole. (Word 2024 merges each leaving cell into its
        // neighbour here too, which gives a first column twice as wide.)
        var bytes = Table(3, Row(Cell("A", mark), Cell("B"), Cell("C")), Row(Cell("D", mark), Cell("E"), Cell("F")));

        Assert.Equal("B C / E F", Resolve(bytes, accept: mark == "cellDel"));
    }

    [Theory]
    [InlineData("RP034-Deleted-Cells.docx", true, 1)]
    [InlineData("RP035-Inserted-Cells.docx", false, 1)]
    public void WordAuthoredMergeAndSplit_ResolveToOneCellAsWideAsTheTable(string fixture, bool accept, int cells)
    {
        // Word files written before it stopped tracking cell structure: three cells merged into one
        // (two marked deleted), and one cell split into three (two marked inserted). Accepting the
        // merge and rejecting the split both leave one cell over all three columns, 9016 twips wide,
        // which is also what Word 2024 makes of the merge. Its recorded cell width must not be
        // added to again, and the split's recorded span must not be widened a second time.
        var bytes = File.ReadAllBytes(Path.GetFullPath(Path.Combine(System.AppContext.BaseDirectory, "../../../../TestFiles/RP", fixture)));
        var document = new WmlDocument(fixture, bytes);
        var processed = accept ? RevisionProcessor.AcceptRevisions(document) : RevisionProcessor.RejectRevisions(document);
        using var session = new DocxSession(bytes);
        Assert.True((accept ? session.AcceptAllRevisions() : session.RejectAllRevisions()).Success);

        foreach (var resolved in new[] { processed.DocumentByteArray, session.Save() })
        {
            using var stream = new MemoryStream(resolved);
            using var doc = WordprocessingDocument.Open(stream, false);
            var firstRow = WordprocessingMLUtil.RowCells(doc.MainDocumentPart!.GetXDocument().Descendants(W.tr).First()).ToList();
            Assert.Equal(cells, firstRow.Count);
            Assert.Equal("3", (string?)firstRow[0].Element(W.tcPr)!.Element(W.gridSpan)?.Attribute(W.val));
            Assert.Equal("9016", (string?)firstRow[0].Element(W.tcPr)!.Element(W.tcW)!.Attribute(W._w));
        }
    }

    [Theory]
    [InlineData("cellDel", Stamp)]
    [InlineData("cellDel", OtherStamp)]
    [InlineData("cellIns", Stamp)]
    [InlineData("cellIns", OtherStamp)]
    public void RowWhoseEveryCellLeaves_IsRemoved(string mark, string secondStamp)
    {
        // One revision marks the whole row, or two reviewers mark it between them.
        var bytes = Table(3, Plain(3).Replace(">x<", ">w<"),
            Row(Cell("A", mark), Cell("B", mark, stamp: secondStamp), Cell("C", mark, stamp: secondStamp)), Plain(3));
        bool leaves = mark == "cellDel";

        Assert.Equal("w w w / x x x", Resolve(bytes, accept: leaves));
        Assert.Equal("w w w / A B C / x x x", Resolve(bytes, accept: !leaves));
    }

    [Fact]
    public void RowLeftWithoutACell_OneRevisionAtATime_GoesWithTheLastOne()
    {
        var bytes = Table(2, Row(Cell("A", "cellDel"), Cell("B", "cellDel", stamp: OtherStamp)), Plain(2));
        using var session = new DocxSession(bytes);

        Assert.True(session.AcceptRevision(session.ListRevisions()[0].Id).Success);
        Assert.Equal(2, MainBody(session.Save()).Descendants(W.tr).Count());
        Assert.True(session.AcceptRevision(Assert.Single(session.ListRevisions()).Id).Success);

        Assert.Equal("x x", Shape(session.Save()));
        Assert.True(session.Undo());
        Assert.Single(session.ListRevisions());
    }

    [Theory]
    [InlineData("cellDel")]
    [InlineData("cellIns")]
    public void TableLeftWithoutARow_IsRemoved(string mark)
    {
        var onlyRow = Row(Cell("A", mark), Cell("B", mark));
        bool leaves = mark == "cellDel";

        // Between paragraphs: the table goes and the paragraphs stay.
        var body = Body("<w:p><w:r><w:t>Before</w:t></w:r></w:p>" + Tbl(2, onlyRow) + "<w:p><w:r><w:t>After</w:t></w:r></w:p>");
        Assert.Equal("no table", Resolve(body, accept: leaves));
        Assert.Equal("A B", Resolve(body, accept: !leaves));

        // As a cell's only content beside its closing paragraph: the cell keeps that paragraph.
        var nested = Body("<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr><w:tblGrid><w:gridCol w:w=\"2400\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"2400\" w:type=\"dxa\"/></w:tcPr>" + Tbl(2, onlyRow) + "<w:p/></w:tc></w:tr></w:tbl><w:p/>");
        Assert.Equal("", Resolve(nested, accept: leaves));
    }

    [Theory]
    [InlineData(Position.After)]
    [InlineData(Position.Before)]
    public void RejectingTheSessionsOwnInsertedColumn_RestoresTheOriginalTable(Position position)
    {
        // Control for the widening rule. A tracked InsertTableColumn records the old grid beside its
        // inserted cells; rejecting restores that grid first, the remaining cells fill it, and so no
        // neighbour takes the inserted cells' columns a second time.
        var original = Table(3, Row(Cell("A"), Cell("B"), Cell("C")), Row(Cell("D"), Cell("E"), Cell("F")));
        byte[] tracked;
        using (var session = new DocxSession(original, new DocxSessionSettings { RevisionAuthor = "Reviewer A", TrackedChanges = TrackedChangeMode.RenderInline }))
        {
            var first = session.Project().AnchorIndex.Values.First(a => a.Anchor.Scope == "body" && a.Anchor.Kind == "tc");
            Assert.True(session.InsertTableColumn(first.Anchor.Id, position).Success);
            tracked = session.Save();
        }

        Assert.Equal("A B C / D E F", Resolve(tracked, accept: false));
        Assert.Equal(3, MainBody(RevisionProcessor.RejectRevisions(new WmlDocument("t.docx", tracked)).DocumentByteArray)
            .Descendants(W.gridCol).Count());
    }

    [Fact]
    public void ComparisonThatRemovedAColumn_AcceptsToTheRevisedTableWithoutWidening()
    {
        // The redline narrows the grid and marks the surplus cell of each row deleted. The cells that
        // stay already fill the narrowed grid, so accepting must not widen one of them over it.
        static WmlDocument Document(params string[][] rows) => new("d.docx", Table(rows[0].Length,
            rows.Select(cells => Row(cells.Select(text => Cell(text)).ToArray())).ToArray()));
        var redline = DocxCompare.Compare(
            Document(new[] { "A", "B", "C" }, new[] { "D", "E", "F" }), Document(new[] { "B", "C" }, new[] { "E", "F" }));

        Assert.Equal("B C / E F", Resolve(redline.DocumentByteArray, accept: true));
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
        Assert.Equal("A C / D F", Resolve(Column(wrapped: true), accept: true));

        // The cells that went are reported removed either way.
        var removed = bare.AcceptAllRevisions().Removed;
        Assert.NotEmpty(removed);
        Assert.Equal(removed.Count, wrapped.AcceptAllRevisions().Removed.Count);
    }
}
