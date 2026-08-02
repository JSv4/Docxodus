#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Tests for post-insert table editing on <see cref="DocxSession"/>: insert/delete row,
/// insert/delete column — addressed by a cell-paragraph anchor. Test IDs use the DT2xx range.
/// </summary>
public class DocxSessionTableEditTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static XElement DocumentXml(byte[] docxBytes)
    {
        using var ms = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        return doc.MainDocumentPart!.GetXDocument().Root!;
    }

    private static string FirstBodyParagraph(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;

    /// <summary>Insert a rows×cols table and return its created cell-paragraph anchors (row-major).</summary>
    private static (DocxSession session, string[] cells) NewTable(int rows, int cols)
    {
        var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        var r = session.InsertTable(anchor, Position.After, rows, cols);
        Assert.True(r.Success, r.Error?.Message);
        return (session, r.Created.Select(a => a.Id).ToArray());
    }

    private static XElement SingleTable(DocxSession session) =>
        DocumentXml(session.Save()).Descendants(W + "tbl").Single();

    private static void AssertSchemaValid(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        var errors = new OpenXmlValidator().Validate(wDoc)
            .Select(e => $"{e.Path?.XPath}: {e.Description}").ToList();
        Assert.True(errors.Count == 0, "OOXML schema errors:\n" + string.Join("\n", errors));
    }

    [Fact]
    public void DT201_InsertTableRow_After_AddsRowWithSameColumnCount()
    {
        var (session, cells) = NewTable(2, 2); // cells row-major: r0c0,r0c1,r1c0,r1c1
        var r = session.InsertTableRow(cells[0], Position.After); // after row 0
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal(2, r.Created.Count); // the new row's two cell paragraphs

        var tbl = SingleTable(session);
        Assert.Equal(3, tbl.Elements(W + "tr").Count());
        Assert.All(tbl.Elements(W + "tr"), tr => Assert.Equal(2, tr.Elements(W + "tc").Count()));
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT202_InsertTableColumn_After_AddsColumnToEveryRow()
    {
        var (session, cells) = NewTable(2, 2);
        var r = session.InsertTableColumn(cells[0], Position.After); // after column 0
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal(2, r.Created.Count); // one new cell per row

        var tbl = SingleTable(session);
        Assert.Equal(3, tbl.Element(W + "tblGrid")!.Elements(W + "gridCol").Count());
        Assert.All(tbl.Elements(W + "tr"), tr => Assert.Equal(3, tr.Elements(W + "tc").Count()));
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT203_DeleteTableRow_RemovesTheRow()
    {
        var (session, cells) = NewTable(3, 2);
        var r = session.DeleteTableRow(cells[2]); // a cell in row 1
        Assert.True(r.Success, r.Error?.Message);

        var tbl = SingleTable(session);
        Assert.Equal(2, tbl.Elements(W + "tr").Count());
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT204_DeleteTableColumn_RemovesTheColumnFromEveryRow()
    {
        var (session, cells) = NewTable(2, 3);
        var r = session.DeleteTableColumn(cells[1]); // column 1
        Assert.True(r.Success, r.Error?.Message);

        var tbl = SingleTable(session);
        Assert.Equal(2, tbl.Element(W + "tblGrid")!.Elements(W + "gridCol").Count());
        Assert.All(tbl.Elements(W + "tr"), tr => Assert.Equal(2, tr.Elements(W + "tc").Count()));
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT205_DeleteLastRow_RemovesTheWholeTable()
    {
        var (session, cells) = NewTable(1, 2);
        var r = session.DeleteTableRow(cells[0]);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Empty(DocumentXml(session.Save()).Descendants(W + "tbl"));
    }

    [Fact]
    public void DT206_DeleteLastColumn_RemovesTheWholeTable()
    {
        var (session, cells) = NewTable(2, 1);
        var r = session.DeleteTableColumn(cells[0]);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Empty(DocumentXml(session.Save()).Descendants(W + "tbl"));
    }

    [Fact]
    public void DT207_InsertRow_NonCellAnchor_IsRejected()
    {
        var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var bodyP = FirstBodyParagraph(session);
        var r = session.InsertTableRow(bodyP, Position.After);
        Assert.False(r.Success); // a body paragraph is not in a table
    }

    // ─── Table styling (issue #315 Stage A) ─────────────────────────────

    [Fact]
    public void DT208_SetColumnWidths_RewritesGridCellsAndTableWidth()
    {
        var (session, cells) = NewTable(2, 3);
        var r = session.SetColumnWidths(cells[0], new[] { 4000, 2000, 1000 });
        Assert.True(r.Success, r.Error?.Message);

        var tbl = SingleTable(session);
        Assert.Equal(new[] { "4000", "2000", "1000" },
            tbl.Element(W + "tblGrid")!.Elements(W + "gridCol")
                .Select(g => (string)g.Attribute(W + "w")!).ToArray());
        Assert.All(tbl.Elements(W + "tr"), tr => Assert.Equal(
            new[] { "4000", "2000", "1000" },
            tr.Elements(W + "tc")
                .Select(tc => (string)tc.Element(W + "tcPr")!.Element(W + "tcW")!.Attribute(W + "w")!)
                .ToArray()));

        var tblW = tbl.Element(W + "tblPr")!.Element(W + "tblW")!;
        Assert.Equal("7000", (string)tblW.Attribute(W + "w"));
        Assert.Equal("dxa", (string)tblW.Attribute(W + "type"));
        Assert.Equal("fixed",
            (string)tbl.Element(W + "tblPr")!.Element(W + "tblLayout")!.Attribute(W + "type"));
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT209_SetColumnWidths_WrongCountOrNonPositive_IsRejected()
    {
        var (session, cells) = NewTable(2, 3);
        var wrongCount = session.SetColumnWidths(cells[0], new[] { 4000, 2000 });
        Assert.False(wrongCount.Success);
        Assert.Equal(EditErrorCode.InvalidTableStyling, wrongCount.Error!.Code);

        var nonPositive = session.SetColumnWidths(cells[0], new[] { 4000, 0, 1000 });
        Assert.False(nonPositive.Success);
        Assert.Equal(EditErrorCode.InvalidTableStyling, nonPositive.Error!.Code);
    }

    [Fact]
    public void DT210_SetTableBorders_OutsideScope_LeavesInsideEdgesUntouched()
    {
        var (session, cells) = NewTable(2, 2);
        var r = session.SetTableBorders(cells[0], new TableBorderSpec
        {
            Scope = TableBorderScope.Outside,
            Style = "double",
            Size = 12,
            Color = "FF0000",
        });
        Assert.True(r.Success, r.Error?.Message);

        var borders = SingleTable(session).Element(W + "tblPr")!.Element(W + "tblBorders")!;
        foreach (var edge in new[] { "top", "left", "bottom", "right" })
        {
            var e = borders.Element(W + edge)!;
            Assert.Equal("double", (string)e.Attribute(W + "val"));
            Assert.Equal("12", (string)e.Attribute(W + "sz"));
            Assert.Equal("FF0000", (string)e.Attribute(W + "color"));
        }
        // InsertTable wrote thin single inside edges; the outside-scoped op must not touch them.
        foreach (var edge in new[] { "insideH", "insideV" })
        {
            var e = borders.Element(W + edge)!;
            Assert.Equal("single", (string)e.Attribute(W + "val"));
            Assert.Equal("4", (string)e.Attribute(W + "sz"));
        }
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT211_SetTableBorders_StyleNone_WritesExplicitNoneEdges()
    {
        var (session, cells) = NewTable(2, 2);
        var r = session.SetTableBorders(cells[0], new TableBorderSpec { Style = "none" });
        Assert.True(r.Success, r.Error?.Message);

        var borders = SingleTable(session).Element(W + "tblPr")!.Element(W + "tblBorders")!;
        foreach (var edge in new[] { "top", "left", "bottom", "right", "insideH", "insideV" })
            Assert.Equal("none", (string)borders.Element(W + edge)!.Attribute(W + "val"));
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT212_SetCellShading_RowScope_ShadesEveryCellInTheRow()
    {
        var (session, cells) = NewTable(2, 2); // row-major: r0c0,r0c1,r1c0,r1c1
        var r = session.SetCellShading(cells[0], "#d9d9d9", TableShadingScope.Row);
        Assert.True(r.Success, r.Error?.Message);

        var rows = SingleTable(session).Elements(W + "tr").ToList();
        Assert.All(rows[0].Elements(W + "tc"), tc =>
        {
            var shd = tc.Element(W + "tcPr")!.Element(W + "shd")!;
            Assert.Equal("clear", (string)shd.Attribute(W + "val"));
            Assert.Equal("D9D9D9", (string)shd.Attribute(W + "fill")); // normalized upper, '#' stripped
        });
        Assert.All(rows[1].Elements(W + "tc"),
            tc => Assert.Null(tc.Element(W + "tcPr")!.Element(W + "shd")));
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT213_SetCellShading_NullFill_ClearsAndBadFillIsRejected()
    {
        var (session, cells) = NewTable(1, 2);
        Assert.True(session.SetCellShading(cells[0], "336699").Success);
        Assert.True(session.SetCellShading(cells[0], null).Success);

        var tc0 = SingleTable(session).Descendants(W + "tc").First();
        Assert.Null(tc0.Element(W + "tcPr")!.Element(W + "shd"));

        var bad = session.SetCellShading(cells[0], "not-a-color");
        Assert.False(bad.Success);
        Assert.Equal(EditErrorCode.InvalidTableStyling, bad.Error!.Code);
    }

    [Fact]
    public void DT214_SetRepeatHeaderRow_TogglesTrPrTblHeader()
    {
        var (session, cells) = NewTable(3, 2);
        var on = session.SetRepeatHeaderRow(cells[0], true);
        Assert.True(on.Success, on.Error?.Message);

        var firstRow = SingleTable(session).Elements(W + "tr").First();
        Assert.NotNull(firstRow.Element(W + "trPr")?.Element(W + "tblHeader"));
        AssertSchemaValid(session.Save());

        var off = session.SetRepeatHeaderRow(cells[0], false);
        Assert.True(off.Success, off.Error?.Message);
        firstRow = SingleTable(session).Elements(W + "tr").First();
        Assert.Null(firstRow.Element(W + "trPr")); // emptied trPr is removed entirely
    }

    [Fact]
    public void DT215_TableStyling_NonCellAnchor_IsRejected()
    {
        var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var bodyP = FirstBodyParagraph(session);
        Assert.False(session.SetColumnWidths(bodyP, new[] { 1000 }).Success);
        Assert.False(session.SetTableBorders(bodyP).Success);
        Assert.False(session.SetCellShading(bodyP, "D9D9D9").Success);
        Assert.False(session.SetRepeatHeaderRow(bodyP, true).Success);
    }

    [Fact]
    public void DT216_TableStyling_IsUndoable()
    {
        var (session, cells) = NewTable(2, 2);
        Assert.True(session.SetCellShading(cells[0], "D9D9D9", TableShadingScope.Row).Success);
        Assert.True(session.SetRepeatHeaderRow(cells[0], true).Success);

        Assert.True(session.Undo()); // repeat-header off
        Assert.True(session.Undo()); // shading gone

        var tbl = SingleTable(session);
        Assert.Empty(tbl.Descendants(W + "shd"));
        Assert.Empty(tbl.Descendants(W + "tblHeader"));

        Assert.True(session.Redo());
        Assert.NotEmpty(SingleTable(session).Descendants(W + "shd"));
    }
}
