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
/// insert/delete column — addressed by a canonical cell anchor. Test IDs use the DT2xx range.
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

    /// <summary>Insert a rows×cols table and return its created canonical cell anchors (row-major).</summary>
    private static (DocxSession session, string[] cells) NewTable(int rows, int cols,
        string[]? contents = null, int[]? widths = null)
    {
        var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        var r = session.InsertTable(anchor, Position.After, rows, cols,
            new TableInsertOptions { CellContents = contents, ColumnWidths = widths });
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal(rows * cols, r.Created.Count);
        return (session, r.Created.Select(a => a.Id).ToArray());
    }

    // ─── Grid readers (the assertion vocabulary for merge tests) ─────────

    private static XElement Cell(XElement tbl, int row, int cell) =>
        tbl.Elements(W + "tr").ElementAt(row).Elements(W + "tc").ElementAt(cell);

    /// <summary>Cells per row — with merges this is the row's cell count, not the grid width.</summary>
    private static int[] CellCounts(XElement tbl) =>
        tbl.Elements(W + "tr").Select(tr => tr.Elements(W + "tc").Count()).ToArray();

    private static int[] GridCols(XElement tbl) =>
        tbl.Element(W + "tblGrid")!.Elements(W + "gridCol")
            .Select(g => (int)g.Attribute(W + "w")!).ToArray();

    private static int Span(XElement tc) =>
        (int?)tc.Element(W + "tcPr")?.Element(W + "gridSpan")?.Attribute(W + "val") ?? 1;

    /// <summary>null = no vertical merge, "restart" = the merge's lead cell, "continue" = a
    /// continuation (Word writes it as a bare w:vMerge).</summary>
    private static string? VMerge(XElement tc)
    {
        var vm = tc.Element(W + "tcPr")?.Element(W + "vMerge");
        return vm is null ? null : (string?)vm.Attribute(W + "val") ?? "continue";
    }

    private static int CellWidth(XElement tc) =>
        (int?)tc.Element(W + "tcPr")?.Element(W + "tcW")?.Attribute(W + "w") ?? 0;

    private static string CellText(XElement tc) =>
        string.Concat(tc.Descendants(W + "t").Select(t => (string)t));

    private static string AllText(byte[] docxBytes) =>
        string.Concat(DocumentXml(docxBytes).Descendants(W + "t").Select(t => (string)t));

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
    public void DT214b_SetTableRowOptions_WritesHeightAndPageSplitPolicy()
    {
        var (session, cells) = NewTable(2, 2);
        var on = session.SetTableRowOptions(cells[0], new TableRowOptions
        {
            RepeatHeader = true,
            AllowBreakAcrossPages = false,
            HeightTwips = 480,
            HeightRule = TableRowHeightRule.AtLeast,
        });
        Assert.True(on.Success, on.Error?.Message);

        var trPr = SingleTable(session).Elements(W + "tr").First().Element(W + "trPr")!;
        Assert.NotNull(trPr.Element(W + "cantSplit"));
        Assert.NotNull(trPr.Element(W + "tblHeader"));
        Assert.Equal("480", (string)trPr.Element(W + "trHeight")!.Attribute(W + "val"));
        Assert.Equal("atLeast", (string)trPr.Element(W + "trHeight")!.Attribute(W + "hRule"));
        AssertSchemaValid(session.Save());

        var off = session.SetTableRowOptions(cells[0], new TableRowOptions
        {
            AllowBreakAcrossPages = true,
            HeightTwips = 0,
        });
        Assert.True(off.Success, off.Error?.Message);
        trPr = SingleTable(session).Elements(W + "tr").First().Element(W + "trPr")!;
        Assert.Null(trPr.Element(W + "cantSplit"));
        Assert.Null(trPr.Element(W + "trHeight"));
        Assert.NotNull(trPr.Element(W + "tblHeader"));
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT214c_SetTableRowOptions_NegativeHeightIsRejected()
    {
        var (session, cells) = NewTable(1, 1);
        var result = session.SetTableRowOptions(cells[0],
            new TableRowOptions { HeightTwips = -1 });
        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.InvalidTableStyling, result.Error!.Code);
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

    // ─── Cell merge / unmerge (issue #340 Stage B) ───────────────────────

    [Fact]
    public void DT220_MergeCells_Horizontal_WritesGridSpanSumsWidthAndKeepsContent()
    {
        var (session, cells) = NewTable(3, 3,
            contents: new[] { "A", "B", "C", "d", "e", "f", "g", "h", "i" },
            widths: new[] { 2000, 3000, 4000 });

        var r = session.MergeCells(cells[0], rowSpan: 1, colSpan: 3);
        Assert.True(r.Success, r.Error?.Message);
        // Content is preserved, while the two absorbed canonical cell identities are invalidated.
        Assert.Equal(2, r.Removed.Count);
        Assert.All(r.Removed, anchor => Assert.Equal("tc", anchor.Kind));
        Assert.Empty(r.Created);
        Assert.Equal(cells[0], Assert.Single(r.Modified).Id);

        var tbl = SingleTable(session);
        Assert.Equal(new[] { 1, 3, 3 }, CellCounts(tbl));
        Assert.Equal(3, Span(Cell(tbl, 0, 0)));
        Assert.Null(VMerge(Cell(tbl, 0, 0)));                  // no vertical extent
        Assert.Equal(9000, CellWidth(Cell(tbl, 0, 0)));        // 2000 + 3000 + 4000
        Assert.Equal("ABC", CellText(Cell(tbl, 0, 0)));
        Assert.Equal(new[] { 2000, 3000, 4000 }, GridCols(tbl)); // the grid itself is untouched
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT221_MergeCells_Vertical_WritesRestartAndEmptiedContinuations()
    {
        var (session, cells) = NewTable(3, 2,
            contents: new[] { "A", "x", "B", "y", "C", "z" });

        var r = session.MergeCells(cells[0], rowSpan: 3, colSpan: 1);
        Assert.True(r.Success, r.Error?.Message);
        // The absorbed text moved into the lead cell and every canonical cell shell survives.
        Assert.Empty(r.Removed);
        Assert.Empty(r.Created);

        var tbl = SingleTable(session);
        Assert.Equal(new[] { 2, 2, 2 }, CellCounts(tbl)); // a vertical merge removes no cells
        Assert.Equal(new[] { "restart", "continue", "continue" },
            Enumerable.Range(0, 3).Select(i => VMerge(Cell(tbl, i, 0))).ToArray());
        Assert.All(Enumerable.Range(0, 3), i => Assert.Equal(1, Span(Cell(tbl, i, 0))));
        Assert.Equal("ABC", CellText(Cell(tbl, 0, 0)));
        Assert.Equal(new[] { "", "" },
            new[] { CellText(Cell(tbl, 1, 0)), CellText(Cell(tbl, 2, 0)) });
        // The continuation cells still hold exactly one w:p — CT_Tc requires a block child.
        Assert.All(new[] { Cell(tbl, 1, 0), Cell(tbl, 2, 0) },
            tc => Assert.Single(tc.Elements(W + "p")));
        Assert.Equal(new[] { "x", "y", "z" },
            Enumerable.Range(0, 3).Select(i => CellText(Cell(tbl, i, 1))).ToArray());
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT222_MergeCells_Rectangle_SpansBothAxesAndLeavesOuterRowsAlone()
    {
        var (session, cells) = NewTable(3, 3,
            contents: new[] { "A", "B", "c", "D", "E", "f", "g", "h", "i" });

        var r = session.MergeCells(cells[0], rowSpan: 2, colSpan: 2);
        Assert.True(r.Success, r.Error?.Message);

        var tbl = SingleTable(session);
        Assert.Equal(new[] { 2, 2, 3 }, CellCounts(tbl));
        Assert.Equal(2, Span(Cell(tbl, 0, 0)));
        Assert.Equal("restart", VMerge(Cell(tbl, 0, 0)));
        Assert.Equal(2, Span(Cell(tbl, 1, 0)));
        Assert.Equal("continue", VMerge(Cell(tbl, 1, 0)));
        Assert.Equal("ABDE", CellText(Cell(tbl, 0, 0)));
        Assert.Equal("", CellText(Cell(tbl, 1, 0)));
        // Column 2 and row 2 are untouched.
        Assert.Equal(new[] { "c", "f" },
            new[] { CellText(Cell(tbl, 0, 1)), CellText(Cell(tbl, 1, 1)) });
        Assert.Equal("ghi", string.Concat(Enumerable.Range(0, 3).Select(c => CellText(Cell(tbl, 2, c)))));
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT223_MergeCells_ContentModes_DiscardDropsAndRejectRefuses()
    {
        var (discard, dCells) = NewTable(1, 2, contents: new[] { "keep", "drop" });
        var dr = discard.MergeCells(dCells[0], 1, 2,
            new TableMergeOptions { Content = TableMergeContent.Discard });
        Assert.True(dr.Success, dr.Error?.Message);
        Assert.Equal(dCells[1], Assert.Single(dr.Removed).Id); // the dropped cell's paragraph
        Assert.Equal("keep", CellText(Cell(SingleTable(discard), 0, 0)));
        AssertSchemaValid(discard.Save());

        var (reject, rCells) = NewTable(1, 2, contents: new[] { "keep", "occupied" });
        var refused = reject.MergeCells(rCells[0], 1, 2,
            new TableMergeOptions { Content = TableMergeContent.Reject });
        Assert.False(refused.Success);
        Assert.Equal(EditErrorCode.InvalidTableMerge, refused.Error!.Code);
        Assert.Equal(2, CellCounts(SingleTable(reject))[0]); // refused, not half-applied

        // Reject only guards non-empty cells: an empty neighbour merges fine.
        var (empty, eCells) = NewTable(1, 2, contents: new[] { "keep", "" });
        Assert.True(empty.MergeCells(eCells[0], 1, 2,
            new TableMergeOptions { Content = TableMergeContent.Reject }).Success);
    }

    [Fact]
    public void DT224_MergeCells_RejectsOutOfRangeAndDegenerateRectangles()
    {
        var (session, cells) = NewTable(2, 2);
        foreach (var (rowSpan, colSpan) in new[] { (3, 1), (1, 3), (1, 1), (0, 2), (2, -1) })
        {
            var bad = session.MergeCells(cells[0], rowSpan, colSpan);
            Assert.False(bad.Success, $"{rowSpan}x{colSpan} should be rejected");
            Assert.Equal(EditErrorCode.InvalidTableMerge, bad.Error!.Code);
        }
        Assert.Equal(new[] { 2, 2 }, CellCounts(SingleTable(session)));
    }

    [Fact]
    public void DT225_MergeCells_RejectsRectangleStraddlingAnExistingSpan()
    {
        var (session, cells) = NewTable(3, 3);
        // Row 1 gets a 2-wide span over grid columns 1–2.
        Assert.True(session.MergeCells(cells[4], 1, 2).Success);

        // A 2×2 rectangle over columns 0–1 would cut that span in half.
        var bad = session.MergeCells(cells[0], 2, 2);
        Assert.False(bad.Success);
        Assert.Equal(EditErrorCode.InvalidTableMerge, bad.Error!.Code);
        Assert.Contains("do not tile", bad.Error.Message);

        // Columns 0–2 tile cleanly in every row, so the same anchor merges at full width.
        Assert.True(session.MergeCells(cells[0], 2, 3).Success);
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT226_MergeCells_RejectsRectangleClippingAVerticalMerge()
    {
        var (session, cells) = NewTable(3, 2);
        var merged = session.MergeCells(cells[0], 3, 1);
        Assert.True(merged.Success, merged.Error?.Message);
        // The absorbed cells were already empty, so their paragraphs — and anchors — survive
        // in place as the continuation cells' bodies.
        Assert.Empty(merged.Created);
        Assert.Empty(merged.Removed);

        // Stopping one row short of the run would strand row 2's continuation.
        var clipped = session.MergeCells(cells[0], 2, 1);
        Assert.False(clipped.Success);
        Assert.Equal(EditErrorCode.InvalidTableMerge, clipped.Error!.Code);
        Assert.Contains("continues past", clipped.Error.Message);

        // Starting from a continuation cell is equally invalid.
        var fromContinuation = session.MergeCells(cells[2], 2, 1); // row 1, column 0
        Assert.False(fromContinuation.Success);
        Assert.Equal(EditErrorCode.InvalidTableMerge, fromContinuation.Error!.Code);
        Assert.Contains("started above", fromContinuation.Error.Message);
    }

    [Fact]
    public void DT227_UnmergeCells_Horizontal_RestoresUnitCellsAtGridWidths()
    {
        var (session, cells) = NewTable(2, 3,
            contents: new[] { "A", "B", "C", "d", "e", "f" },
            widths: new[] { 2000, 3000, 4000 });
        Assert.True(session.MergeCells(cells[0], 1, 3).Success);

        var r = session.UnmergeCells(cells[0]);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal(2, r.Created.Count); // the two restored cells' paragraphs

        var tbl = SingleTable(session);
        Assert.Equal(new[] { 3, 3 }, CellCounts(tbl));
        Assert.All(tbl.Descendants(W + "tc"), tc => Assert.Equal(1, Span(tc)));
        Assert.Equal(new[] { 2000, 3000, 4000 },
            Enumerable.Range(0, 3).Select(c => CellWidth(Cell(tbl, 0, c))).ToArray());
        // The merged cell kept the absorbed content; the restored cells start empty.
        Assert.Equal(new[] { "ABC", "", "" },
            Enumerable.Range(0, 3).Select(c => CellText(Cell(tbl, 0, c))).ToArray());
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT228_UnmergeCells_FromAContinuation_UnmergesTheWholeRun()
    {
        var (session, cells) = NewTable(3, 2);
        Assert.True(session.MergeCells(cells[0], 3, 1).Success);

        // Address the middle row's continuation paragraph, not the lead cell — an empty cell's
        // anchor survives the merge unchanged.
        var r = session.UnmergeCells(cells[2]);
        Assert.True(r.Success, r.Error?.Message);

        var tbl = SingleTable(session);
        Assert.All(Enumerable.Range(0, 3), i => Assert.Null(VMerge(Cell(tbl, i, 0))));
        Assert.Empty(tbl.Descendants(W + "vMerge"));
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT229_UnmergeCells_UnmergedCellIsRejected()
    {
        var (session, cells) = NewTable(2, 2);
        var bad = session.UnmergeCells(cells[0]);
        Assert.False(bad.Success);
        Assert.Equal(EditErrorCode.InvalidTableMerge, bad.Error!.Code);

        var bodyP = FirstBodyParagraph(session);
        Assert.Equal(EditErrorCode.TableAnchorMigrationRequired, session.MergeCells(bodyP, 2, 2).Error!.Code);
        Assert.Equal(EditErrorCode.TableAnchorMigrationRequired, session.UnmergeCells(bodyP).Error!.Code);
    }

    [Fact]
    public void DT230_MergeThenUnmerge_RestoresTheGridShape()
    {
        var (session, cells) = NewTable(3, 3, widths: new[] { 2000, 3000, 4000 });
        Assert.True(session.MergeCells(cells[0], 2, 2).Success);
        Assert.True(session.UnmergeCells(cells[0]).Success);

        var tbl = SingleTable(session);
        Assert.Equal(new[] { 3, 3, 3 }, CellCounts(tbl));
        Assert.Equal(new[] { 2000, 3000, 4000 }, GridCols(tbl));
        Assert.Empty(tbl.Descendants(W + "gridSpan"));
        Assert.Empty(tbl.Descendants(W + "vMerge"));
        Assert.All(tbl.Elements(W + "tr"), tr => Assert.Equal(
            new[] { 2000, 3000, 4000 },
            tr.Elements(W + "tc").Select(CellWidth).ToArray()));
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT231_MergeCells_IsUndoableAndRedoable()
    {
        var (session, cells) = NewTable(2, 2, contents: new[] { "A", "B", "c", "d" });
        Assert.True(session.MergeCells(cells[0], 1, 2).Success);
        Assert.Equal("AB", CellText(Cell(SingleTable(session), 0, 0)));

        Assert.True(session.Undo());
        var tbl = SingleTable(session);
        Assert.Equal(new[] { 2, 2 }, CellCounts(tbl));
        Assert.Empty(tbl.Descendants(W + "gridSpan"));
        Assert.Equal(new[] { "A", "B" },
            new[] { CellText(Cell(tbl, 0, 0)), CellText(Cell(tbl, 0, 1)) });

        Assert.True(session.Redo());
        Assert.Equal(2, Span(Cell(SingleTable(session), 0, 0)));
    }

    // ─── Span-aware row / column CRUD ────────────────────────────────────

    [Fact]
    public void DT232_InsertTableRow_InsideAVerticalMerge_ExtendsIt()
    {
        var (session, cells) = NewTable(3, 2);
        Assert.True(session.MergeCells(cells[0], 3, 1).Success);

        // cells[1] is row 0's unmerged right-hand cell; inserting below it lands the new row
        // inside the run spanning column 0.
        var r = session.InsertTableRow(cells[1], Position.After);
        Assert.True(r.Success, r.Error?.Message);

        var tbl = SingleTable(session);
        Assert.Equal(4, tbl.Elements(W + "tr").Count());
        Assert.Equal(new[] { "restart", "continue", "continue", "continue" },
            Enumerable.Range(0, 4).Select(i => VMerge(Cell(tbl, i, 0))).ToArray());
        Assert.Null(VMerge(Cell(tbl, 1, 1))); // the new row's other cell is an ordinary cell
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT233_InsertTableRow_OutsideAVerticalMerge_StaysUnmerged()
    {
        var (session, cells) = NewTable(3, 2);
        Assert.True(session.MergeCells(cells[0], 3, 1).Success);

        // Above the merge's restart row: the new row must NOT join the run.
        Assert.True(session.InsertTableRow(cells[1], Position.Before).Success);

        var tbl = SingleTable(session);
        Assert.Equal(new[] { null, "restart", "continue", "continue" },
            Enumerable.Range(0, 4).Select(i => VMerge(Cell(tbl, i, 0))).ToArray());
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT234_InsertTableRow_MirrorsAHorizontalSpan()
    {
        var (session, cells) = NewTable(2, 3);
        Assert.True(session.MergeCells(cells[0], 1, 2).Success); // row 0: [span 2][1]

        var r = session.InsertTableRow(cells[0], Position.After);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal(2, r.Created.Count); // one anchor per cell of the cloned shape

        var tbl = SingleTable(session);
        Assert.Equal(new[] { 2, 2, 3 }, CellCounts(tbl));
        Assert.Equal(2, Span(Cell(tbl, 1, 0)));      // grid shape mirrored…
        Assert.Null(VMerge(Cell(tbl, 1, 0)));        // …but never half of someone's merge
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT235_DeleteTableRow_OfAMergeLeadRow_PromotesTheNextContinuation()
    {
        var (session, cells) = NewTable(3, 2);
        Assert.True(session.MergeCells(cells[0], 3, 1).Success);

        Assert.True(session.DeleteTableRow(cells[1]).Success); // cells[1] is in row 0

        var tbl = SingleTable(session);
        Assert.Equal(2, tbl.Elements(W + "tr").Count());
        Assert.Equal(new[] { "restart", "continue" },
            Enumerable.Range(0, 2).Select(i => VMerge(Cell(tbl, i, 0))).ToArray());
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT236_InsertTableColumn_ThroughASpan_WidensItInsteadOfSplittingIt()
    {
        var (session, cells) = NewTable(3, 3, widths: new[] { 2000, 2000, 2000 });
        Assert.True(session.MergeCells(cells[0], 1, 3).Success); // row 0 spans all three columns

        // Boundary after grid column 0 falls INSIDE row 0's span.
        var r = session.InsertTableColumn(cells[3], Position.After);
        Assert.True(r.Success, r.Error?.Message);
        Assert.Equal(2, r.Created.Count); // rows 1 and 2 gain a cell; row 0 only widens

        var tbl = SingleTable(session);
        Assert.Equal(new[] { 2000, 2000, 2000, 2000 }, GridCols(tbl));
        Assert.Equal(new[] { 1, 4, 4 }, CellCounts(tbl));
        Assert.Equal(4, Span(Cell(tbl, 0, 0)));
        Assert.Equal(8000, CellWidth(Cell(tbl, 0, 0)));
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT237_DeleteTableColumn_ThroughASpan_NarrowsItInsteadOfDroppingTheCell()
    {
        var (session, cells) = NewTable(3, 3, widths: new[] { 2000, 3000, 4000 });
        Assert.True(session.MergeCells(cells[0], 1, 3).Success);

        Assert.True(session.DeleteTableColumn(cells[4]).Success); // grid column 1 (width 3000)

        var tbl = SingleTable(session);
        Assert.Equal(new[] { 2000, 4000 }, GridCols(tbl));
        Assert.Equal(new[] { 1, 2, 2 }, CellCounts(tbl));
        Assert.Equal(2, Span(Cell(tbl, 0, 0)));
        Assert.Equal(6000, CellWidth(Cell(tbl, 0, 0))); // 9000 − 3000
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void DT238_SetColumnWidths_SizesAMergedCellToTheColumnsItSpans()
    {
        var (session, cells) = NewTable(2, 3);
        Assert.True(session.MergeCells(cells[0], 1, 2).Success);

        Assert.True(session.SetColumnWidths(cells[0], new[] { 1000, 2000, 3000 }).Success);

        var tbl = SingleTable(session);
        Assert.Equal(3000, CellWidth(Cell(tbl, 0, 0))); // 1000 + 2000
        Assert.Equal(3000, CellWidth(Cell(tbl, 0, 1)));
        Assert.Equal(new[] { 1000, 2000, 3000 },
            Enumerable.Range(0, 3).Select(c => CellWidth(Cell(tbl, 1, c))).ToArray());
        AssertSchemaValid(session.Save());
    }

    // ─── Merged tables downstream: projection and the diff engine ────────

    [Fact]
    public void DT239_MergedTable_ProjectsOpaquelyAndKeepsItsCellsAddressable()
    {
        var (session, cells) = NewTable(2, 3, contents: new[] { "A", "B", "C", "d", "e", "f" });
        Assert.True(session.MergeCells(cells[0], 2, 2).Success);

        var projection = session.Project();
        // Merged cells disqualify GFM rendering, so the table projects as an opaque block…
        Assert.Contains("```table", projection.Markdown);
        // …whose width is the GRID extent (3), not the merged first row's cell count (2).
        Assert.Contains("rows: 2\ncols: 3", projection.Markdown.Replace("\r\n", "\n"));
        // …while every surviving canonical cell stays individually addressable.
        Assert.Contains(cells[0], projection.AnchorIndex.Keys);
        Assert.Contains(cells[2], projection.AnchorIndex.Keys);
        Assert.True(session.ReplaceCellContent(cells[2], "still editable").Success);
        Assert.Equal("still editable", CellText(Cell(SingleTable(session), 0, 1)));
    }

    [Fact]
    public void DT240_MergedTable_RoundTripsThroughDocxDiff()
    {
        var (session, cells) = NewTable(3, 3,
            contents: new[] { "A", "B", "c", "D", "E", "f", "g", "h", "i" },
            widths: new[] { 2000, 3000, 4000 });
        Assert.True(session.MergeCells(cells[0], 2, 2).Success);
        var left = new WmlDocument("left.docx", session.Save());

        Assert.True(session.ReplaceCellContent(cells[0], "revised headline").Success);
        var right = new WmlDocument("right.docx", session.Save());

        var redline = DocxDiff.Compare(left, right);
        Assert.NotEmpty(DocxDiff.GetRevisions(left, right));
        Assert.Equal(AllText(right.DocumentByteArray),
            AllText(RevisionProcessor.AcceptRevisions(redline).DocumentByteArray));
        Assert.Equal(AllText(left.DocumentByteArray),
            AllText(RevisionProcessor.RejectRevisions(redline).DocumentByteArray));
        AssertSchemaValid(redline.DocumentByteArray);
    }
}
