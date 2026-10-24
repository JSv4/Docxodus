#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus;
using Docxodus.Tests.Ir;
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

    private static void AssertSchemaValid(byte[] bytes) =>
        PackageValidation.AssertValid(bytes, DocumentFormat.OpenXml.FileFormatVersions.Office2007);

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
    // ─── Untracked column edits under a pending grid change (issue #1342) ─────────

    /// <summary>A 2x2 table whose grid carries a pending <c>w:tblGridChange</c>: a real one (the
    /// old widths differ) or Word's no-op one (the old grid is the live grid).</summary>
    private static byte[] TableWithPendingGridChange(bool real, bool pending = true, bool unaligned = false)
    {
        string Cell(string text) =>
            $"<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";
        var oldGrid = real
            ? "<w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"1000\"/>"
            : "<w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/>";
        if (unaligned) oldGrid += "<w:gridCol w:w=\"2000\"/>";
        return IrTestDocuments.FromBodyXml(
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/>"
            + (pending ? $"<w:tblGridChange w:id=\"5\"><w:tblGrid>{oldGrid}</w:tblGrid></w:tblGridChange>" : "") + "</w:tblGrid>"
            + $"<w:tr>{Cell("A1")}{Cell("B1")}</w:tr><w:tr>{Cell("A2")}{Cell("B2")}</w:tr></w:tbl>"
            + "<w:p><w:r><w:t>after</w:t></w:r></w:p>").DocumentByteArray;
    }

    private static byte[] PendingGridChangeFixture(string fixture) => fixture switch
    {
        "synthetic_real" => TableWithPendingGridChange(real: true),
        "synthetic_no_op" => TableWithPendingGridChange(real: false),
        _ => File.ReadAllBytes(Path.GetFullPath(Path.Combine(
            System.AppContext.BaseDirectory, "../../../../TestFiles", fixture))),
    };

    /// <summary>Every row's grid span (omitted columns, then each cell's span) next to the
    /// grid's column count, for the first table.</summary>
    private static (int GridColumns, int[] RowSpans) GridAndRowSpans(byte[] bytes)
    {
        var tbl = DocumentXml(bytes).Descendants(W + "tbl").First();
        int Omitted(XElement tr, string name) =>
            (int?)tr.Element(W + "trPr")?.Element(W + name)?.Attribute(W + "val") ?? 0;
        return (tbl.Element(W + "tblGrid")!.Elements(W + "gridCol").Count(),
            tbl.Elements(W + "tr")
                .Select(tr => Omitted(tr, "gridBefore") + Omitted(tr, "gridAfter") + tr.Elements(W + "tc").Sum(Span))
                .ToArray());
    }

    /// <summary>An untracked column insertion or deletion is mirrored in the old grid a pending
    /// <c>w:tblGridChange</c> keeps, so rejecting that change, all at once or revision by
    /// revision, restores a grid exactly as wide as the rows, and the widths can be set again.</summary>
    [Theory]
    [InlineData("FA/RevTracking/021-TableGridChange.docx", "insert")]
    [InlineData("FA/RevTracking/021-TableGridChange.docx", "delete")]
    [InlineData("RP/RP029-Table-Row-Props-Change.docx", "insert")]
    [InlineData("RP/RP029-Table-Row-Props-Change.docx", "delete")]
    [InlineData("synthetic_real", "insert")]
    [InlineData("synthetic_real", "delete")]
    [InlineData("synthetic_no_op", "insert")]
    [InlineData("synthetic_no_op", "delete")]
    public void DT265_UntrackedColumnEdit_KeepsAPendingGridChangeAsWideAsTheRows(string fixture, string operation)
    {
        var baseline = PendingGridChangeFixture(fixture);
        byte[] edited;
        using (var session = new DocxSession(baseline))
        {
            var cell = session.Project().AnchorIndex.Values.First(a => a.Anchor.Kind == "tc").Anchor.Id;
            var result = operation == "insert"
                ? session.InsertTableColumn(cell, Position.After)
                : session.DeleteTableColumn(cell);
            Assert.True(result.Success, result.Error?.Message);
            edited = session.Save();

            Assert.True(session.Undo());
            Assert.True(XNode.DeepEquals(DocumentXml(baseline), DocumentXml(session.Save())));
        }
        foreach (var version in new[] { DocumentFormat.OpenXml.FileFormatVersions.Office2019,
                     DocumentFormat.OpenXml.FileFormatVersions.Office2021 })
            Assert.Subset(PackageValidation.ErrorKeys(baseline, version, PackageValidation.Keys.IdDescriptionXPath),
                PackageValidation.ErrorKeys(edited, version, PackageValidation.Keys.IdDescriptionXPath));

        foreach (var resolution in new[] { "reject_all", "reject_each", "accept_all" })
        {
            using var session = new DocxSession(edited);
            if (resolution == "reject_each")
                while (session.ListRevisions().FirstOrDefault() is { } revision)
                    Assert.True(session.RejectRevision(revision.Id).Success, revision.Id);
            else
                Assert.True((resolution == "accept_all"
                    ? session.AcceptAllRevisions() : session.RejectAllRevisions()).Success);

            var (gridColumns, rowSpans) = GridAndRowSpans(session.Save());
            Assert.All(rowSpans, span => Assert.Equal(gridColumns, span));
            var cell = session.Project().AnchorIndex.Values.First(a => a.Anchor.Kind == "tc").Anchor.Id;
            var widths = session.SetColumnWidths(cell, Enumerable.Repeat(1000, gridColumns).ToArray());
            Assert.True(widths.Success, $"{resolution}: {widths.Error?.Message}");
        }
    }

    /// <summary>The mirrored column is the old grid's copy of the column the new one was cloned
    /// from, at the same position: rejecting the grid change gives the new column that column's
    /// old width, as it gives every other column its own.</summary>
    [Fact]
    public void DT266_UntrackedColumnInsert_RejectGivesTheNewColumnItsSourcesOldWidth()
    {
        using var session = new DocxSession(TableWithPendingGridChange(real: true));
        var cell = session.Project().AnchorIndex.Values.First(a => a.Anchor.Kind == "tc").Anchor.Id;
        Assert.True(session.InsertTableColumn(cell, Position.After).Success);
        Assert.Equal(new[] { 2000, 2000, 2000 }, GridCols(SingleTable(session)));

        Assert.True(session.RejectAllRevisions().Success);

        Assert.Equal(new[] { 3000, 3000, 1000 }, GridCols(SingleTable(session)));
    }

    private static byte[] PendingColumnGeometry(string shape, bool incompatible = false)
    {
        string Cell(string text, string properties = "<w:tcW w:w=\"2000\" w:type=\"dxa\"/>") =>
            $"<w:tc><w:tcPr>{properties}</w:tcPr><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";
        const string stamp = "w:id=\"6\" w:author=\"Prior Reviewer\" w:date=\"2026-01-01T00:00:00Z\"";
        var current = "<w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/>";
        var prior = "<w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"1000\"/>";
        string rows;
        if (shape == "merged")
        {
            const string width = "<w:tcW w:w=\"4000\" w:type=\"dxa\"/><w:gridSpan w:val=\"2\"/>";
            var previousWidth = incompatible ? width.Replace("w:val=\"2\"", "w:val=\"1\"") : width;
            var properties = width + "<w:shd w:val=\"clear\" w:fill=\"00AA00\"/>"
                + $"<w:tcPrChange {stamp}><w:tcPr>{previousWidth}<w:shd w:val=\"clear\" w:fill=\"AA0000\"/></w:tcPr></w:tcPrChange>";
            var previousRow = incompatible ? "<w:trPr><w:trPrChange w:id=\"7\" w:author=\"Prior Reviewer\" w:date=\"2026-01-01T00:00:00Z\"><w:trPr><w:gridAfter w:val=\"1\"/></w:trPr></w:trPrChange></w:trPr>" : "";
            rows = $"<w:tr>{previousRow}{Cell("Merged", properties)}</w:tr><w:tr>{Cell("A")}{Cell("B")}</w:tr>";
        }
        else
        {
            current += "<w:gridCol w:w=\"2000\"/>";
            prior += "<w:gridCol w:w=\"2000\"/>";
            var rowProperties = "<w:gridBefore w:val=\"1\"/><w:trHeight w:val=\"480\"/>"
                + $"<w:trPrChange {stamp}><w:trPr><w:gridBefore w:val=\"{(incompatible ? 0 : 1)}\"/>"
                + (incompatible ? "<w:gridAfter w:val=\"1\"/>" : "")
                + "<w:trHeight w:val=\"240\"/></w:trPr></w:trPrChange>";
            rows = $"<w:tr><w:trPr>{rowProperties}</w:trPr>{Cell("A")}{Cell("B")}</w:tr>"
                + $"<w:tr>{Cell("C")}{Cell("D")}{Cell("E")}</w:tr>";
        }
        return IrTestDocuments.FromBodyXml("<w:tbl><w:tblPr/><w:tblGrid>" + current
            + "<w:tblGridChange w:id=\"5\"><w:tblGrid>" + prior + "</w:tblGrid></w:tblGridChange></w:tblGrid>"
            + rows + "</w:tbl><w:p><w:r><w:t>After</w:t></w:r></w:p>").DocumentByteArray;
    }

    [Theory]
    [InlineData("merged", "insert")]
    [InlineData("merged", "delete")]
    [InlineData("omitted", "insert")]
    [InlineData("omitted", "delete")]
    public void UntrackedColumnEdit_KeepsPendingCellAndRowGeometryConsistent(string shape, string operation)
    {
        var baseline = PendingColumnGeometry(shape);
        AssertSchemaValid(baseline);
        byte[] edited;
        using (var session = new DocxSession(baseline))
        {
            var cells = session.Project().AnchorIndex.Values.Where(a => a.Anchor.Kind == "tc").ToArray();
            var anchor = cells[shape == "merged" ? 1 : 2].Anchor.Id;
            Assert.True((operation == "insert" ? session.InsertTableColumn(anchor,
                shape == "merged" ? Position.After : Position.Before) : session.DeleteTableColumn(anchor)).Success);
            edited = session.Save();
            AssertSchemaValid(edited);
        }
        foreach (var bulk in new[] { false, true })
        {
            using var session = new DocxSession(edited);
            if (bulk)
                Assert.True(session.RejectAllRevisions().Success);
            else
                while (session.ListRevisions().FirstOrDefault() is { } revision)
                    Assert.True(session.RejectRevision(revision.Id).Success);
            var (columns, spans) = GridAndRowSpans(session.Save());
            Assert.All(spans, span => Assert.Equal(columns, span));
            var firstRow = SingleTable(session).Elements(W + "tr").First();
            if (shape == "merged")
            {
                var properties = firstRow.Element(W + "tc")!.Element(W + "tcPr")!;
                Assert.Equal(operation == "insert" ? 7000 : 1000, (int?)properties.Element(W + "tcW")!.Attribute(W + "w"));
                Assert.Equal("AA0000", (string?)properties.Element(W + "shd")!.Attribute(W + "fill"));
            }
            else
                Assert.Equal(240, (int?)firstRow.Element(W + "trPr")!.Element(W + "trHeight")!.Attribute(W + "val"));
            AssertSchemaValid(session.Save());
        }
    }

    [Theory]
    [InlineData("merged", "insert")]
    [InlineData("merged", "delete")]
    [InlineData("omitted", "insert")]
    [InlineData("omitted", "delete")]
    public void UntrackedColumnEdit_RefusesAnIncompatiblePendingLayoutBeforeHistory(string shape, string operation)
    {
        var baseline = PendingColumnGeometry(shape, incompatible: true);
        AssertSchemaValid(baseline);
        using var session = new DocxSession(baseline);
        var cells = session.Project().AnchorIndex.Values.Where(a => a.Anchor.Kind == "tc").ToArray();
        var anchor = cells[shape == "merged" ? 1 : 2].Anchor.Id;
        var before = DocumentXml(session.Save());
        var version = session.Version;
        var result = operation == "insert" ? session.InsertTableColumn(anchor,
            shape == "merged" ? Position.After : Position.Before) : session.DeleteTableColumn(anchor);
        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.UnresolvedStructuralRevision, result.Error!.Code);
        Assert.Equal(version, session.Version);
        Assert.True(XNode.DeepEquals(before, DocumentXml(session.Save())));
        Assert.False(session.Undo());
    }

    [Theory]
    [InlineData("insert")]
    [InlineData("delete")]
    public void UntrackedColumnEdit_WithoutPendingPropertiesStaysRevisionFree(string operation)
    {
        var baseline = TableWithPendingGridChange(real: false, pending: false);
        using var session = new DocxSession(baseline);
        var anchor = session.Project().AnchorIndex.Values.First(a => a.Anchor.Kind == "tc").Anchor.Id;
        Assert.True((operation == "insert" ? session.InsertTableColumn(anchor, Position.After) : session.DeleteTableColumn(anchor)).Success);
        Assert.Empty(session.ListRevisions());
        var edited = DocumentXml(session.Save());
        Assert.True(session.Undo());
        Assert.True(XNode.DeepEquals(DocumentXml(baseline), DocumentXml(session.Save())));
        Assert.True(session.Redo());
        Assert.True(XNode.DeepEquals(edited, DocumentXml(session.Save())));
        AssertSchemaValid(session.Save());
    }

    [Theory]
    [InlineData(Position.Before, 3000)]
    [InlineData(Position.After, 1000)]
    public void UntrackedNewColumn_CopiesAUnitCellBaselineFromAMergedReference(Position position, int oldWidth)
    {
        using var session = new DocxSession(PendingColumnGeometry("merged"));
        var anchor = session.Project().AnchorIndex.Values.First(a => a.Anchor.Kind == "tc").Anchor.Id;
        Assert.True(session.InsertTableColumn(anchor, position).Success);
        var row = SingleTable(session).Elements(W + "tr").First();
        var inserted = position == Position.Before ? row.Elements(W + "tc").First() : row.Elements(W + "tc").Last();
        var previous = inserted.Element(W + "tcPr")!.Element(W + "tcPrChange")!.Element(W + "tcPr")!;
        Assert.Null(previous.Element(W + "gridSpan"));
        Assert.Equal(oldWidth, (int?)previous.Element(W + "tcW")!.Attribute(W + "w"));
        Assert.True(session.RejectAllRevisions().Success);
        var (columns, spans) = GridAndRowSpans(session.Save());
        Assert.Equal(3, columns);
        Assert.All(spans, span => Assert.Equal(columns, span));
        AssertSchemaValid(session.Save());
    }

    [Theory]
    [InlineData(Position.Before)]
    [InlineData(Position.After)]
    public void UntrackedColumnInsert_UsesTheLastColumnsPreviousWidth(Position position)
    {
        using var session = new DocxSession(TableWithPendingGridChange(real: true));
        var anchor = session.Project().AnchorIndex.Values.Where(a => a.Anchor.Kind == "tc").Skip(1).First().Anchor.Id;
        Assert.True(session.InsertTableColumn(anchor, position).Success);
        Assert.True(session.RejectAllRevisions().Success);
        Assert.Equal(new[] { 3000, 1000, 1000 }, GridCols(SingleTable(session)));
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void UntrackedColumnDelete_RemovesTheSameLastColumnFromThePreviousGrid()
    {
        using var session = new DocxSession(TableWithPendingGridChange(real: true));
        var anchor = session.Project().AnchorIndex.Values.Where(a => a.Anchor.Kind == "tc").Skip(1).First().Anchor.Id;
        Assert.True(session.DeleteTableColumn(anchor).Success);
        Assert.True(session.RejectAllRevisions().Success);
        Assert.Equal(new[] { 3000 }, GridCols(SingleTable(session)));
        AssertSchemaValid(session.Save());
    }

    [Theory]
    [InlineData("insert")]
    [InlineData("delete")]
    public void UntrackedColumnEdit_RequiresAnAlignedPreviousGrid(string operation)
    {
        var baseline = TableWithPendingGridChange(real: true, unaligned: true);
        AssertSchemaValid(baseline);
        using var session = new DocxSession(baseline);
        var anchor = session.Project().AnchorIndex.Values.First(a => a.Anchor.Kind == "tc").Anchor.Id;
        var before = DocumentXml(session.Save());
        var result = operation == "insert" ? session.InsertTableColumn(anchor, Position.After) : session.DeleteTableColumn(anchor);
        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.UnresolvedStructuralRevision, result.Error!.Code);
        Assert.True(XNode.DeepEquals(before, DocumentXml(session.Save())));
        Assert.False(session.Undo());
        Assert.True(session.AcceptAllRevisions().Success);
        anchor = session.Project().AnchorIndex.Values.First(a => a.Anchor.Kind == "tc").Anchor.Id;
        Assert.True((operation == "insert" ? session.InsertTableColumn(anchor, Position.After) : session.DeleteTableColumn(anchor)).Success);
        AssertSchemaValid(session.Save());
    }

    [Fact]
    public void UntrackedColumnDelete_RefusesAUnitCellWhosePreviousLayoutSpannedTwoColumns()
    {
        const string stamp = "w:author=\"Prior Reviewer\" w:date=\"2026-01-01T00:00:00Z\"";
        var baseline = IrTestDocuments.FromBodyXml(
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"2000\"/>"
            + "<w:tblGridChange w:id=\"5\"><w:tblGrid><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid></w:tblGridChange></w:tblGrid>"
            + $"<w:tr><w:trPr><w:gridBefore w:val=\"1\"/><w:trPrChange w:id=\"7\" {stamp}><w:trPr><w:gridBefore w:val=\"0\"/></w:trPr></w:trPrChange></w:trPr>"
            + $"<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/><w:tcPrChange w:id=\"6\" {stamp}><w:tcPr><w:tcW w:w=\"4000\" w:type=\"dxa\"/><w:gridSpan w:val=\"2\"/></w:tcPr></w:tcPrChange></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>"
            + "<w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>").DocumentByteArray;
        AssertSchemaValid(baseline);
        using var session = new DocxSession(baseline);
        var anchor = session.Project().AnchorIndex.Values.First(a => a.Anchor.Kind == "tc").Anchor.Id;
        var before = DocumentXml(session.Save());
        var result = session.DeleteTableColumn(anchor);
        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.UnresolvedStructuralRevision, result.Error!.Code);
        Assert.True(XNode.DeepEquals(before, DocumentXml(session.Save())));
        Assert.False(session.Undo());
    }

    // ─── Block insertion beside a table part (issue #1331) ─────────

    /// <summary>A one-cell table, a block content control and a paragraph holding an inline
    /// content control: every anchor kind a block insertion might be handed.</summary>
    private static byte[] BlockAnchorDocument() =>
        IrTestDocuments.FromBodyXml(
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>"
            + "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc></w:tr>"
            + "</w:tbl>"
            + "<w:sdt><w:sdtPr/><w:sdtContent><w:p><w:r><w:t>block control</w:t></w:r></w:p></w:sdtContent></w:sdt>"
            + "<w:customXml w:element=\"record\"><w:p><w:r><w:t>custom block control</w:t></w:r></w:p></w:customXml>"
            + "<w:p><w:r><w:t>before </w:t></w:r><w:sdt><w:sdtPr/><w:sdtContent><w:r><w:t>inline control</w:t></w:r></w:sdtContent></w:sdt></w:p>"
            + "<w:p><w:r><w:t>after</w:t></w:r></w:p>").DocumentByteArray;

    /// <summary>The anchor of the element <paramref name="kind"/> names in
    /// <see cref="BlockAnchorDocument"/>.</summary>
    private static string BlockAnchor(DocxSession session, string kind)
    {
        var anchors = session.Project().AnchorIndex.Values;
        return kind switch
        {
            "cell_paragraph" => anchors.First(a => a.Anchor.Kind == "p" && a.TextPreview.StartsWith("A1")).Anchor.Id,
            "body_paragraph" => anchors.First(a => a.Anchor.Kind == "p" && a.TextPreview.StartsWith("after")).Anchor.Id,
            "block_sdt" => anchors.First(a => a.Anchor.Kind == "sdt" && a.TextPreview.StartsWith("block")).Anchor.Id,
            "custom_paragraph" => anchors.First(a => a.Anchor.Kind == "p" && a.TextPreview.StartsWith("custom")).Anchor.Id,
            "inline_sdt" => anchors.First(a => a.Anchor.Kind == "sdt" && a.TextPreview.StartsWith("inline")).Anchor.Id,
            _ => anchors.First(a => a.Anchor.Kind == kind).Anchor.Id,
        };
    }

    private static EditResult InsertBlock(DocxSession session, string operation, string anchor, Position position) =>
        operation switch
        {
            "paragraph" => session.InsertParagraph(anchor, position, "new text"),
            "rule" => session.InsertHorizontalRule(anchor, position),
            "table" => session.InsertTable(anchor, position, 1, 1),
            _ => throw new System.ArgumentOutOfRangeException(nameof(operation)),
        };

    /// <summary>A cell, row, grid column or inline content control has no block siblings, so a
    /// paragraph, rule or table placed beside one would be schema-invalid (a <c>w:p</c> directly
    /// inside a <c>w:tr</c>). The insertion is refused before anything is written or recorded.</summary>
    [Theory]
    [InlineData("paragraph", "tc", false)]
    [InlineData("paragraph", "tc", true)]
    [InlineData("paragraph", "tr", false)]
    [InlineData("paragraph", "col", false)]
    [InlineData("paragraph", "inline_sdt", false)]
    [InlineData("rule", "tc", false)]
    [InlineData("rule", "tc", true)]
    [InlineData("rule", "tr", false)]
    [InlineData("table", "tc", false)]
    [InlineData("table", "col", false)]
    public void DT260_BlockInsertion_BesideANonBlock_IsRefused(string operation, string kind, bool tracked)
    {
        var baseline = BlockAnchorDocument();
        foreach (var position in new[] { Position.Before, Position.After })
        {
            using var session = new DocxSession(baseline, new DocxSessionSettings
            {
                TrackedChanges = tracked ? TrackedChangeMode.RenderInline : TrackedChangeMode.Accept,
                RevisionAuthor = "Bob",
            });
            long version = session.Version;

            var result = InsertBlock(session, operation, BlockAnchor(session, kind), position);

            Assert.False(result.Success);
            Assert.Equal(EditErrorCode.AnchorWrongKind, result.Error!.Code);
            Assert.Equal(version, session.Version);
            Assert.False(session.Undo());
            Assert.True(XNode.DeepEquals(DocumentXml(baseline), DocumentXml(session.Save())));
        }
    }

    /// <summary>Control: a paragraph (in the body or in a cell), a table and a block content
    /// control all take a block beside them, and the result validates.</summary>
    [Theory]
    [InlineData("paragraph", "body_paragraph")]
    [InlineData("paragraph", "cell_paragraph")]
    [InlineData("paragraph", "tbl")]
    [InlineData("paragraph", "block_sdt")]
    [InlineData("paragraph", "custom_paragraph")]
    [InlineData("rule", "custom_paragraph")]
    [InlineData("table", "custom_paragraph")]
    [InlineData("rule", "cell_paragraph")]
    [InlineData("rule", "tbl")]
    [InlineData("table", "body_paragraph")]
    [InlineData("table", "block_sdt")]
    public void DT261_BlockInsertion_BesideABlock_StillSucceeds(string operation, string kind)
    {
        foreach (var position in new[] { Position.Before, Position.After })
        {
            using var session = new DocxSession(BlockAnchorDocument());

            var result = InsertBlock(session, operation, BlockAnchor(session, kind), position);

            Assert.True(result.Success, result.Error?.Message);
            var saved = session.Save();
            PackageValidation.AssertValid(saved, DocumentFormat.OpenXml.FileFormatVersions.Office2019);
            PackageValidation.AssertValid(saved, DocumentFormat.OpenXml.FileFormatVersions.Office2021);
        }
    }
}
