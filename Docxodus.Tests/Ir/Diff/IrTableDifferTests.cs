#nullable enable

using System.Linq;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.2 Task 4 tests for row/cell table granularity: a Modified table pair produces a nested
/// <see cref="IrTableDiff"/> whose row ops align by content and whose cell-text edit surfaces as a TOKEN
/// DIFF inside the affected cell (THE headline test), not a whole-table blob. Also covers row
/// insert/delete, unchanged rows on the spine, JSON round-trip of the nested diff, and apply-verify.
/// </summary>
public class IrTableDifferTests
{
    private static readonly IrReaderOptions NoSources = new() { RetainSources = false };
    private static readonly IrDiffSettings Default = new();

    private static IrDocument FromXml(string bodyInnerXml) =>
        IrReader.Read(IrTestDocuments.FromBodyXml(bodyInnerXml), NoSources);

    private static string Cell(string text) =>
        $"<w:tc><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";

    private static string Cell(string text, int width) =>
        $"<w:tc><w:tcPr><w:tcW w:w=\"{width}\" w:type=\"dxa\"/></w:tcPr>" +
        $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:tc>";

    private static string Row(params string[] cells) =>
        $"<w:tr>{string.Concat(cells)}</w:tr>";

    private static string Table(params string[] rows) =>
        $"<w:tbl><w:tblPr/><w:tblGrid/>{string.Concat(rows)}</w:tbl>";

    private static string GridTable(int width, int columns, params string[] rows) =>
        "<w:tbl><w:tblPr/><w:tblGrid>" +
        string.Concat(Enumerable.Repeat($"<w:gridCol w:w=\"{width}\"/>", columns)) +
        $"</w:tblGrid>{string.Concat(rows)}</w:tbl>";

    private static IrEditOp TableOp(IrDocument l, IrDocument r)
    {
        var script = IrEditScriptBuilder.Build(l, r, Default);
        var op = script.Operations.Single(o => o.Kind == IrEditOpKind.ModifyBlock);
        Assert.NotNull(op.TableDiff);
        // Always apply-verify + JSON-round-trip the produced script.
        IrEditScriptVerifier.Verify(l, r, script, Default);
        var back = IrEditScriptJson.Read(IrEditScriptJson.Write(script));
        Assert.Equal(script, back);
        return op;
    }

    [Fact]
    public void Cell_text_edit_surfaces_as_token_diff_in_that_cell()
    {
        // Two-row, two-column table; one cell's text edited.
        var left = FromXml(Table(
            Row(Cell("alpha one"), Cell("beta two")),
            Row(Cell("gamma three"), Cell("delta four"))));
        var right = FromXml(Table(
            Row(Cell("alpha one"), Cell("beta two")),
            Row(Cell("gamma three"), Cell("delta EDITED"))));

        var op = TableOp(left, right);
        var table = op.TableDiff!;

        // Row 0 unchanged (on the spine), row 1 modified.
        Assert.Equal(IrRowOpKind.EqualRow, table.RowOps[0].Kind);
        var modRow = table.RowOps[1];
        Assert.Equal(IrRowOpKind.ModifyRow, modRow.Kind);
        Assert.NotNull(modRow.CellOps);

        // Cell 0 of the modified row unchanged (no block ops); cell 1 carries a block-level ModifyBlock
        // whose TOKEN diff describes the in-cell text edit — NOT a whole-table blob.
        var cell0 = modRow.CellOps![0];
        Assert.Null(cell0.BlockOps); // content-equal cell ⇒ no recursion

        var cell1 = modRow.CellOps![1];
        Assert.NotNull(cell1.BlockOps);
        var blockOp = cell1.BlockOps!.Single();
        Assert.Equal(IrEditOpKind.ModifyBlock, blockOp.Kind);
        Assert.NotNull(blockOp.TokenDiff); // the cell paragraph's token diff
        // The token diff has at least one Delete or Insert (the edited word) and some Equal.
        Assert.Contains(blockOp.TokenDiff!.Ops, o => o.Kind is IrTokenOpKind.Insert or IrTokenOpKind.Delete);
        Assert.Contains(blockOp.TokenDiff!.Ops, o => o.Kind == IrTokenOpKind.Equal);
    }

    [Fact]
    public void Row_inserted_and_deleted()
    {
        var left = FromXml(Table(
            Row(Cell("keep me")),
            Row(Cell("delete me")),
            Row(Cell("also keep"))));
        var right = FromXml(Table(
            Row(Cell("keep me")),
            Row(Cell("brand new")),
            Row(Cell("also keep"))));

        var op = TableOp(left, right);
        var rowOps = op.TableDiff!.RowOps;

        // "keep me" + "also keep" are unique-hash row anchors on the spine; the middle row is a
        // delete (old) + insert (new) OR a modify. With both surviving rows anchored, the middle gap
        // has one free left + one free right → positional ModifyRow.
        Assert.Equal(2, rowOps.Count(o => o.Kind == IrRowOpKind.EqualRow));
        Assert.Equal(1, rowOps.Count(o => o.Kind == IrRowOpKind.ModifyRow));
    }

    [Fact]
    public void Row_only_added()
    {
        var left = FromXml(Table(Row(Cell("one")), Row(Cell("two"))));
        var right = FromXml(Table(Row(Cell("one")), Row(Cell("two")), Row(Cell("three"))));

        var op = TableOp(left, right);
        var rowOps = op.TableDiff!.RowOps;

        Assert.Equal(2, rowOps.Count(o => o.Kind == IrRowOpKind.EqualRow));
        Assert.Equal(1, rowOps.Count(o => o.Kind == IrRowOpKind.InsertRow));
        Assert.Equal(0, rowOps.Count(o => o.Kind == IrRowOpKind.DeleteRow));
    }

    [Fact]
    public void Deterministic_table_diff()
    {
        var left = FromXml(Table(
            Row(Cell("a"), Cell("b")),
            Row(Cell("c"), Cell("d"))));
        var right = FromXml(Table(
            Row(Cell("a"), Cell("b")),
            Row(Cell("c"), Cell("D-edited"))));

        var first = IrEditScriptBuilder.Build(left, right, Default);
        var second = IrEditScriptBuilder.Build(left, right, Default);
        Assert.Equal(first, second);
    }

    [Fact]
    public void Ordinary_grid_middle_cell_insert_uses_body_spine_when_all_cell_widths_change()
    {
        // The retained A/B/C text is identical, but every tcPr/tcW changes as the table grows 3 → 4
        // columns.  Full cell ContentHash therefore cannot anchor B/C; the ordinary-grid path must use
        // the shell-free cell body key and place NEW at its actual right-side position.
        var left = FromXml(GridTable(2000, 3,
            Row(Cell("A", 2000), Cell("B", 2000), Cell("C", 2000)),
            Row(Cell("D", 2000), Cell("E", 2000), Cell("F", 2000))));
        var right = FromXml(GridTable(1500, 4,
            Row(Cell("A", 1500), Cell("NEW-1", 1500), Cell("B", 1500), Cell("C", 1500)),
            Row(Cell("D", 1500), Cell("NEW-2", 1500), Cell("E", 1500), Cell("F", 1500))));

        var op = TableOp(left, right);
        var leftTable = Assert.IsType<IrTable>(left.Body.Blocks.Single());
        var rightTable = Assert.IsType<IrTable>(right.Body.Blocks.Single());
        var rows = op.TableDiff!.RowOps.Where(r => r.Kind == IrRowOpKind.ModifyRow).ToList();
        Assert.Equal(2, rows.Count);

        for (int row = 0; row < rows.Count; row++)
        {
            var cells = rows[row].CellOps!;
            Assert.Equal(4, cells.Count);
            Assert.Equal(leftTable.Rows[row].Cells[0].Anchor.ToString(), cells[0].LeftCellAnchor);
            Assert.Equal(rightTable.Rows[row].Cells[0].Anchor.ToString(), cells[0].RightCellAnchor);
            Assert.Null(cells[1].LeftCellAnchor);
            Assert.Equal(rightTable.Rows[row].Cells[1].Anchor.ToString(), cells[1].RightCellAnchor);
            Assert.Equal(leftTable.Rows[row].Cells[1].Anchor.ToString(), cells[2].LeftCellAnchor);
            Assert.Equal(rightTable.Rows[row].Cells[2].Anchor.ToString(), cells[2].RightCellAnchor);
            Assert.Equal(leftTable.Rows[row].Cells[2].Anchor.ToString(), cells[3].LeftCellAnchor);
            Assert.Equal(rightTable.Rows[row].Cells[3].Anchor.ToString(), cells[3].RightCellAnchor);
            Assert.DoesNotContain(cells, c => c.RightCellAnchor == null);
        }

        Assert.NotEqual(leftTable.Rows[0].Cells[1].ContentHash, rightTable.Rows[0].Cells[2].ContentHash);
    }

    [Fact]
    public void Ordinary_grid_mixed_middle_insert_and_edit_aligns_insert_before_edited_retained_cell()
    {
        // B is edited while X is inserted before it. A/C are exact body anchors, leaving B versus X/B2 in
        // one free gap. The costed monotone fill must prefer B→B2 and identify X as the real inserted cell;
        // positional pairing would instead shift B onto X and call C an inserted tail cell.
        var left = FromXml(GridTable(2000, 3,
            Row(Cell("A", 2000), Cell("B", 2000), Cell("C", 2000))));
        var right = FromXml(GridTable(1500, 4,
            Row(Cell("A", 1500), Cell("X", 1500), Cell("B2", 1500), Cell("C", 1500))));

        var op = TableOp(left, right);
        var leftTable = Assert.IsType<IrTable>(left.Body.Blocks.Single());
        var rightTable = Assert.IsType<IrTable>(right.Body.Blocks.Single());
        var cells = Assert.Single(op.TableDiff!.RowOps, r => r.Kind == IrRowOpKind.ModifyRow).CellOps!;

        // A/A, +X, B/B2, C/C. The edited retained cell remains paired so its paragraph diff is available;
        // only the actual X cell is emitted as a right-only native cell insertion.
        Assert.Equal(4, cells.Count);
        Assert.Equal(leftTable.Rows[0].Cells[0].Anchor.ToString(), cells[0].LeftCellAnchor);
        Assert.Equal(rightTable.Rows[0].Cells[0].Anchor.ToString(), cells[0].RightCellAnchor);
        Assert.Null(cells[1].LeftCellAnchor);
        Assert.Equal(rightTable.Rows[0].Cells[1].Anchor.ToString(), cells[1].RightCellAnchor);
        Assert.Equal(leftTable.Rows[0].Cells[1].Anchor.ToString(), cells[2].LeftCellAnchor);
        Assert.Equal(rightTable.Rows[0].Cells[2].Anchor.ToString(), cells[2].RightCellAnchor);
        Assert.NotNull(cells[2].BlockOps);
        Assert.Equal(leftTable.Rows[0].Cells[2].Anchor.ToString(), cells[3].LeftCellAnchor);
        Assert.Equal(rightTable.Rows[0].Cells[3].Anchor.ToString(), cells[3].RightCellAnchor);
        Assert.DoesNotContain(cells, cell => cell.RightCellAnchor is null);
    }

    [Fact]
    public void Ordinary_rows_mixed_middle_insert_and_edit_align_insert_before_edited_retained_row()
    {
        // The row analogue of the cell case above. The row renderer natively supports both insertions and
        // deletions, so the same bounded alignment can always surface an inserted row plus a paired edit.
        var left = FromXml(Table(Row(Cell("A")), Row(Cell("B")), Row(Cell("C"))));
        var right = FromXml(Table(Row(Cell("A")), Row(Cell("X")), Row(Cell("B2")), Row(Cell("C"))));

        var op = TableOp(left, right);
        var leftTable = Assert.IsType<IrTable>(left.Body.Blocks.Single());
        var rightTable = Assert.IsType<IrTable>(right.Body.Blocks.Single());
        var rows = op.TableDiff!.RowOps;

        Assert.Equal(4, rows.Count);
        Assert.Equal(IrRowOpKind.EqualRow, rows[0].Kind);
        Assert.Equal(leftTable.Rows[0].Anchor.ToString(), rows[0].LeftRowAnchor);
        Assert.Equal(rightTable.Rows[0].Anchor.ToString(), rows[0].RightRowAnchor);
        Assert.Equal(IrRowOpKind.InsertRow, rows[1].Kind);
        Assert.Null(rows[1].LeftRowAnchor);
        Assert.Equal(rightTable.Rows[1].Anchor.ToString(), rows[1].RightRowAnchor);
        Assert.Equal(IrRowOpKind.ModifyRow, rows[2].Kind);
        Assert.Equal(leftTable.Rows[1].Anchor.ToString(), rows[2].LeftRowAnchor);
        Assert.Equal(rightTable.Rows[2].Anchor.ToString(), rows[2].RightRowAnchor);
        Assert.NotNull(rows[2].CellOps);
        Assert.Equal(IrRowOpKind.EqualRow, rows[3].Kind);
        Assert.Equal(leftTable.Rows[2].Anchor.ToString(), rows[3].LeftRowAnchor);
        Assert.Equal(rightTable.Rows[3].Anchor.ToString(), rows[3].RightRowAnchor);
    }

    [Fact]
    public void Merged_or_offset_rows_keep_conservative_positional_cell_pairs()
    {
        // gridSpan and gridBefore need a topology-aware phase.  Phase 1 must not reinterpret either as an
        // ordinary-column insertion just because their cell text happens to form a unique sequence.
        var left = FromXml(
            "<w:tbl><w:tblPr/><w:tblGrid/>" +
            "<w:tr><w:trPr><w:gridBefore w:val=\"1\"/></w:trPr>" +
            "<w:tc><w:tcPr><w:gridSpan w:val=\"2\"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
            Cell("B") + "</w:tr></w:tbl>");
        var right = FromXml(
            "<w:tbl><w:tblPr/><w:tblGrid/>" +
            "<w:tr><w:trPr><w:gridBefore w:val=\"1\"/></w:trPr>" +
            "<w:tc><w:tcPr><w:gridSpan w:val=\"2\"/></w:tcPr><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>" +
            Cell("NEW") + Cell("B") + "</w:tr></w:tbl>");

        var op = TableOp(left, right);
        var leftTable = Assert.IsType<IrTable>(left.Body.Blocks.Single());
        Assert.Equal(1, leftTable.Rows[0].GridBefore);
        var cells = Assert.Single(op.TableDiff!.RowOps, r => r.Kind == IrRowOpKind.ModifyRow).CellOps!;

        // The existing positional result remains: B is paired with NEW and the right B is a surplus tail.
        Assert.Equal(3, cells.Count);
        Assert.NotNull(cells[1].LeftCellAnchor);
        Assert.NotNull(cells[1].RightCellAnchor);
        Assert.Null(cells[2].LeftCellAnchor);
        Assert.NotNull(cells[2].RightCellAnchor);
    }
}
