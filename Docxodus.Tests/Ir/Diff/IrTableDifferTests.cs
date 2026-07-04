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

    private static string Row(params string[] cells) =>
        $"<w:tr>{string.Concat(cells)}</w:tr>";

    private static string Table(params string[] rows) =>
        $"<w:tbl><w:tblPr/><w:tblGrid/>{string.Concat(rows)}</w:tbl>";

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
}
