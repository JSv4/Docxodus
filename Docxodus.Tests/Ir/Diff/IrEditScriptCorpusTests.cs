#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using Docxodus.Ir.Diff;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.2 Task 2 corpus exit-invariant: over every WC base↔variant pair (and the reversed direction),
/// build the <see cref="IrEditScript"/>, run the apply-verifier (apply(script, left) reconstructs right
/// at text level), and JSON-round-trip it. Logs the corpus-wide op-kind histogram totals.
/// </summary>
[Trait("Category", "Corpus")]
public class IrEditScriptCorpusTests
{
    private static readonly IrDiffSettings Diff = new();

    private readonly ITestOutputHelper _out;

    public IrEditScriptCorpusTests(ITestOutputHelper output) => _out = output;

    [Fact]
    public void WC_corpus_edit_scripts_apply_verify_and_json_round_trip_both_directions()
    {
        var pairs = WcCorpus.BuildPairs();
        Assert.True(pairs.Count >= 30,
            $"Expected a substantial WC pair list; only inferred {pairs.Count}. Naming convention drift?");

        var totals = new Dictionary<IrEditOpKind, int>();
        foreach (var k in Enum.GetValues<IrEditOpKind>())
            totals[k] = 0;
        var tableStats = new TableStats();

        _out.WriteLine($"WC corpus: {pairs.Count} pairs (each built + verified + round-tripped, fwd + rev)");
        _out.WriteLine("");

        foreach (var (baseName, variantName) in pairs)
        {
            var baseDoc = WcCorpus.ReadWc(baseName);
            var variantDoc = WcCorpus.ReadWc(variantName);

            VerifyOne(baseDoc, variantDoc, totals, tableStats, accumulate: true);   // forward (accumulates)
            VerifyOne(variantDoc, baseDoc, totals, tableStats, accumulate: false);  // reversed
        }

        _out.WriteLine("Corpus op-kind totals (forward direction):");
        foreach (var kv in totals.OrderBy(k => (int)k.Key))
            _out.WriteLine($"  {kv.Key} = {kv.Value}");
        _out.WriteLine("");
        _out.WriteLine("Table-diff stats (forward, M2.2 Task 4):");
        _out.WriteLine($"  ModifyBlock-with-table-diff = {tableStats.TablesDiffed}");
        _out.WriteLine($"  row ops total = {tableStats.RowOps} " +
            $"(Equal={tableStats.EqualRows} Modify={tableStats.ModifyRows} " +
            $"Insert={tableStats.InsertRows} Delete={tableStats.DeleteRows} Moved={tableStats.MovedRows})");
        _out.WriteLine($"  cell ops total = {tableStats.CellOps}; cells with a block token diff = {tableStats.CellsWithTokenDiff}");
    }

    private sealed class TableStats
    {
        public int TablesDiffed, RowOps, EqualRows, ModifyRows, InsertRows, DeleteRows, MovedRows;
        public int CellOps, CellsWithTokenDiff;
    }

    private static void VerifyOne(
        Docxodus.Ir.IrDocument left, Docxodus.Ir.IrDocument right,
        Dictionary<IrEditOpKind, int> totals, TableStats tableStats, bool accumulate)
    {
        var script = IrEditScriptBuilder.Build(left, right, Diff);

        // Exit invariant: apply(script, left) reconstructs right at text level (validates nested table
        // row/cell diffs + their anchors too).
        IrEditScriptVerifier.Verify(left, right, script, Diff);

        // JSON round-trip: Read(Write(s)) is record-equal to s, and Write is deterministic.
        var json = IrEditScriptJson.Write(script);
        var back = IrEditScriptJson.Read(json);
        Assert.Equal(script, back);
        Assert.Equal(json, IrEditScriptJson.Write(back));

        if (accumulate)
            foreach (var op in script.Operations)
            {
                totals[op.Kind]++;
                if (op.TableDiff is { } td)
                    AccumulateTable(td, tableStats);
            }
    }

    private static void AccumulateTable(IrTableDiff td, TableStats s)
    {
        s.TablesDiffed++;
        foreach (var rowOp in td.RowOps)
        {
            s.RowOps++;
            switch (rowOp.Kind)
            {
                case IrRowOpKind.EqualRow: s.EqualRows++; break;
                case IrRowOpKind.ModifyRow: s.ModifyRows++; break;
                case IrRowOpKind.InsertRow: s.InsertRows++; break;
                case IrRowOpKind.DeleteRow: s.DeleteRows++; break;
                case IrRowOpKind.MovedRow: s.MovedRows++; break;
            }
            if (rowOp.CellOps is { } cellOps)
                foreach (var cellOp in cellOps)
                {
                    s.CellOps++;
                    if (cellOp.BlockOps is { } blockOps &&
                        blockOps.Any(b => b.TokenDiff is not null))
                        s.CellsWithTokenDiff++;
                }
        }
    }
}
