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

        _out.WriteLine($"WC corpus: {pairs.Count} pairs (each built + verified + round-tripped, fwd + rev)");
        _out.WriteLine("");

        foreach (var (baseName, variantName) in pairs)
        {
            var baseDoc = WcCorpus.ReadWc(baseName);
            var variantDoc = WcCorpus.ReadWc(variantName);

            VerifyOne(baseDoc, variantDoc, totals, accumulate: true);   // forward (accumulates totals)
            VerifyOne(variantDoc, baseDoc, totals, accumulate: false);  // reversed
        }

        _out.WriteLine("Corpus op-kind totals (forward direction):");
        foreach (var kv in totals.OrderBy(k => (int)k.Key))
            _out.WriteLine($"  {kv.Key} = {kv.Value}");
    }

    private static void VerifyOne(
        Docxodus.Ir.IrDocument left, Docxodus.Ir.IrDocument right,
        Dictionary<IrEditOpKind, int> totals, bool accumulate)
    {
        var script = IrEditScriptBuilder.Build(left, right, Diff);

        // Exit invariant: apply(script, left) reconstructs right at text level.
        IrEditScriptVerifier.Verify(left, right, script, Diff);

        // JSON round-trip: Read(Write(s)) is record-equal to s, and Write is deterministic.
        var json = IrEditScriptJson.Write(script);
        var back = IrEditScriptJson.Read(json);
        Assert.Equal(script, back);
        Assert.Equal(json, IrEditScriptJson.Write(back));

        if (accumulate)
            foreach (var op in script.Operations)
                totals[op.Kind]++;
    }
}
