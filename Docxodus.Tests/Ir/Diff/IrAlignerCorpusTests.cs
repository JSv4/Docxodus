#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.1 Task 3 corpus smoke: run <see cref="IrBlockAligner"/> over every base↔variant DOCX pair we can
/// infer from <c>TestFiles/WC/</c> (the pair list + reader options live in <see cref="WcCorpus"/>),
/// asserting totality (no throw) + the shared invariants both forward (before→after) and reversed
/// (after→before), and logging a per-pair entry-kind histogram plus corpus totals.
/// </summary>
[Trait("Category", "Corpus")]
public class IrAlignerCorpusTests
{
    private static readonly IrDiffSettings Diff = new();

    private readonly ITestOutputHelper _out;

    public IrAlignerCorpusTests(ITestOutputHelper output) => _out = output;

    [Fact]
    public void WC_corpus_pairs_align_without_throwing_invariants_hold_both_directions()
    {
        var pairs = WcCorpus.BuildPairs();
        Assert.True(pairs.Count >= 30,
            $"Expected a substantial WC pair list; only inferred {pairs.Count}. Naming convention drift?");

        var totals = new Dictionary<IrAlignmentKind, int>();
        foreach (var k in Enum.GetValues<IrAlignmentKind>())
            totals[k] = 0;

        _out.WriteLine($"WC corpus: {pairs.Count} base↔variant pairs (each run forward + reversed)");
        _out.WriteLine("");

        foreach (var (baseName, variantName) in pairs)
        {
            var baseDoc = WcCorpus.ReadWc(baseName);
            var variantDoc = WcCorpus.ReadWc(variantName);

            var fwd = IrBlockAligner.Align(baseDoc, variantDoc, Diff);
            IrAlignmentAsserts.AssertInvariants(baseDoc, variantDoc, fwd);

            var rev = IrBlockAligner.Align(variantDoc, baseDoc, Diff);
            IrAlignmentAsserts.AssertInvariants(variantDoc, baseDoc, rev);

            foreach (var e in fwd.Entries)
                totals[e.Kind]++;

            _out.WriteLine($"  {baseName} -> {variantName}");
            _out.WriteLine($"      fwd: {IrAlignmentAsserts.Histogram(fwd)}");
            _out.WriteLine($"      rev: {IrAlignmentAsserts.Histogram(rev)}");
        }

        _out.WriteLine("");
        _out.WriteLine("Corpus totals (forward direction):");
        foreach (var kv in totals.OrderBy(k => (int)k.Key))
            _out.WriteLine($"  {kv.Key} = {kv.Value}");
    }
}
