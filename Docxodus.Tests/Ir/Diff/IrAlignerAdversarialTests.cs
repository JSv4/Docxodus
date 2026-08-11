#nullable enable

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.1 Task 3 adversarial + scale coverage for <see cref="IrBlockAligner"/>: boilerplate-heavy and
/// fully-rewritten stress fixtures, a contiguous block-move LIS check, and an anti-O(n²) scale guard.
/// All documents are built programmatically via <see cref="IrTestDocuments"/> + <see cref="IrReader"/>.
/// </summary>
[Collection(IrAlignerPerformanceCollection.Name)]
public class IrAlignerAdversarialTests
{
    private static readonly IrReaderOptions NoSources =
        new() { RetainSources = false, RevisionView = RevisionView.Accept };
    private static readonly IrDiffSettings Diff = new();

    private readonly ITestOutputHelper _out;

    public IrAlignerAdversarialTests(ITestOutputHelper output) => _out = output;

    private static IrDocument Doc(IEnumerable<string> paras) =>
        IrReader.Read(IrTestDocuments.Create(paras.ToArray()), NoSources);

    private static IrBlockAlignment Align(IrDocument l, IrDocument r) =>
        IrBlockAligner.Align(l, r, Diff);

    private static int Count(IrBlockAlignment a, IrAlignmentKind k) => IrAlignmentAsserts.Count(a, k);

    // Each paragraph unique by its clause number → all hashes distinct (no boilerplate collisions).
    private static string[] DistinctClauses(int n) =>
        Enumerable.Range(0, n)
            .Select(i => $"Clause {i}: standard wording for this section of the agreement.")
            .ToArray();

    // ------------------------------------------------------------------ near-identical, one edit

    [Fact]
    public void NearIdentical_500_one_word_changed_yields_499_unchanged_1_modified_0_moved()
    {
        var left = DistinctClauses(500);
        var right = (string[])left.Clone();
        // Change one word in exactly one paragraph (kept unique, no hash collision with any sibling).
        right[250] = "Clause 250: REVISED wording for this section of the agreement.";

        var l = Doc(left);
        var r = Doc(right);
        var a = Align(l, r);

        Assert.Equal(499, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(1, Count(a, IrAlignmentKind.Modified));
        Assert.Equal(0, Count(a, IrAlignmentKind.Moved));
        Assert.Equal(0, Count(a, IrAlignmentKind.Inserted));
        Assert.Equal(0, Count(a, IrAlignmentKind.Deleted));
        IrAlignmentAsserts.AssertInvariants(l, r, a);
    }

    // ------------------------------------------------------------------ identical boilerplate, one deleted

    [Fact]
    public void Identical_500_delete_one_yields_499_unchanged_1_deleted_0_moved_0_modified()
    {
        var left = Enumerable.Repeat("Standard boilerplate clause.", 500).ToArray();
        var right = Enumerable.Repeat("Standard boilerplate clause.", 499).ToArray();

        var l = Doc(left);
        var r = Doc(right);
        var a = Align(l, r);

        Assert.Equal(499, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(1, Count(a, IrAlignmentKind.Deleted));
        Assert.Equal(0, Count(a, IrAlignmentKind.Moved));
        Assert.Equal(0, Count(a, IrAlignmentKind.Modified));
        IrAlignmentAsserts.AssertInvariants(l, r, a);
    }

    // ------------------------------------------------------------------ fully rewritten

    [Fact]
    public void Fully_rewritten_200_vs_200_no_throw_invariants_hold_runtime_sane()
    {
        var left = Enumerable.Range(0, 200)
            .Select(i => $"Original paragraph {i} with its own distinct content.")
            .ToArray();
        var right = Enumerable.Range(0, 200)
            .Select(i => $"Completely different replacement line {i} sharing nothing.")
            .ToArray();

        var l = Doc(left);
        var r = Doc(right);

        var sw = Stopwatch.StartNew();
        var a = Align(l, r);
        sw.Stop();

        IrAlignmentAsserts.AssertInvariants(l, r, a);
        // M2.2 Task 3 re-baseline. No exact-hash anchors exist; everything resolves in one big head↔tail
        // gap. The two sides share NOTHING (every left line is "Original paragraph i …", every right line
        // is "Completely different replacement line i …"), so every candidate pair scores 0 < the 0.5
        // BlockSimilarityThreshold and NONE pair as Modified. The 1×1 unambiguous-residue fallback does not
        // apply (200 free on each side, not 1×1). So the correct classification is 200 Deleted + 200
        // Inserted — claiming 200 in-place Modified edits (the M2.1 blind-positional behavior this pass
        // replaces) would falsely assert each replacement line is a revision of the i-th original line.
        // Cross-gap move detection finds nothing either (no pair clears the 0.8 MoveSimilarityThreshold).
        Assert.Equal(0, Count(a, IrAlignmentKind.Modified));
        Assert.Equal(200, Count(a, IrAlignmentKind.Deleted));
        Assert.Equal(200, Count(a, IrAlignmentKind.Inserted));
        Assert.Equal(0, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(0, Count(a, IrAlignmentKind.Moved));
        Assert.Equal(0, Count(a, IrAlignmentKind.MovedModified));
        _out.WriteLine($"Fully-rewritten 200x200: {IrAlignmentAsserts.Histogram(a)} in {sw.ElapsedMilliseconds} ms");
        Assert.True(sw.ElapsedMilliseconds < 5000, $"Rewrite align took {sw.ElapsedMilliseconds} ms — too slow.");
    }

    // ------------------------------------------------------------------ contiguous block move

    [Fact]
    public void Move_10_unique_paragraph_block_front_to_back_of_300_yields_exactly_10_moved()
    {
        // 300 unique paragraphs. Take the FIRST 10 and relocate them (as a contiguous block, order
        // preserved) to the very end. The other 290 stay in their original relative order.
        //
        // LIS reasoning: anchors pair all 300 by exact content. The (leftIndex, rightIndex) pairs are:
        //   moved block:   left 0..9   -> right 290..299  (still increasing AMONG themselves)
        //   stationary:    left 10..299 -> right 0..289   (also increasing among themselves)
        // The longest increasing subsequence by right index picks the larger monotone chain. The
        // stationary 290 occupy right positions 0..289 with left positions 10..299 (increasing), so they
        // ARE a length-290 increasing subsequence. The moved 10 land at right 290..299 with left 0..9 —
        // increasing among themselves, but to JOIN the spine after the stationary chain they'd need a
        // left index > 299, which they don't have (their left indices 0..9 are the smallest). So the LIS
        // keeps the 290 stationary and drops the 10-block off the spine → exactly 10 Moved.
        const int total = 300;
        const int blockSize = 10;
        var all = DistinctClauses(total);

        var left = all.ToArray();
        var movedBlock = all.Take(blockSize).ToArray();
        var rest = all.Skip(blockSize).ToArray();
        var right = rest.Concat(movedBlock).ToArray();

        var l = Doc(left);
        var r = Doc(right);
        var a = Align(l, r);

        IrAlignmentAsserts.AssertInvariants(l, r, a);
        _out.WriteLine($"Block-move 10-of-300: {IrAlignmentAsserts.Histogram(a)}");

        Assert.Equal(blockSize, Count(a, IrAlignmentKind.Moved));
        Assert.Equal(total - blockSize, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(0, Count(a, IrAlignmentKind.Modified));
        Assert.Equal(0, Count(a, IrAlignmentKind.Inserted));
        Assert.Equal(0, Count(a, IrAlignmentKind.Deleted));

        // The moved entries are exactly the relocated block's content.
        var movedTexts = a.Entries
            .Where(e => e.Kind == IrAlignmentKind.Moved)
            .Select(e => Text(e.Right!))
            .ToHashSet();
        Assert.Equal(movedBlock.ToHashSet(), movedTexts);
    }

    // ------------------------------------------------------------------ scale guard (anti-O(n²))

    [Trait("Category", "Perf")]
    [Fact]
    public void Scale_guard_1000_vs_4000_cpu_ratio_within_15x()
    {
        // Both inputs are the near-identical fixture (distinct clauses) self-paired with ONE edit, so
        // every block anchors uniquely and the only gap is a single 1-block Modified gap — i.e. NO large
        // all-distinct gap that would trip the InOrderRefine G²/2 worst case. This isolates the spine /
        // anchoring cost, which should scale ~linearly: 4× the blocks ⇒ well under 12× the CPU time
        // (a true O(n²) regression reads ~16×). Process CPU time keeps shared-runner scheduling noise
        // from looking like algorithmic work.
        // Best of up to three independent rounds. Scheduling noise on a shared runner can only ADD CPU
        // time, so it can only inflate the ratio — which makes the MINIMUM across rounds the closest
        // estimate of the true algorithmic one, and makes a single noisy round unable to fail the test.
        // The guard keeps its teeth: a real O(n²) regression reads ~16x in EVERY round, so all three
        // have to exceed the limit before this fails. Rounds stop as soon as one comes in under.
        //
        // Sizes are 1000/4000 (was 500/2000): at 500 paragraphs the baseline sample sits at ~2-3 ms,
        // where the independent per-size minimum is biased — the small input reaches its ideal sample
        // far more often than the cache-pressured large one, inflating the ratio. CI tripped the 12x
        // guard three times in a row on unchanged aligner code (12.20x, 12.29x, 15.37x best-of-rounds)
        // with baselines of 2.28-2.97 ms. Quadrupling both sizes keeps the 4x scale and the limit's
        // meaning while making the denominator large enough that scheduler noise stops deciding the
        // verdict.
        //
        // Limit calibration: with the solid 7-8 ms baseline the measured ratio on GitHub's shared
        // runners is 12.5-13x in EVERY round (e.g. 1000=7.64 ms, 4000=97.88 ms ⇒ 12.82x best-of-3) —
        // that is the aligner's true linear-plus-cache profile at these sizes, not noise: 4000
        // paragraphs outgrow cache and raise the per-item constant. A genuine O(n²) regression reads
        // ≥16x from the algorithm alone, before the same cache multiplier pushes it higher, so 15x
        // still separates the two regimes cleanly while sitting above the measured healthy band.
        const double limit = 15.0;
        const int rounds = 3;

        double bestRatio = double.MaxValue, bestSmall = 0, bestLarge = 0;
        for (int round = 0; round < rounds; round++)
        {
            double small = BestSampleCpuMs(1000);
            double large = BestSampleCpuMs(4000);
            double ratio = large / Math.Max(small, 0.0001);
            _out.WriteLine($"Scale guard CPU round {round + 1}: 1000-para = {small:F2} ms, " +
                $"4000-para = {large:F2} ms, ratio = {ratio:F2}x (n=4x)");

            if (ratio < bestRatio)
                (bestRatio, bestSmall, bestLarge) = (ratio, small, large);
            if (bestRatio <= limit)
                break;
        }

        Assert.True(bestRatio <= limit,
            $"Align CPU-time ratio {bestRatio:F2}x for 4x input exceeds the {limit:F0}x anti-O(n²) guard " +
            $"in every one of {rounds} rounds (best round: 1000={bestSmall:F2}ms, 4000={bestLarge:F2}ms).");
    }

    /// <summary>
    /// Collect first, warm up once, then take the best of five process-CPU samples (ms per align)
    /// for an n-paragraph self-pair with one edit. Each sample times a batch of 10 aligns so the
    /// measurement stays well above platform timer granularity.
    /// </summary>
    private static double BestSampleCpuMs(int n)
    {
        const int alignsPerSample = 10;
        var baseParas = DistinctClauses(n);
        var edited = (string[])baseParas.Clone();
        edited[n / 2] = $"Clause {n / 2}: REVISED wording for this section of the agreement.";

        var l = Doc(baseParas);
        var r = Doc(edited);

        GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: true);
        GC.WaitForPendingFinalizers();
        _ = Align(l, r); // warm-up (JIT, dictionary growth)

        using var process = Process.GetCurrentProcess();
        double best = double.MaxValue;
        for (int i = 0; i < 5; i++)
        {
            TimeSpan start = process.TotalProcessorTime;
            for (int j = 0; j < alignsPerSample; j++)
                _ = Align(l, r);
            TimeSpan elapsed = process.TotalProcessorTime - start;
            best = Math.Min(best, elapsed.TotalMilliseconds / alignsPerSample);
        }
        return best;
    }

    private static string Text(IrBlock b) =>
        b is IrParagraph p
            ? string.Concat(p.Inlines.OfType<IrTextRun>().Select(t => t.Text))
            : string.Empty;
}

/// <summary>
/// Keeps the scale guard isolated from the default parallel suite so unrelated test work is not
/// counted in this process's CPU-time measurement.
/// </summary>
[CollectionDefinition(Name, DisableParallelization = true)]
public sealed class IrAlignerPerformanceCollection
{
    public const string Name = "IR aligner performance";
}
