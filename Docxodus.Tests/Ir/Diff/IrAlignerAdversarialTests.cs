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
    public void Scale_guard_1000_vs_4000_allocation_ratio_within_12x()
    {
        // Both inputs are the near-identical fixture (distinct clauses) self-paired with ONE edit, so
        // every block anchors uniquely and the only gap is a single 1-block Modified gap — i.e. NO large
        // all-distinct gap that would trip the InOrderRefine G²/2 worst case. This isolates the spine /
        // anchoring storage, which should scale near-linearly: 4× the blocks must stay well below the
        // 16× signature of quadratic materialization. Per-thread managed allocations are deterministic
        // for this synchronous path and cannot be distorted by runner scheduling or CPU-cache pressure.
        const double limit = 12.0;

        long small = AllocatedBytesForAlign(1000);
        long large = AllocatedBytesForAlign(4000);
        double ratio = (double)large / Math.Max(small, 1L);

        _out.WriteLine($"Scale guard allocations: 1000-para = {small:N0} bytes, " +
            $"4000-para = {large:N0} bytes, ratio = {ratio:F2}x (n=4x)");
        Assert.True(ratio <= limit,
            $"Align allocation ratio {ratio:F2}x for 4x input exceeds the {limit:F0}x " +
            $"anti-quadratic-materialization guard (1000={small:N0}, 4000={large:N0} bytes).");
    }

    /// <summary>
    /// Warm up once, then count allocations made on the calling thread by one alignment of an
    /// n-paragraph self-pair with one edit. Document construction is deliberately outside the sample.
    /// </summary>
    private static long AllocatedBytesForAlign(int n)
    {
        var baseParas = DistinctClauses(n);
        var edited = (string[])baseParas.Clone();
        edited[n / 2] = $"Clause {n / 2}: REVISED wording for this section of the agreement.";

        var l = Doc(baseParas);
        var r = Doc(edited);

        _ = Align(l, r); // warm-up (JIT and any one-time runtime initialization)
        long start = GC.GetAllocatedBytesForCurrentThread();
        _ = Align(l, r);
        return GC.GetAllocatedBytesForCurrentThread() - start;
    }

    private static string Text(IrBlock b) =>
        b is IrParagraph p
            ? string.Concat(p.Inlines.OfType<IrTextRun>().Select(t => t.Text))
            : string.Empty;
}

/// <summary>
/// Keeps the scale guard isolated from the default parallel suite so unrelated test work cannot
/// perturb process-wide runtime state during its measurement.
/// </summary>
[CollectionDefinition(Name, DisableParallelization = true)]
public sealed class IrAlignerPerformanceCollection
{
    public const string Name = "IR aligner performance";
}
