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
        // No exact-hash anchors exist; everything resolves in one big head↔tail gap as positional
        // Modified pairs (200) with no surplus, zero Moved.
        Assert.Equal(200, Count(a, IrAlignmentKind.Modified));
        Assert.Equal(0, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(0, Count(a, IrAlignmentKind.Moved));
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
    public void Scale_guard_500_vs_2000_wall_ratio_within_8x()
    {
        // Both inputs are the near-identical fixture (distinct clauses) self-paired with ONE edit, so
        // every block anchors uniquely and the only gap is a single 1-block Modified gap — i.e. NO large
        // all-distinct gap that would trip the InOrderRefine G²/2 worst case. This isolates the spine /
        // anchoring cost, which should scale ~linearly: 4× the blocks ⇒ well under 8× the wall time.
        double small = BestOfThree(500);
        double large = BestOfThree(2000);
        double ratio = large / Math.Max(small, 0.0001);

        _out.WriteLine($"Scale guard: 500-para = {small:F2} ms, 2000-para = {large:F2} ms, ratio = {ratio:F2}x (n=4x)");
        Assert.True(ratio <= 8.0,
            $"Align wall-time ratio {ratio:F2}x for 4x input exceeds the 8x anti-O(n²) guard " +
            $"(500={small:F2}ms, 2000={large:F2}ms).");
    }

    /// <summary>Warm up once, then best-of-3 wall-time (ms) for an n-para self-pair with one edit.</summary>
    private static double BestOfThree(int n)
    {
        var baseParas = DistinctClauses(n);
        var edited = (string[])baseParas.Clone();
        edited[n / 2] = $"Clause {n / 2}: REVISED wording for this section of the agreement.";

        var l = Doc(baseParas);
        var r = Doc(edited);

        _ = Align(l, r); // warm-up (JIT, dictionary growth)

        double best = double.MaxValue;
        for (int i = 0; i < 3; i++)
        {
            var sw = Stopwatch.StartNew();
            _ = Align(l, r);
            sw.Stop();
            best = Math.Min(best, sw.Elapsed.TotalMilliseconds);
        }
        return best;
    }

    private static string Text(IrBlock b) =>
        b is IrParagraph p
            ? string.Concat(p.Inlines.OfType<IrTextRun>().Select(t => t.Text))
            : string.Empty;
}
