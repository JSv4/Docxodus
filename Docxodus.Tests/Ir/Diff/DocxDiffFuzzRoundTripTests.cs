#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using Docxodus;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Generative content-correctness fuzzer for the BYTE-LEVEL redline round trip — the exact accept≡right /
/// reject≡left contract a consumer relies on. For each seed, <see cref="DiffFuzzer.Generate"/> synthesizes a
/// random (left, right) document pair; this test then runs the FULL public path
/// <c>DocxDiff.Compare(left, right)</c> → <see cref="RevisionProcessor.AcceptRevisions(WmlDocument)"/> /
/// <see cref="RevisionProcessor.RejectRevisions(WmlDocument)"/> and asserts the accepted document's whole-body
/// text (paragraphs AND tables) equals the RIGHT input's, and the rejected document's equals the LEFT input's —
/// with zero content lost or mangled.
/// <para><b>Why this exists.</b> The <see cref="IrDiffFuzzTests"/> fuzzer validates content correctness at the
/// EDIT-SCRIPT level (<see cref="IrEditScriptVerifier"/> reconstructs right at the token level) — a strong
/// proxy, but it does not materialize the redline docx and round-trip it through Accept/Reject. Before this
/// test that byte-level path was only spot-checked on a handful of hand-built pairs
/// (<c>DocxDiffOpsRoundTripTests</c>, <c>DocxDiffInputRevisionsRoundTripTests</c>). This turns those few into
/// hundreds of reproducible fuzzed cases (seed-count knob <c>DOCXODUS_FUZZ_SEEDS</c>, default 50), the same
/// determinism + repro discipline as the sibling fuzzer.</para>
/// <para>Inputs are CLEAN (the generator seeds no pre-existing revisions), so the accepted view is the raw text:
/// <c>accept(Compare(l,r))</c> text == right text and <c>reject</c> text == left text directly.</para>
/// </summary>
public class DocxDiffFuzzRoundTripTests
{
    /// <summary>
    /// Default sweep width. Raised from 50 with the order assertion (issue #288): the reordering seeds are
    /// rare — 3 in the first 2000 — and 50 seeds never reached one, so the bug lived under a green fuzzer.
    /// 250 seeds (~8s) covers seed 184, the canonical duplicate-paragraph-crossing-a-merge repro; the wide
    /// 2000-seed sweep that also finds 760 and 1714 runs via <c>DOCXODUS_FUZZ_SEEDS</c>.
    /// </summary>
    private const int DefaultSeedCount = 250;
    private readonly ITestOutputHelper _out;

    public DocxDiffFuzzRoundTripTests(ITestOutputHelper output) => _out = output;

    private static int ResolveSeedCount()
    {
        var env = Environment.GetEnvironmentVariable("DOCXODUS_FUZZ_SEEDS");
        return int.TryParse(env, out var n) && n > 0 ? n : DefaultSeedCount;
    }

    [Fact]
    public void Fuzz_byte_level_accept_reject_round_trip_preserves_content()
    {
        int seedCount = ResolveSeedCount();
        var failures = new List<string>();
        int withTable = 0;

        for (int seed = 1; seed <= seedCount; seed++)
        {
            var c = DiffFuzzer.Generate(seed);

            var expectedRight = Docs.PlainTextWithTables(c.Right);
            var expectedLeft = Docs.PlainTextWithTables(c.Left);
            if (HasTable(c.Right) || HasTable(c.Left)) withTable++;

            WmlDocument redline;
            WmlDocument accepted;
            WmlDocument rejected;
            try
            {
                redline = DocxDiff.Compare(c.Left, c.Right);
                accepted = RevisionProcessor.AcceptRevisions(redline);
                rejected = RevisionProcessor.RejectRevisions(redline);
            }
            catch (Exception ex)
            {
                failures.Add($"seed {seed}: THREW {ex.GetType().Name}: {ex.Message}  [{c.DescribeMutations()}]");
                continue;
            }

            // HARD guarantee — the primary contract: NO CONTENT LOSS. The accepted document's word bag equals
            // the right input's and the rejected document's equals the left input's (order-independent, so a
            // move/split reorder does not mask a genuine drop or duplication — a lost/duplicated word flips the
            // multiset). This is the guarantee a consumer relies on: accept keeps exactly the revised content,
            // reject exactly the original content, nothing added or dropped.
            if (!WordBagEqual(gotRight: Docs.PlainTextWithTables(accepted), expectedRight))
                failures.Add($"seed {seed}: ACCEPT lost/added content vs right  [{c.DescribeMutations()}]\n" +
                             BagDelta(expectedRight, Docs.PlainTextWithTables(accepted)));
            if (!WordBagEqual(gotRight: Docs.PlainTextWithTables(rejected), expectedLeft))
                failures.Add($"seed {seed}: REJECT lost/added content vs left  [{c.DescribeMutations()}]\n" +
                             BagDelta(expectedLeft, Docs.PlainTextWithTables(rejected)));

            // EXACT-ORDER guarantee (issue #288). The word-bag checks above are order-INDEPENDENT by
            // design, so they cannot see a redline that restores every word but rebuilds the blocks
            // permuted — which is precisely what a duplicate-content pairing crossing a split/merge group
            // used to do (seeds 184, 760 and 1714 of the first 2000). Comparing the whole-body
            // BLOCK SEQUENCE, not just the multiset, is what turns `accept ≡ right` / `reject ≡ left` into
            // the exact contract a consumer relies on.
            if (BlockSequence(accepted) != BlockSequence(c.Right))
                failures.Add($"seed {seed}: ACCEPT block ORDER differs from right  [{c.DescribeMutations()}]\n" +
                             FirstOrderDelta(BlockSequence(c.Right), BlockSequence(accepted)));
            if (BlockSequence(rejected) != BlockSequence(c.Left))
                failures.Add($"seed {seed}: REJECT block ORDER differs from left  [{c.DescribeMutations()}]\n" +
                             FirstOrderDelta(BlockSequence(c.Left), BlockSequence(rejected)));
        }

        _out.WriteLine($"Byte-level round-trip fuzz: {seedCount} seeds ({withTable} with a table), " +
                       $"zero content loss AND exact block order. " +
                       $"env DOCXODUS_FUZZ_SEEDS overrides (default {DefaultSeedCount}).");
        Assert.True(failures.Count == 0,
            $"{failures.Count} byte-level accept/reject CONTENT or ORDER failures:\n" +
            string.Join("\n", failures.Take(15)));
    }

    /// <summary>
    /// The document's whole body as one newline-joined, whitespace-normalized BLOCK sequence — the
    /// order-SENSITIVE projection <see cref="WordBagEqual"/> deliberately discards. Empty lines are dropped
    /// so a redline is judged on the blocks that carry content, not on paragraph-mark bookkeeping (an
    /// inserted or deleted bare mark is a separate, already-covered concern).
    /// </summary>
    private static string BlockSequence(WmlDocument d) =>
        string.Join("\n", Docs.PlainTextWithTables(d)
            .Split('\n')
            .Select(line => string.Join(" ", line.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)))
            .Where(line => line.Length > 0));

    /// <summary>The first differing block of two <see cref="BlockSequence"/> projections, for the repro dump.</summary>
    private static string FirstOrderDelta(string expected, string actual)
    {
        var e = expected.Split('\n');
        var a = actual.Split('\n');
        for (int i = 0; i < Math.Max(e.Length, a.Length); i++)
        {
            string ei = i < e.Length ? e[i] : "<none>";
            string ai = i < a.Length ? a[i] : "<none>";
            if (ei != ai)
                return $"      first divergence at block {i}:\n        expected: [{ei}]\n        actual:   [{ai}]";
        }
        return $"      block counts differ: expected {e.Length}, actual {a.Length}";
    }

    private static bool WordBagEqual(string gotRight, string expected)
    {
        var a = new Dictionary<string, int>();
        foreach (var w in gotRight.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
            a[w] = a.GetValueOrDefault(w) + 1;
        foreach (var w in expected.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
        {
            if (!a.TryGetValue(w, out var n) || n == 0) return false;
            a[w] = n - 1;
        }
        return a.Values.All(v => v == 0);
    }

    private static string BagDelta(string expected, string actual)
    {
        Dictionary<string, int> Bag(string s)
        {
            var d = new Dictionary<string, int>();
            foreach (var w in s.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))
                d[w] = d.GetValueOrDefault(w) + 1;
            return d;
        }
        var e = Bag(expected); var a = Bag(actual);
        var lost = e.Where(kv => a.GetValueOrDefault(kv.Key) < kv.Value)
                    .Select(kv => $"{kv.Key}×{kv.Value - a.GetValueOrDefault(kv.Key)}").Take(8);
        var extra = a.Where(kv => e.GetValueOrDefault(kv.Key) < kv.Value)
                     .Select(kv => $"{kv.Key}×{kv.Value - e.GetValueOrDefault(kv.Key)}").Take(8);
        return $"      lost: [{string.Join(" ", lost)}]  extra: [{string.Join(" ", extra)}]";
    }

    private static bool HasTable(WmlDocument d) => Docs.MainPartXml(d).Contains("<w:tbl>");

}
