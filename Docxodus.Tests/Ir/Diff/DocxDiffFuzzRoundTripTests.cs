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
    private const int DefaultSeedCount = 50;
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
        }

        _out.WriteLine($"Byte-level round-trip content fuzz: {seedCount} seeds ({withTable} with a table), " +
                       $"zero content loss. env DOCXODUS_FUZZ_SEEDS overrides (default {DefaultSeedCount}).");
        Assert.True(failures.Count == 0,
            $"{failures.Count} byte-level accept/reject CONTENT-LOSS failures:\n" +
            string.Join("\n", failures.Take(15)));
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
