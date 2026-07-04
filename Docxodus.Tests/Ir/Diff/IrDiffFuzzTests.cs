#nullable enable

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Docxodus.Tests.Ir;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.3 Task 3 — the deterministic generative fuzzer. For each integer seed, <see cref="DiffFuzzer"/>
/// synthesizes a (left, right) document pair via a seeded mutation engine; this test then runs the strongest
/// oracle we own — the IR pipeline's own invariants — and, where the mutation class is cross-engine
/// comparable, a differential spot check against the shipped <see cref="WmlComparer"/>.
///
/// <para><b>Determinism (the binding constraint).</b> Every case is a pure function of its seed: the
/// document, the mutation list, and the diff are fully reproducible from the seed alone (<see cref="Random"/>
/// is always seeded; nothing reads the clock or the environment except the seed-COUNT knob below). A failure
/// therefore dumps just the seed + a one-line mutation list, and <see cref="ReproduceCase"/> regenerates the
/// exact case in a debugger.</para>
///
/// <para><b>(a) Own-oracle invariants — ALWAYS, every case.</b> IrReader both sides (RetainSources=false) →
/// <see cref="IrBlockAligner"/> + <see cref="IrAlignmentAsserts"/> totality/per-kind invariants →
/// <see cref="IrEditScriptBuilder"/> → <see cref="IrEditScriptVerifier"/> apply-verification →
/// <see cref="IrEditScriptJson"/> round-trip record-equality + determinism. ANY failure here is a hard test
/// failure (the seed + mutation list are dumped in the assertion message).</para>
///
/// <para><b>(b) Differential spot check vs WmlComparer — comparable cases only.</b> A case is comparable iff
/// every mutation is a text edit / paragraph insert-delete / table-cell edit / row insert-delete
/// (<see cref="DiffFuzzer.FuzzCase.IsComparableClass"/>). Cases containing <c>RelocateParagraph</c> (the two
/// engines disagree on move-vs-insert+delete framing) or <c>BoldWord</c> (format-change reporting differs —
/// WmlComparer emits rPrChange revisions the IR renderer models differently) are EXCLUDED, because the
/// cross-engine equivalence relation does not hold for those kinds by construction. For a comparable case we
/// compare under the Task 2 combined-char-bag equivalence (<see cref="RevisionEquivalence"/>). A mismatch is
/// NOT an automatic failure — the two engines legitimately atomize and report at different grains (and
/// WmlComparer has documented oracle quirks, e.g. its U+2011/U+00AD/PUA special-char drops). (The former
/// "French-apostrophe under-report" example here was retracted 2026-06-11: that was an IR NBSP-tokenizer
/// bug, since fixed, not an oracle under-report.) We FAIL ONLY on the one asymmetric signal that
/// is a genuine regression: the NEW engine surfaced NOTHING while the OLD engine saw real content
/// (one-sided-new-empty). All other mismatches are counted, characterized, and written to the artifacts
/// dir.</para>
///
/// <para><b>Scale knob.</b> The seed count defaults to 50 (CI). Set the <c>DOCXODUS_FUZZ_SEEDS</c> env var to
/// override (e.g. 500 for a nightly-style local run). Seeds are always 1..N, so a larger N is a superset of a
/// smaller one — a seed's identity never shifts.</para>
///
/// <para><b>Minimization / repro.</b> On any own-oracle failure the assertion message carries the seed,
/// base-doc paragraph count, table presence, and the full mutation list — enough to call
/// <see cref="ReproduceCase"/>(seed) in a debugger and step the exact case. <see cref="ReproduceCase"/>
/// regenerates the case and re-runs the own-oracle battery, throwing at the first broken invariant.</para>
/// </summary>
[Trait("Category", "Fuzz")]
public class IrDiffFuzzTests
{
    private const int DefaultSeedCount = 50;
    private const string Author = "Open-Xml-PowerTools";

    private static readonly IrReaderOptions ReadOpts =
        new() { RetainSources = false, RevisionView = RevisionView.Accept };
    private static readonly IrDiffSettings Diff = new();

    private readonly ITestOutputHelper _out;

    public IrDiffFuzzTests(ITestOutputHelper output) => _out = output;

    // ---------------------------------------------------------------------- the fuzz run

    [Fact]
    public void Seeded_fuzz_cases_satisfy_own_oracle_and_differential_invariants()
    {
        int seedCount = ResolveSeedCount();
        var sw = Stopwatch.StartNew();
        var artifactsDir = ArtifactsDir();
        Directory.CreateDirectory(artifactsDir);
        ClearStaleArtifacts(artifactsDir);

        // Differential characterization counters (NOT pass/fail — the report adjudicates these).
        int comparable = 0, diffMatch = 0, diffCharBag = 0, diffMismatch = 0, oldEngineThrew = 0;
        int withTable = 0, totalMutations = 0;
        var mismatchExamples = new List<string>();
        var regressions = new List<string>();

        for (int seed = 1; seed <= seedCount; seed++)
        {
            var c = DiffFuzzer.Generate(seed);
            if (c.HasTable) withTable++;
            totalMutations += c.Mutations.Count;

            // ---- (a) own-oracle invariants — ALWAYS. A throw/assert here fails the whole test. ----------
            IrDocument left, right;
            IrEditScript script;
            try
            {
                (left, right, script) = RunOwnOracle(c);
            }
            catch (Exception ex)
            {
                Assert.Fail(
                    $"OWN-ORACLE failure on seed {seed}.\n" +
                    $"  repro: IrDiffFuzzTests.ReproduceCase({seed})\n" +
                    $"  base paragraphs = {c.BaseParagraphCount}, table = {c.HasTable}\n" +
                    $"  mutations = [{c.DescribeMutations()}]\n" +
                    $"  {ex.GetType().Name}: {ex.Message}");
                throw; // unreachable; satisfies the definite-assignment analyzer
            }

            // ---- (b) differential spot check — comparable cases only. -------------------------------------
            if (!c.IsComparableClass)
                continue;
            comparable++;

            RevisionEquivalence.RevisionBag oldBag;
            try
            {
                oldBag = RunOldEngine(c);
            }
            catch (Exception ex)
            {
                // The legacy engine can throw on a synthetic shape; that is not a NEW-engine regression, so we
                // record + skip the differential (own-oracle already passed). Mirrors the Task 2 OLD_ERROR path.
                oldEngineThrew++;
                File.WriteAllText(Path.Combine(artifactsDir, $"seed{seed:D4}.olderror.txt"),
                    $"seed {seed} OLD_ENGINE threw (differential skipped; own-oracle PASSED)\n" +
                    $"mutations = [{c.DescribeMutations()}]\n\n{ex}");
                continue;
            }

            var newRevs = IrRevisionRenderer.Render(script, left, right, Diff);
            var newBag = RevisionEquivalence.RevisionBag.FromIr(newRevs);

            if (oldBag.MultisetsEqual(newBag))
            {
                diffMatch++;
            }
            else if (oldBag.IsCombinedCharBagEquivalent(newBag))
            {
                diffCharBag++; // same content, different atomization — expected and fine
            }
            else
            {
                diffMismatch++;
                // The ONLY hard-failure signal: the new engine missed content the old engine saw. Everything
                // else (old under-reports, granularity beyond the char bag, punctuation attachment) is a
                // characterized non-failure.
                bool oneSidedNewEmpty = newBag.Total == 0 && oldBag.Total > 0;
                WriteMismatchDetail(artifactsDir, seed, c, oldBag, newBag, oneSidedNewEmpty);
                if (oneSidedNewEmpty)
                    regressions.Add($"seed {seed}: new=0 old={oldBag.Total}  [{c.DescribeMutations()}]");
                else if (mismatchExamples.Count < 12)
                    mismatchExamples.Add($"seed {seed}: old={oldBag.Total} new={newBag.Total}  [{c.DescribeMutations()}]");
            }
        }

        sw.Stop();

        // ----- report -------------------------------------------------------------------------------
        _out.WriteLine($"Fuzz run: {seedCount} seeds (env DOCXODUS_FUZZ_SEEDS overrides; default {DefaultSeedCount})");
        _out.WriteLine($"Wall time: {sw.Elapsed.TotalSeconds:F1}s   ({sw.Elapsed.TotalMilliseconds / seedCount:F1} ms/seed)");
        _out.WriteLine($"Cases with a table: {withTable}/{seedCount}   total mutations applied: {totalMutations}");
        _out.WriteLine("");
        _out.WriteLine("OWN-ORACLE: all seeds passed alignment + apply-verify + JSON round-trip.");
        _out.WriteLine("");
        _out.WriteLine("DIFFERENTIAL (comparable cases only):");
        _out.WriteLine($"  comparable cases      = {comparable}");
        _out.WriteLine($"  exact multiset MATCH  = {diffMatch}");
        _out.WriteLine($"  char-bag equivalent   = {diffCharBag}");
        _out.WriteLine($"  mismatch (counted)    = {diffMismatch}");
        _out.WriteLine($"    of which new-empty REGRESSIONS = {regressions.Count}");
        _out.WriteLine($"  old-engine threw      = {oldEngineThrew} (differential skipped, own-oracle passed)");
        if (mismatchExamples.Count > 0)
        {
            _out.WriteLine("");
            _out.WriteLine("  sample non-regression mismatches:");
            foreach (var e in mismatchExamples)
                _out.WriteLine($"    {e}");
        }
        _out.WriteLine("");
        _out.WriteLine($"Artifacts: {artifactsDir}");

        // ----- assertions ---------------------------------------------------------------------------
        // (Own-oracle failures already threw above.) The only differential hard-failure is a one-sided
        // new-empty regression: the IR engine missed content the old engine reported on a COMPARABLE case.
        Assert.True(regressions.Count == 0,
            $"NEW engine reported ZERO revisions on {regressions.Count} comparable case(s) where the OLD " +
            $"engine saw content — a regression:\n  " + string.Join("\n  ", regressions));
    }

    // ---------------------------------------------------------------------- own-oracle battery

    /// <summary>
    /// Run the full own-oracle battery for a case and return the IR docs + script (so the caller can render
    /// revisions for the differential check without re-reading). Throws at the first broken invariant.
    /// </summary>
    private static (IrDocument Left, IrDocument Right, IrEditScript Script) RunOwnOracle(DiffFuzzer.FuzzCase c)
    {
        var left = IrReader.Read(c.Left, ReadOpts);
        var right = IrReader.Read(c.Right, ReadOpts);

        // Alignment totality + per-kind hash/format invariants.
        var alignment = IrBlockAligner.Align(left, right, Diff);
        IrAlignmentAsserts.AssertInvariants(left, right, alignment, Diff);

        // Edit script + apply-verification (apply(script, left) reconstructs right at text level; also
        // re-checks alignment anchors, move pairing, and nested table diffs).
        var script = IrEditScriptBuilder.Build(left, right, Diff);
        IrEditScriptVerifier.Verify(left, right, script, Diff);

        // JSON round-trip: Read(Write(s)) is record-equal to s, and Write is deterministic.
        var json = IrEditScriptJson.Write(script);
        var back = IrEditScriptJson.Read(json);
        Assert.Equal(script, back);
        Assert.Equal(json, IrEditScriptJson.Write(back));

        return (left, right, script);
    }

    private static RevisionEquivalence.RevisionBag RunOldEngine(DiffFuzzer.FuzzCase c)
    {
        var settings = new WmlComparerSettings { AuthorForRevisions = Author };
        var compared = WmlComparer.Compare(c.Left, c.Right, settings);
        var revs = WmlComparer.GetRevisions(compared, settings);
        return RevisionEquivalence.RevisionBag.FromWmlComparer(revs);
    }

    // ---------------------------------------------------------------------- repro affordance

    /// <summary>
    /// Regenerate the case for <paramref name="seed"/> and re-run the own-oracle battery, throwing at the
    /// first broken invariant. The minimization affordance: a failing fuzz seed dumped by the main test can
    /// be reproduced with a single call here (e.g. in a debugger or an ad-hoc <c>[Fact]</c>):
    /// <code>IrDiffFuzzTests.ReproduceCase(42);</code>
    /// </summary>
    public static void ReproduceCase(int seed) => ReproduceCaseInternal(seed);

    /// <summary>
    /// As <see cref="ReproduceCase"/>, but returns the resolved <see cref="DiffFuzzer.FuzzCase"/> so a
    /// caller can inspect the documents / mutation list after the (green) own-oracle re-run. Internal
    /// because <see cref="DiffFuzzer.FuzzCase"/> is internal; the public <see cref="ReproduceCase"/> is the
    /// debugger entry point named in failure dumps.
    /// </summary>
    internal static DiffFuzzer.FuzzCase ReproduceCaseInternal(int seed)
    {
        var c = DiffFuzzer.Generate(seed);
        RunOwnOracle(c);
        return c;
    }

    /// <summary>
    /// A standing smoke for the repro affordance itself (and a regression guard for a few specific seeds):
    /// <see cref="ReproduceCase"/> must pass for these seeds, proving the helper is wired and the engine is
    /// green on a fixed sample independent of the env-driven main run.
    /// </summary>
    [Theory]
    [InlineData(1)]
    [InlineData(7)]
    [InlineData(13)]
    [InlineData(42)]
    public void ReproduceCase_is_green_for_fixed_seeds(int seed)
    {
        var c = ReproduceCaseInternal(seed);
        Assert.Equal(seed, c.Seed);
        Assert.InRange(c.BaseParagraphCount, 10, 40);
    }

    // ---------------------------------------------------------------------- knobs + artifacts

    private int ResolveSeedCount()
    {
        var raw = Environment.GetEnvironmentVariable("DOCXODUS_FUZZ_SEEDS");
        if (!string.IsNullOrWhiteSpace(raw) && int.TryParse(raw, out var n) && n > 0)
        {
            _out.WriteLine($"DOCXODUS_FUZZ_SEEDS={n} (overriding default {DefaultSeedCount})");
            return n;
        }
        return DefaultSeedCount;
    }

    private static void WriteMismatchDetail(
        string dir, int seed, DiffFuzzer.FuzzCase c,
        RevisionEquivalence.RevisionBag oldB, RevisionEquivalence.RevisionBag newB, bool regression)
    {
        var sb = new StringBuilder();
        sb.AppendLine(regression ? "MISMATCH (REGRESSION: new-empty)" : "MISMATCH (characterized non-failure)");
        sb.AppendLine($"seed = {seed}    repro: IrDiffFuzzTests.ReproduceCase({seed})");
        sb.AppendLine($"base paragraphs = {c.BaseParagraphCount}, table = {c.HasTable}");
        sb.AppendLine($"mutations = [{c.DescribeMutations()}]");
        sb.AppendLine($"old total = {oldB.Total}    new total = {newB.Total}");
        sb.AppendLine();
        foreach (var kind in Enum.GetValues<IrRevisionType>())
        {
            sb.AppendLine($"[{kind}]");
            sb.AppendLine($"  OLD: {oldB.Dump(kind)}");
            sb.AppendLine($"  NEW: {newB.Dump(kind)}");
        }
        File.WriteAllText(Path.Combine(dir, $"seed{seed:D4}.mismatch.txt"), sb.ToString());
    }

    private static void ClearStaleArtifacts(string dir)
    {
        foreach (var f in Directory.GetFiles(dir, "seed*.txt"))
            File.Delete(f);
    }

    private static string ArtifactsDir([CallerFilePath] string thisFile = "") =>
        Path.Combine(Path.GetDirectoryName(thisFile)!, "FuzzArtifacts");
}
