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
/// oracle we own — the IR pipeline's own invariants — and, where the mutation class warrants it, checks
/// that the engine actually produces revisions.
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
/// <para><b>(b) Revision production — comparable cases only.</b> A case is comparable iff every mutation
/// is a text edit / paragraph insert-delete / table-cell edit / row insert-delete
/// (<see cref="DiffFuzzer.FuzzCase.IsComparableClass"/>). Cases containing <c>RelocateParagraph</c> or
/// <c>BoldWord</c> are excluded, as they were when this half was a differential check against the legacy
/// engine and the cross-engine equivalence did not hold for those kinds by construction.
///
/// <para>Through v10 this half compared the IR engine's revision set against <c>WmlComparer</c>'s, failing
/// only on one asymmetric signal: the new engine surfaced NOTHING where the legacy engine saw real content.
/// The legacy engine went in v11.0.0, but that signal does not need it — a comparable case producing zero
/// revisions is the regression, whatever a second engine would have said. Because every case is a pure
/// function of its seed, the number of comparable cases that legitimately yield nothing is a constant
/// (<see cref="ExpectedZeroRevisionCases"/>), so it is PINNED rather than merely reported: more means the
/// engine is missing content, fewer means it improved and the pin must move.</para></para>
/// </summary>
[Trait("Category", "Fuzz")]
public class IrDiffFuzzTests
{
    private const int DefaultSeedCount = 50;

    /// <summary>
    /// Comparable cases that legitimately produce no revisions at <see cref="DefaultSeedCount"/> seeds.
    /// Deterministic (every case is a pure function of its seed), so it is asserted, not just reported.
    /// Measured when the differential half was retired in v11.0.0.
    /// </summary>
    private const int ExpectedZeroRevisionCases = 0;
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

        int comparable = 0, zeroRevisionCases = 0;
        int withTable = 0, totalMutations = 0;
        var zeroRevisionExamples = new List<string>();

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

            // ---- (b) revision-production check — comparable cases only. ----------------------------
            // Through v10 this was a DIFFERENTIAL check against WmlComparer, whose only hard-failure
            // signal was "the IR engine reported zero revisions on a comparable case where the legacy
            // engine saw content". With the legacy engine removed the signal survives without it: the
            // fuzzer is seeded and deterministic, so the set of comparable cases that legitimately yield
            // no revisions is a fixed number. Pin it. A NEW zero-revision case is the same regression the
            // differential caught; a case that stops being zero is an improvement that must be re-pinned.
            if (!c.IsComparableClass)
                continue;
            comparable++;

            var newRevs = IrRevisionRenderer.Render(script, left, right, Diff);
            var newBag = RevisionEquivalence.RevisionBag.FromIr(newRevs);
            if (newBag.Total == 0)
            {
                zeroRevisionCases++;
                if (zeroRevisionExamples.Count < 12)
                    zeroRevisionExamples.Add($"seed {seed}: [{c.DescribeMutations()}]");
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
        _out.WriteLine("REVISION PRODUCTION (comparable cases only):");
        _out.WriteLine($"  comparable cases          = {comparable}");
        _out.WriteLine($"  yielding zero revisions   = {zeroRevisionCases} (pinned at {ExpectedZeroRevisionCases})");
        if (zeroRevisionExamples.Count > 0)
        {
            _out.WriteLine("");
            _out.WriteLine("  sample zero-revision cases:");
            foreach (var e in zeroRevisionExamples)
                _out.WriteLine($"    {e}");
        }
        _out.WriteLine("");
        _out.WriteLine($"Artifacts: {artifactsDir}");

        // ----- assertions ---------------------------------------------------------------------------
        // (Own-oracle failures already threw above.) A comparable case producing no revisions at all is
        // the regression the removed differential check existed to catch. The count is deterministic for
        // a fixed seed count, so it is pinned rather than merely reported.
        Assert.True(
            seedCount != DefaultSeedCount || zeroRevisionCases == ExpectedZeroRevisionCases,
            $"Comparable cases yielding ZERO revisions moved from {ExpectedZeroRevisionCases} to " +
            $"{zeroRevisionCases}. More means the engine is missing content it used to report; fewer means " +
            $"it improved and the pin needs updating. Samples:\n  " +
            string.Join("\n  ", zeroRevisionExamples));
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


    private static void ClearStaleArtifacts(string dir)
    {
        foreach (var f in Directory.GetFiles(dir, "seed*.txt"))
            File.Delete(f);
    }

    private static string ArtifactsDir([CallerFilePath] string thisFile = "") =>
        Path.Combine(Path.GetDirectoryName(thisFile)!, "FuzzArtifacts");
}
