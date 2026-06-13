#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.6 split-threshold sweep (design-review finding F4.1): the spec's 0.90 coverage / 0.34 foreign-slack
/// values were STARTING HYPOTHESES tuned on two fixtures; this sweep is the GATE that proves they sit on a
/// stable plateau over the corpus before they ship as defaults. The grid re-runs the WC003 revision-count
/// row set (the GetRevisions parity scoreboard's file-based corpus) under every candidate (coverage, slack)
/// pair and counts count-exact passes. The shipped pair must attain the grid MAXIMUM and every NEIGHBORING
/// cell (±1 step on each axis) must attain the same count — i.e. the choice has ≥1 full grid step of margin
/// to the nearest flip, the margin report the review required. If this test ever fails after a corpus or
/// engine change, the plateau has moved: re-run the sweep, re-pin the values, and update the spec — do NOT
/// weaken the assertion.
/// </summary>
[Trait("Category", "Parity")]
public class IrSplitThresholdSweepTests
{
    private readonly ITestOutputHelper _out;
    public IrSplitThresholdSweepTests(ITestOutputHelper output) => _out = output;

    private static readonly DirectoryInfo SourceDir = new("../../../../TestFiles/");

    private static readonly double[] CoverageGrid = { 0.80, 0.85, 0.88, 0.90, 0.92, 0.95 };
    private static readonly double[] SlackGrid = { 0.20, 0.27, 0.34, 0.40, 0.50 };

    private const double ShippedCoverage = 0.90;
    private const double ShippedSlack = 0.34;

    [Fact]
    public void Sweep_split_thresholds_over_the_scoreboard_corpus()
    {
        // Cache the IR reads — each (left, right) pair is read once and re-diffed per grid cell (the
        // aligner + renderer are pure functions of the documents + settings).
        var rows = IrParityScoreboardTests.WC003_Compare_Rows().ToList();
        var docs = new Dictionary<string, IrDocument>(StringComparer.Ordinal);
        IrDocument Read(string rel)
        {
            if (!docs.TryGetValue(rel, out var doc))
                docs[rel] = doc = IrReader.Read(
                    new WmlDocument(Path.Combine(SourceDir.FullName, rel)), WcCorpus.ReadOpts);
            return doc;
        }

        var baseSettings = IrWmlComparerAdapter.MapSettings(new WmlComparerSettings());
        var passes = new int[CoverageGrid.Length, SlackGrid.Length];

        for (int ci = 0; ci < CoverageGrid.Length; ci++)
        for (int si = 0; si < SlackGrid.Length; si++)
        {
            var settings = baseSettings with
            {
                DetectSplitMerge = true,
                SplitCoverageThreshold = CoverageGrid[ci],
                SplitForeignSlack = SlackGrid[si],
            };
            int pass = 0;
            foreach (var (_, left, right, expected) in rows)
            {
                try
                {
                    var l = Read(left);
                    var r = Read(right);
                    var script = IrEditScriptBuilder.Build(l, r, settings);
                    var revs = IrRevisionRenderer.Render(script, l, r, settings);
                    if (revs.Count == expected)
                        pass++;
                }
                catch
                {
                    // A throwing row counts as a non-pass for this cell (same row throws for every cell,
                    // so it cannot tilt the plateau comparison).
                }
            }
            passes[ci, si] = pass;
        }

        // ---- report the full grid ----
        _out.WriteLine("WC003 count-exact passes per (coverage, slack) cell:");
        _out.WriteLine("          " + string.Join("  ", SlackGrid.Select(s => $"s={s:F2}")));
        int max = 0;
        for (int ci = 0; ci < CoverageGrid.Length; ci++)
        {
            _out.WriteLine($"  c={CoverageGrid[ci]:F2}  " +
                string.Join("  ", Enumerable.Range(0, SlackGrid.Length).Select(si => $"{passes[ci, si],5}")));
            for (int si = 0; si < SlackGrid.Length; si++)
                max = Math.Max(max, passes[ci, si]);
        }

        int shippedCi = Array.IndexOf(CoverageGrid, ShippedCoverage);
        int shippedSi = Array.IndexOf(SlackGrid, ShippedSlack);
        Assert.True(shippedCi >= 0 && shippedSi >= 0, "shipped thresholds must be grid members");

        // The shipped cell attains the grid maximum…
        Assert.True(passes[shippedCi, shippedSi] == max,
            $"shipped ({ShippedCoverage}, {ShippedSlack}) scores {passes[shippedCi, shippedSi]} < grid max {max} " +
            "— the plateau has moved; re-pin the thresholds (F4.1).");

        // …and every neighboring cell attains the same count (≥1 grid step of margin to the nearest flip).
        for (int dc = -1; dc <= 1; dc++)
        for (int ds = -1; ds <= 1; ds++)
        {
            int ci = shippedCi + dc, si = shippedSi + ds;
            if (ci < 0 || ci >= CoverageGrid.Length || si < 0 || si >= SlackGrid.Length)
                continue;
            Assert.True(passes[ci, si] == max,
                $"neighbor (c={CoverageGrid[ci]}, s={SlackGrid[si]}) scores {passes[ci, si]} ≠ max {max}: the " +
                "shipped thresholds sit at a plateau EDGE, not on a plateau — margin to the nearest flip is " +
                "less than one grid step (F4.1 blocker: re-adjudicate before shipping).");
        }
    }

    [Fact]
    public void Shipped_thresholds_are_pinned()
    {
        var s = new IrDiffSettings();
        Assert.True(s.DetectSplitMerge);
        Assert.Equal(ShippedCoverage, s.SplitCoverageThreshold);
        Assert.Equal(ShippedSlack, s.SplitForeignSlack);
        Assert.Equal(8, s.SplitMaxRunLength);
    }
}
