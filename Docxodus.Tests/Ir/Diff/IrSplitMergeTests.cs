#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Docxodus.Tests.Ir;
using Xunit;
using Xunit.Abstractions;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>M2.6 split/merge unit tests — op model + JSON wire (extended by later tasks: segmenter, detection, projection).</summary>
public class IrSplitMergeTests
{
    private readonly ITestOutputHelper _out;

    public IrSplitMergeTests(ITestOutputHelper output) => _out = output;

    private static IrTokenDiff Diff(params IrTokenOp[] ops) => new(IrNodeList.From(ops.ToList()));

    private static IrEditOp SplitOp() => new(
        IrEditOpKind.SplitBlock,
        LeftAnchor: "p:body:aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
        RightAnchor: null, TokenDiff: null, MoveGroupId: null, IsMoveSource: null,
        SplitMergeAnchors: IrNodeList.From(new List<string>
        {
            "p:body:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb",
            "p:body:cccccccccccccccccccccccccccccccc",
        }),
        SegmentDiffs: IrNodeList.From(new List<IrTokenDiff>
        {
            Diff(new IrTokenOp(IrTokenOpKind.Equal, 0, 3, 0, 3)),
            Diff(new IrTokenOp(IrTokenOpKind.Equal, 0, 2, 0, 2),
                 new IrTokenOp(IrTokenOpKind.Insert, 2, 2, 2, 4)),
        }));

    private static IrEditOp MergeOp() => new(
        IrEditOpKind.MergeBlock,
        LeftAnchor: null,
        RightAnchor: "p:body:99999999999999999999999999999999",
        TokenDiff: null, MoveGroupId: null, IsMoveSource: null,
        SplitMergeAnchors: IrNodeList.From(new List<string>
        {
            "p:body:11111111111111111111111111111111",
            "p:body:22222222222222222222222222222222",
        }),
        SegmentDiffs: IrNodeList.From(new List<IrTokenDiff>
        {
            Diff(new IrTokenOp(IrTokenOpKind.Equal, 0, 2, 0, 2)),
            Diff(new IrTokenOp(IrTokenOpKind.Equal, 0, 2, 2, 4)),
        }));

    [Fact]
    public void Split_op_json_round_trips_and_is_deterministic()
    {
        var script = new IrEditScript(IrNodeList.From(new List<IrEditOp> { SplitOp() }));
        var json = IrEditScriptJson.Write(script);
        var back = IrEditScriptJson.Read(json);
        Assert.Equal(script, back);
        Assert.Equal(json, IrEditScriptJson.Write(back));
    }

    [Fact]
    public void Split_op_json_golden_shape()
    {
        var json = IrEditScriptJson.Write(new IrEditScript(IrNodeList.From(new List<IrEditOp> { SplitOp() })));
        Assert.Contains("\"kind\": \"SplitBlock\"", json);
        Assert.Contains("\"splitMergeAnchors\"", json);
        Assert.Contains("\"segmentDiffs\"", json);
        // The singular side rides the EXISTING field; no rightAnchor on a split op.
        Assert.Contains("\"leftAnchor\"", json);
        Assert.DoesNotContain("\"rightAnchor\"", json);
    }

    [Fact]
    public void Merge_op_json_round_trips_and_is_deterministic()
    {
        var script = new IrEditScript(IrNodeList.From(new List<IrEditOp> { MergeOp() }));
        var json = IrEditScriptJson.Write(script);
        var back = IrEditScriptJson.Read(json);
        Assert.Equal(script, back);
        Assert.Equal(json, IrEditScriptJson.Write(back));
    }

    [Fact]
    public void Merge_op_json_golden_shape()
    {
        var json = IrEditScriptJson.Write(new IrEditScript(IrNodeList.From(new List<IrEditOp> { MergeOp() })));
        Assert.Contains("\"kind\": \"MergeBlock\"", json);
        Assert.Contains("\"splitMergeAnchors\"", json);
        Assert.Contains("\"segmentDiffs\"", json);
        Assert.Contains("\"rightAnchor\"", json);
        Assert.DoesNotContain("\"leftAnchor\"", json);
    }

    [Fact]
    public void Scripts_without_splits_serialize_without_new_fields()
    {
        var op = new IrEditOp(IrEditOpKind.InsertBlock, null, "p:body:dddddddddddddddddddddddddddddddd",
            null, null, null);
        var json = IrEditScriptJson.Write(new IrEditScript(IrNodeList.From(new List<IrEditOp> { op })));
        Assert.DoesNotContain("splitMergeAnchors", json);
        Assert.DoesNotContain("segmentDiffs", json);
    }

    [Fact]
    public void Pairing_assert_rejects_a_split_op_that_also_sets_RightAnchor() // F1.1: the assert is load-bearing
    {
        var bad = SplitOp() with { RightAnchor = "p:body:eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee" };
        var script = new IrEditScript(IrNodeList.From(new List<IrEditOp> { bad }));
        Assert.ThrowsAny<Xunit.Sdk.XunitException>(() => IrEditScriptVerifier.AssertSplitMergePairing(script));
    }

    [Fact]
    public void Pairing_assert_rejects_anchor_shared_between_two_split_ops() // F2.2 overlap ceiling
    {
        var a = SplitOp();
        var b = SplitOp() with { LeftAnchor = "p:body:ffffffffffffffffffffffffffffffff" };
        var script = new IrEditScript(IrNodeList.From(new List<IrEditOp> { a, b }));
        Assert.ThrowsAny<Xunit.Sdk.XunitException>(() => IrEditScriptVerifier.AssertSplitMergePairing(script));
    }

    [Fact]
    public void Pairing_assert_rejects_count_mismatch_and_short_anchor_lists()
    {
        var oneAnchor = SplitOp() with
        {
            SplitMergeAnchors = IrNodeList.From(new List<string> { "p:body:bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb" }),
        };
        Assert.ThrowsAny<Xunit.Sdk.XunitException>(() => IrEditScriptVerifier.AssertSplitMergePairing(
            new IrEditScript(IrNodeList.From(new List<IrEditOp> { oneAnchor }))));
    }

    [Fact]
    public void Pairing_assert_rejects_segment_diff_count_mismatch()
    {
        // 2 anchors (passes the ≥2 gate) but only 1 segment diff — the count-equality rule must fire.
        var mismatch = SplitOp() with
        {
            SegmentDiffs = IrNodeList.From(new List<IrTokenDiff>
            {
                Diff(new IrTokenOp(IrTokenOpKind.Equal, 0, 3, 0, 3)),
            }),
        };
        Assert.ThrowsAny<Xunit.Sdk.XunitException>(() => IrEditScriptVerifier.AssertSplitMergePairing(
            new IrEditScript(IrNodeList.From(new List<IrEditOp> { mismatch }))));
    }

    [Fact]
    public void Pairing_assert_accepts_a_well_formed_split_and_merge()
    {
        IrEditScriptVerifier.AssertSplitMergePairing(
            new IrEditScript(IrNodeList.From(new List<IrEditOp> { SplitOp(), MergeOp() })));
    }

    // -------- segmenter (Task 3) --------

    private static (IrDocument Doc, List<IrParagraph> Paras) ReadParas(params string[] texts)
    {
        var xml = string.Concat(texts.Select(t =>
            $"<w:p><w:r><w:t xml:space=\"preserve\">{t}</w:t></w:r></w:p>"));
        var doc = IrReader.Read(IrTestDocuments.FromBodyXml(xml),
            new IrReaderOptions { RetainSources = false, RevisionView = RevisionView.Accept });
        return (doc, doc.Body.Blocks.OfType<IrParagraph>().ToList());
    }

    private static readonly IrDiffSettings S = new() { DetectSplitMerge = true };

    [Fact]
    public void Segmenter_scores_a_clean_split_at_full_coverage_zero_slack()
    {
        var (_, lp) = ReadParas("alpha bravo charlie. delta echo foxtrot.");
        var (_, rp) = ReadParas("alpha bravo charlie. ", "delta echo foxtrot.");
        var score = IrSplitSegmenter.Score(lp[0], new List<IrParagraph> { rp[0], rp[1] }, S);
        Assert.True(score.Coverage >= 0.99, $"coverage {score.Coverage}");
        Assert.True(score.ForeignSlack <= 0.01, $"slack {score.ForeignSlack}");
    }

    [Fact]
    public void Segmenter_scores_keyword_coincidence_below_threshold()
    {
        var (_, lp) = ReadParas("the contract terminates on delivery of the goods.");
        var (_, rp) = ReadParas("the parties agree on many things.", "delivery of pizza is unrelated to the goods here.");
        var score = IrSplitSegmenter.Score(lp[0], new List<IrParagraph> { rp[0], rp[1] }, S);
        Assert.True(score.Coverage < S.SplitCoverageThreshold || score.ForeignSlack > S.SplitForeignSlack,
            $"coincidence must not qualify (cov={score.Coverage}, slack={score.ForeignSlack})");
    }

    [Fact]
    public void Segmenter_segment_diffs_tile_the_left_token_stream_exactly() // F3.3 partition invariant
    {
        var (_, lp) = ReadParas("alpha bravo charlie. delta echo foxtrot.");
        var (_, rp) = ReadParas("alpha bravo charlie. ", "NEW WORDS HERE", "delta echo foxtrot.");
        var rights = new List<IrParagraph> { rp[0], rp[1], rp[2] };
        var diffs = IrSplitSegmenter.ComputeSegmentDiffs(lp[0], rights, S);
        Assert.Equal(3, diffs.Count);
        // F3.3: the segment slices tile the left token stream exactly — slice i's length is the sum
        // of its non-Insert left-span lengths, and the slice lengths sum to the full left token count.
        int leftTotal = IrDiffTokenizer.Tokenize(lp[0], S).Count;
        Assert.Equal(leftTotal, diffs.Sum(d => d.Ops.Where(o => o.Kind != IrTokenOpKind.Insert).Sum(o => o.LeftLength)));
        // And each segment diff right-tiles its right block (IrTokenDiffer invariant, re-checked):
        for (int i = 0; i < 3; i++)
        {
            int rightCount = IrDiffTokenizer.Tokenize(rights[i], S).Count;
            Assert.Equal(rightCount, diffs[i].Ops.Where(o => o.Kind != IrTokenOpKind.Delete).Sum(o => o.RightLength));
        }
    }

    // -------- Task 4 Step 1: permanent fixture diagnostic (evidence for the detection design) --------

    /// <summary>
    /// Dumps the DEFAULT-settings (detection OFF) edit-script op states for the two corpus fixture
    /// pairs whose 1:N split deviations motivated M2.6 (WC-1830, WC-1450), recursing into TableDiff
    /// cell BlockOps. This pins the entry states the detection scan must convert.
    /// </summary>
    /// <remarks>
    /// OBSERVED 2026-06-12 (detection OFF), both deviations live INSIDE one cell's block alignment,
    /// each in a SINGLE gap — the state-(b) shape the Task 4 scan converts:
    /// <list type="bullet">
    /// <item><b>WC-1830</b> (WC041-Table-5 vs Mod), cell tc:7e5725…: the split-out run is
    /// [Insert "Video provides…" | Insert "" (the net-new interior math paragraph, zero content
    /// tokens) | Modify L"Video provides…"↔R"When you click Online Video…"]. I.e. the singular LEFT
    /// paragraph is similarity-paired (Modified) to the LAST run member; the two preceding members
    /// are free Inserts, all adjacent in the same gap. A genuine Delete ("You can also type…")
    /// follows in the same gap and must remain a Delete.</item>
    /// <item><b>WC-1450</b> (WC023-Table-4-Row-Image Before vs After-Delete-1-Row), cell
    /// tc:cd1109…: identical shape — [Insert "Video provides…" | Insert "" | Modify L"Video
    /// provides…"↔R"When you click…"] + trailing genuine Delete. Cell tc:7e5725… additionally shows
    /// the partner-FIRST variant: [Modify L↔R-prefix | Insert "When you click…" (the split tail)].</item>
    /// </list>
    /// No fixture puts the halves in different gaps and no prefix is consumed by a spine anchor
    /// outside the gap, so the in-gap containment scan covers both corpus deviations.
    /// </remarks>
    [Fact]
    public void Diagnostic_fixture_cell_alignment_states()
    {
        DumpFixture("WC-1830", "WC/WC041-Table-5.docx", "WC/WC041-Table-5-Mod.docx");
        DumpFixture("WC-1450", "WC/WC023-Table-4-Row-Image-Before.docx",
            "WC/WC023-Table-4-Row-Image-After-Delete-1-Row.docx");
    }

    private void DumpFixture(string label, string leftRel, string rightRel)
    {
        const string root = "../../../../TestFiles/";
        var settings = new IrDiffSettings { DetectSplitMerge = false }; // detection OFF — baseline states
        var left = IrReader.Read(new WmlDocument(Path.Combine(root, leftRel)), WcCorpus.ReadOpts);
        var right = IrReader.Read(new WmlDocument(Path.Combine(root, rightRel)), WcCorpus.ReadOpts);
        var script = IrEditScriptBuilder.Build(left, right, settings);

        _out.WriteLine($"==== {label}: {leftRel} vs {rightRel} ====");
        DumpOps(script.Operations, left, right, settings, indent: "");
        _out.WriteLine("");

        // Minimal regression canary (the dump itself is the diagnostic): the baseline (detection-OFF)
        // script must still contain the cell-level Modified pairing the state-(b) account documents —
        // if this fails, the recorded entry states above are stale and must be re-derived.
        bool anyCellModify = script.Operations.Any(o => o.TableDiff is { } td2 &&
            td2.RowOps.Any(rw => rw.CellOps is { } cs && cs.Any(c => c.BlockOps is { } bs &&
                bs.Any(b => b.Kind == IrEditOpKind.ModifyBlock))));
        Assert.True(anyCellModify, $"{label}: expected a cell-level ModifyBlock in the baseline script.");
    }

    private void DumpOps(
        IEnumerable<IrEditOp> ops, IrDocument left, IrDocument right, IrDiffSettings settings, string indent)
    {
        foreach (var op in ops)
        {
            _out.WriteLine(
                $"{indent}{op.Kind} L={op.LeftAnchor ?? "-"} [{BlockText(left, op.LeftAnchor, settings)}] " +
                $"R={op.RightAnchor ?? "-"} [{BlockText(right, op.RightAnchor, settings)}]");
            if (op.TableDiff is { } td)
            {
                foreach (var row in td.RowOps)
                {
                    _out.WriteLine($"{indent}  row {row.Kind} L={row.LeftRowAnchor ?? "-"} R={row.RightRowAnchor ?? "-"}");
                    if (row.CellOps is { } cells)
                        foreach (var cell in cells)
                        {
                            _out.WriteLine($"{indent}    cell L={cell.LeftCellAnchor ?? "-"} R={cell.RightCellAnchor ?? "-"}");
                            if (cell.BlockOps is { } blockOps)
                                DumpOps(blockOps, left, right, settings, indent + "      ");
                        }
                }
            }
        }
    }

    private static string BlockText(IrDocument doc, string? anchor, IrDiffSettings settings)
    {
        if (anchor is null || !doc.AnchorIndex.TryGetValue(anchor, out var block))
            return "";
        if (block is not IrParagraph p)
            return $"<{block.GetType().Name}>";
        string text = string.Concat(IrDiffTokenizer.Tokenize(p, settings).Select(t => t.Text));
        return text.Length <= 40 ? text : text[..40] + "…";
    }

    [Fact]
    public void MirrorDiff_swaps_sides_and_flips_insert_delete()
    {
        // A Delete (left-only span) followed by an Equal must mirror to an Insert (right-only span)
        // followed by an Equal with the side spans swapped — the merge path's orientation correction.
        var diff = Diff(
            new IrTokenOp(IrTokenOpKind.Delete, 0, 2, 0, 0),
            new IrTokenOp(IrTokenOpKind.Equal, 2, 5, 0, 3));
        var mirrored = IrSplitSegmenter.MirrorDiff(diff);
        Assert.Equal(new IrTokenOp(IrTokenOpKind.Insert, 0, 0, 0, 2), mirrored.Ops[0]);
        Assert.Equal(new IrTokenOp(IrTokenOpKind.Equal, 0, 3, 2, 5), mirrored.Ops[1]);
        // Mirroring twice is the identity.
        Assert.Equal(diff, IrSplitSegmenter.MirrorDiff(mirrored));
    }

    // -------- detection (Task 4) --------

    private static IrBlockAlignment Align(IrDocument l, IrDocument r, IrDiffSettings s) =>
        IrBlockAligner.Align(l, r, s);

    [Fact]
    public void Detection_fires_for_a_clean_two_way_split()
    {
        // NB: this enters detection as state (b) — each half's Jaccard vs the whole (~0.56) clears the
        // 0.5 BlockSimilarityThreshold, so SimilarityPair pairs L with R0 first and the scan PROMOTES
        // the pairing. The genuinely-free state-(a) path is exercised by the three-way test below.
        var (l, _) = ReadParas("aaa bbb ccc ddd. eee fff ggg hhh.", "unrelated anchor paragraph one two three.");
        var (r, _) = ReadParas("aaa bbb ccc ddd. ", "eee fff ggg hhh.", "unrelated anchor paragraph one two three.");
        var a = Align(l, r, S);
        IrAlignmentAsserts.AssertInvariants(l, r, a, S);
        var split = a.Entries.Single(e => e.Kind == IrAlignmentKind.Split);
        Assert.Equal(2, split.MultiBlocks!.Count);
        Assert.Equal(0, IrAlignmentAsserts.Count(a, IrAlignmentKind.Inserted));
        Assert.Equal(0, IrAlignmentAsserts.Count(a, IrAlignmentKind.Deleted));
    }

    [Fact]
    public void Detection_fires_for_a_fully_free_three_way_split() // state (a): SimilarityPair declines
    {
        // Each third's Jaccard vs the whole is ~1/3 < 0.5, so SimilarityPair declines every pairing and
        // the candidate L reaches the scan genuinely FREE — the partner == -1 fire path.
        var (l, _) = ReadParas("aaa bbb ccc ddd. eee fff ggg hhh. iii jjj kkk lll.",
            "anchor one two three four five.");
        var (r, _) = ReadParas("aaa bbb ccc ddd. ", "eee fff ggg hhh. ", "iii jjj kkk lll.",
            "anchor one two three four five.");
        var a = Align(l, r, S);
        IrAlignmentAsserts.AssertInvariants(l, r, a, S);
        var split = a.Entries.Single(e => e.Kind == IrAlignmentKind.Split);
        Assert.Equal(3, split.MultiBlocks!.Count);
        Assert.Equal(0, IrAlignmentAsserts.Count(a, IrAlignmentKind.Inserted));
        Assert.Equal(0, IrAlignmentAsserts.Count(a, IrAlignmentKind.Deleted));
    }

    [Fact]
    public void Detection_absorbs_an_interior_net_new_block() // the WC-1830 math-paragraph shape
    {
        var (l, _) = ReadParas("aaa bbb ccc ddd eee fff. ggg hhh iii jjj kkk lll.");
        var (r, _) = ReadParas("aaa bbb ccc ddd eee fff. ", "zzz", "ggg hhh iii jjj kkk lll.");
        var a = Align(l, r, S);
        var split = a.Entries.Single(e => e.Kind == IrAlignmentKind.Split);
        Assert.Equal(3, split.MultiBlocks!.Count); // interior net-new member absorbed
        // The absorbed member must not ALSO surface as a plain Insert (consumed exactly once).
        Assert.Equal(0, IrAlignmentAsserts.Count(a, IrAlignmentKind.Inserted));
    }

    [Fact]
    public void Detection_promotes_a_similarity_paired_prefix_with_trailing_tail_inserts() // state (b)
    {
        // Prefix dominant enough that SimilarityPair pairs L with R0 (Jaccard > 0.5), tail falls out free.
        var (l, _) = ReadParas("aaa bbb ccc ddd eee fff ggg hhh iii jjj. kkk lll.");
        var (r, _) = ReadParas("aaa bbb ccc ddd eee fff ggg hhh iii jjj. ", "kkk lll.");
        var a = Align(l, r, S);
        Assert.Single(a.Entries.Where(e => e.Kind == IrAlignmentKind.Split));
        Assert.Equal(0, IrAlignmentAsserts.Count(a, IrAlignmentKind.Modified));
    }

    [Fact]
    public void Detection_does_not_fire_on_keyword_coincidence()
    {
        var (l, _) = ReadParas("the contract terminates on delivery of the goods.");
        var (r, _) = ReadParas("the parties agree on many things today.", "delivery of pizza is unrelated to goods.");
        var a = Align(l, r, S);
        Assert.Empty(a.Entries.Where(e => e.Kind is IrAlignmentKind.Split or IrAlignmentKind.Merge));
    }

    [Fact]
    public void Detection_excludes_an_unrelated_edge_insert_from_the_run() // R2 guard
    {
        var (l, _) = ReadParas("aaa bbb ccc ddd. eee fff ggg hhh.", "anchor one two three four five.");
        var (r, _) = ReadParas("aaa bbb ccc ddd. ", "eee fff ggg hhh.",
            "totally unrelated new paragraph words.", "anchor one two three four five.");
        var a = Align(l, r, S);
        var split = a.Entries.Single(e => e.Kind == IrAlignmentKind.Split);
        Assert.Equal(2, split.MultiBlocks!.Count); // edge net-new EXCLUDED → stays a plain Insert
        Assert.Equal(1, IrAlignmentAsserts.Count(a, IrAlignmentKind.Inserted));
    }

    [Fact]
    public void Detection_never_promotes_an_identity_reserved_unchanged_pair() // F4.2 / WC022 guard
    {
        // Content-equal pair (same text) + a following insert: an Unchanged pair has NO unmatched
        // tail (ContentHash-equal ⇒ all tokens matched), so the insert is genuinely new — no split.
        var (l, _) = ReadParas("same text here one two three.");
        var (r, _) = ReadParas("same text here one two three.", "a new paragraph appended after.");
        var a = Align(l, r, S);
        Assert.Empty(a.Entries.Where(e => e.Kind == IrAlignmentKind.Split));
        Assert.Equal(1, IrAlignmentAsserts.Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(1, IrAlignmentAsserts.Count(a, IrAlignmentKind.Inserted));
    }

    [Fact]
    public void Detection_merge_mirror_fires_for_a_clean_merge()
    {
        var (l, _) = ReadParas("aaa bbb ccc ddd. ", "eee fff ggg hhh.", "anchor one two three four.");
        var (r, _) = ReadParas("aaa bbb ccc ddd. eee fff ggg hhh.", "anchor one two three four.");
        var a = Align(l, r, S);
        var merge = a.Entries.Single(e => e.Kind == IrAlignmentKind.Merge);
        Assert.Equal(2, merge.MultiBlocks!.Count);
    }

    [Fact]
    public void Detection_two_adjacent_splits_never_share_a_right_block() // F2.2
    {
        var (l, _) = ReadParas("aaa bbb ccc. ddd eee fff.", "ggg hhh iii. jjj kkk lll.");
        var (r, _) = ReadParas("aaa bbb ccc. ", "ddd eee fff.", "ggg hhh iii. ", "jjj kkk lll.");
        var a = Align(l, r, S);
        var splits = a.Entries.Where(e => e.Kind == IrAlignmentKind.Split).ToList();
        Assert.Equal(2, splits.Count);
        var members = splits.SelectMany(e => e.MultiBlocks!).ToList();
        Assert.Equal(members.Count, members.Distinct(ReferenceEqualityComparer.Instance).Count());
    }

    [Fact]
    public void Detection_disabled_changes_nothing()
    {
        var (l, _) = ReadParas("aaa bbb ccc ddd. eee fff ggg hhh.");
        var (r, _) = ReadParas("aaa bbb ccc ddd. ", "eee fff ggg hhh.");
        var a = Align(l, r, new IrDiffSettings { DetectSplitMerge = false }); // the strict-1:1 opt-out
        Assert.Empty(a.Entries.Where(e => e.Kind is IrAlignmentKind.Split or IrAlignmentKind.Merge));
    }

    // -------- builder projection + apply-verifier (Task 5) --------

    private static (IrDocument L, IrDocument R, IrEditScript S) BuildScript(string[] left, string[] right)
    {
        var (l, _) = ReadParas(left);
        var (r, _) = ReadParas(right);
        var script = IrEditScriptBuilder.Build(l, r, S);
        return (l, r, script);
    }

    [Fact]
    public void Split_script_carries_one_SplitBlock_and_apply_verifies()
    {
        var (l, r, script) = BuildScript(
            new[] { "aaa bbb ccc ddd. eee fff ggg hhh.", "anchor one two three four five." },
            new[] { "aaa bbb ccc ddd. ", "eee fff ggg hhh.", "anchor one two three four five." });
        var split = Assert.Single(script.Operations.Where(o => o.Kind == IrEditOpKind.SplitBlock));
        Assert.Equal(2, split.SplitMergeAnchors!.Count);
        IrEditScriptVerifier.Verify(l, r, script, S); // count/order/ReferenceEquals proves apply (F3.1)
    }

    [Fact]
    public void Merge_script_carries_one_MergeBlock_and_apply_verifies()
    {
        var (l, r, script) = BuildScript(
            new[] { "aaa bbb ccc ddd. ", "eee fff ggg hhh.", "anchor one two three four five." },
            new[] { "aaa bbb ccc ddd. eee fff ggg hhh.", "anchor one two three four five." });
        Assert.Single(script.Operations.Where(o => o.Kind == IrEditOpKind.MergeBlock));
        IrEditScriptVerifier.Verify(l, r, script, S);
    }

    [Fact]
    public void Split_with_interior_insert_and_prefix_edit_apply_verifies()
    {
        var (l, r, script) = BuildScript(
            new[] { "aaa bbb ccc ddd eee fff. ggg hhh iii jjj kkk lll.", "anchor one two three." },
            new[] { "PRE aaa bbb ccc ddd eee fff. ", "zzz", "ggg hhh iii jjj kkk lll.", "anchor one two three." });
        Assert.Single(script.Operations.Where(o => o.Kind == IrEditOpKind.SplitBlock));
        IrEditScriptVerifier.Verify(l, r, script, S);
    }

    [Fact]
    public void Split_script_json_round_trips()
    {
        var (_, _, script) = BuildScript(
            new[] { "aaa bbb ccc ddd. eee fff ggg hhh." },
            new[] { "aaa bbb ccc ddd. ", "eee fff ggg hhh." });
        var json = IrEditScriptJson.Write(script);
        Assert.Equal(script, IrEditScriptJson.Read(json));
        Assert.Equal(json, IrEditScriptJson.Write(IrEditScriptJson.Read(json)));
    }

    // -------- revision renderer (Task 6) --------

    private static List<IrRevision> FixtureRevisions(string l, string r)
    {
        var settings = IrWmlComparerAdapter.MapSettings(new WmlComparerSettings()) with { DetectSplitMerge = true };
        var left = IrReader.Read(new WmlDocument(Path.Combine("../../../../TestFiles/", l)), WcCorpus.ReadOpts);
        var right = IrReader.Read(new WmlDocument(Path.Combine("../../../../TestFiles/", r)), WcCorpus.ReadOpts);
        var script = IrEditScriptBuilder.Build(left, right, settings);
        return IrRevisionRenderer.Render(script, left, right, settings).ToList();
    }

    [Fact]
    public void WC1830_compat_revisions_match_oracle_count()
    {
        // Oracle (WmlComparer): 2 — Deleted "When you click…add." + Inserted "\n" (the inserted mark).
        var revs = FixtureRevisions("WC/WC041-Table-5.docx", "WC/WC041-Table-5-Mod.docx");
        Assert.True(revs.Count == 2,
            $"expected 2 revisions, got {revs.Count}:\n" +
            string.Join("\n", revs.Select(rv => $"  {rv.Type}: [{rv.Text}]")));
    }

    [Fact]
    public void WC1450_compat_revisions_match_oracle_count()
    {
        var revs = FixtureRevisions("WC/WC023-Table-4-Row-Image-Before.docx",
            "WC/WC023-Table-4-Row-Image-After-Delete-1-Row.docx");
        Assert.True(revs.Count == 7,
            $"expected 7 revisions, got {revs.Count}:\n" +
            string.Join("\n", revs.Select(rv => $"  {rv.Type}: [{rv.Text}]")));
    }

    [Fact]
    public void Fine_mode_split_reports_per_segment_revisions_only()
    {
        var (l, r, script) = BuildScript(
            new[] { "aaa bbb ccc ddd. eee fff ggg hhh." },
            new[] { "aaa bbb ccc ddd. ", "NEW eee fff ggg hhh." });
        var revs = IrRevisionRenderer.Render(script, l, r, S).ToList(); // Fine granularity
        // Engine truth: the only content change is the inserted "NEW " inside segment 1.
        Assert.All(revs, rv => Assert.NotEqual(IrRevisionType.Deleted, rv.Type));
        Assert.Contains(revs, rv => rv.Type == IrRevisionType.Inserted && rv.Text.Contains("NEW"));
    }

    [Fact]
    public void Cell_scope_empty_mark_prune_fires() // F4.3 verification, pinned as a test
    {
        // A body-table cell paragraph anchors as p:body:… (IrReader assigns scope "body" throughout the
        // body, including table cells; only p:fn:/p:en: are excluded from the prune at
        // IrRevisionRenderer.IsZeroWidthBlock), so the compat empty-mark prune applies in cells.
        // RIGHT's cell gains one EMPTY paragraph: compat mode must report NO revision for it.
        const string cellL =
            "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>" +
            "<w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>" +
            "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>" +
            "<w:p><w:r><w:t>cell text here</w:t></w:r></w:p>" +
            "</w:tc></w:tr></w:tbl>";
        const string cellR =
            "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/></w:tblPr>" +
            "<w:tblGrid><w:gridCol w:w=\"2000\"/></w:tblGrid>" +
            "<w:tr><w:tc><w:tcPr><w:tcW w:w=\"2000\" w:type=\"dxa\"/></w:tcPr>" +
            "<w:p><w:r><w:t>cell text here</w:t></w:r></w:p><w:p/>" +
            "</w:tc></w:tr></w:tbl>";
        var ro = new IrReaderOptions { RetainSources = false, RevisionView = RevisionView.Accept };
        var l = IrReader.Read(IrTestDocuments.FromBodyXml(cellL), ro);
        var r = IrReader.Read(IrTestDocuments.FromBodyXml(cellR), ro);
        var compat = new IrDiffSettings
        {
            DetectSplitMerge = true,
            RevisionGranularity = RevisionGranularity.WmlComparerCompatible,
        };
        var script = IrEditScriptBuilder.Build(l, r, compat);
        var revs = IrRevisionRenderer.Render(script, l, r, compat);
        Assert.Empty(revs); // the empty-mark insert in a CELL is pruned (body-scope anchor)
    }
}
