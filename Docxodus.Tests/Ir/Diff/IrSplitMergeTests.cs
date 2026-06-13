#nullable enable

using System.Collections.Generic;
using System.Linq;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>M2.6 split/merge unit tests — op model + JSON wire (extended by later tasks: segmenter, detection, projection).</summary>
public class IrSplitMergeTests
{
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
}
