#nullable enable
using System.Linq;
using Docxodus;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

public class IrCompositeJsonTests
{
    [Fact]
    public void Consolidated_edit_script_json_has_author_and_conflicts()
    {
        var b = Docs.Para("the quick brown fox");
        var r1 = Docs.Para("the SLOW brown fox"); var r2 = Docs.Para("the FAST brown fox");
        var json = DocxDiff.GetConsolidatedEditScriptJson(b, new[]{
            new DocxDiffReviewer{Document=r1,Author="Bob"}, new DocxDiffReviewer{Document=r2,Author="Fred"}});
        Assert.Contains("\"author\"", json);
        Assert.Contains("\"conflicts\"", json);
        using var _ = System.Text.Json.JsonDocument.Parse(json); // valid JSON
    }

    [Fact]
    public void Consolidated_edit_script_json_carries_source_reviewer_and_anchors()
    {
        var b = Docs.Para("alpha one", "gamma three");
        var r1 = Docs.Para("alpha one EDITED", "gamma three");
        var r2 = Docs.Para("alpha one", "gamma three EDITED");
        var json = DocxDiff.GetConsolidatedEditScriptJson(b, new[]{
            new DocxDiffReviewer{Document=r1,Author="Bob"}, new DocxDiffReviewer{Document=r2,Author="Fred"}});
        Assert.Contains("\"sourceReviewer\"", json);
        Assert.Contains("\"operations\"", json);
        using var doc = System.Text.Json.JsonDocument.Parse(json);
        Assert.True(doc.RootElement.TryGetProperty("conflicts", out _));
        Assert.True(doc.RootElement.TryGetProperty("operations", out _));
    }

    [Fact]
    public void Zero_reviewers_json_is_empty_operations_and_conflicts()
    {
        var b = Docs.Para("alpha one");
        var json = DocxDiff.GetConsolidatedEditScriptJson(b, System.Array.Empty<DocxDiffReviewer>());
        using var doc = System.Text.Json.JsonDocument.Parse(json);
        Assert.Empty(doc.RootElement.GetProperty("operations").EnumerateArray());
        Assert.Empty(doc.RootElement.GetProperty("conflicts").EnumerateArray());
    }

    /// <summary>
    /// CONTRACT: any DeleteBlock/InsertBlock in the composite edit-script JSON must carry NULL move fields.
    /// The merger lowers COLLIDING reviewer moves to Insert/Delete and RETAINS the MoveGroupId/IsMoveSource
    /// marker INTERNALLY on a lowered move-source DeleteBlock (for contested-relocation detection); that
    /// marker is stripped before emission. The documented <c>IrEditOp</c> field-presence contract
    /// (IrEditScript.cs) says a DeleteBlock/InsertBlock carries NULL move fields, so the EMITTED, serialized
    /// op must NOT leak <c>moveGroupId</c>/<c>isMoveSource</c>. (Native — non-colliding — moves keep their
    /// MoveBlock kind, which DOES legitimately serialize move fields; that is covered separately below.)
    /// </summary>
    [Fact]
    public void Consolidated_json_delete_insert_ops_carry_no_move_fields()
    {
        // Two reviewers move the SAME ≥4-word paragraph to different places → a CONTESTED relocation: both
        // movers' sources lower to a marked DeleteBlock that co-anchors, so the move is NOT native and the
        // emitted delete/insert ops must be contract-clean (no leaked move fields). A 5-paragraph base with a
        // clear MIDDLE mover (P3) so the aligner anchors BOTH move sources at P3 (co-anchored → contested).
        const string p1 = "First paragraph alpha bravo charlie";
        const string p2 = "Second paragraph delta echo foxtrot";
        const string p3 = "Third paragraph golf hotel india";
        const string p4 = "Fourth paragraph juliet kilo lima";
        const string p5 = "Fifth paragraph mike november oscar";
        var b = Docs.Para(p1, p2, p3, p4, p5);
        var alice = Docs.Para(p1, p2, p4, p5, p3); // P3 → end
        var bob = Docs.Para(p3, p1, p2, p4, p5);   // P3 → front (contested with Alice)

        var json = DocxDiff.GetConsolidatedEditScriptJson(
            b, new[] { new DocxDiffReviewer { Document = alice, Author = "Alice" },
                       new DocxDiffReviewer { Document = bob, Author = "Bob" } });

        using var doc = System.Text.Json.JsonDocument.Parse(json);
        foreach (var op in doc.RootElement.GetProperty("operations").EnumerateArray())
        {
            var kind = op.GetProperty("kind").GetString();
            if (kind is "DeleteBlock" or "InsertBlock")
            {
                Assert.False(op.TryGetProperty("moveGroupId", out _),
                    $"{kind} leaked moveGroupId into the public edit-script JSON (IrEditOp contract violation).");
                Assert.False(op.TryGetProperty("isMoveSource", out _),
                    $"{kind} leaked isMoveSource into the public edit-script JSON (IrEditOp contract violation).");
            }
        }

        // Sanity: the contested move WAS lowered (a DeleteBlock for the move source is present), so the
        // assertion above is not vacuously true.
        Assert.Contains(doc.RootElement.GetProperty("operations").EnumerateArray(),
            op => op.GetProperty("kind").GetString() == "DeleteBlock");
    }

    /// <summary>
    /// FOLLOW-ON A: a single-reviewer non-colliding move serializes as a NATIVE MoveBlock pair — both halves
    /// carry the (global) <c>moveGroupId</c> and <c>isMoveSource</c>.
    /// </summary>
    [Fact]
    public void Consolidated_json_native_move_serializes_moveGroupId_and_isMoveSource_for_both_halves()
    {
        const string p1 = "First paragraph alpha bravo charlie";
        const string p2 = "Second paragraph delta echo foxtrot";
        const string p3 = "Third paragraph golf hotel india";
        const string p4 = "Fourth paragraph juliet kilo lima";
        var b = Docs.Para(p1, p2, p3, p4);
        var alice = Docs.Para(p1, p3, p4, p2);   // P2 relocated to the end (native move)

        var json = DocxDiff.GetConsolidatedEditScriptJson(
            b, new[] { new DocxDiffReviewer { Document = alice, Author = "Alice" } });

        using var doc = System.Text.Json.JsonDocument.Parse(json);
        var moveOps = doc.RootElement.GetProperty("operations").EnumerateArray()
            .Where(op => op.GetProperty("kind").GetString() is "MoveBlock" or "MoveModifyBlock")
            .ToList();
        Assert.Equal(2, moveOps.Count); // source + destination

        foreach (var op in moveOps)
        {
            Assert.True(op.TryGetProperty("moveGroupId", out _),
                "native composite move op must serialize moveGroupId.");
            Assert.True(op.TryGetProperty("isMoveSource", out _),
                "native composite move op must serialize isMoveSource.");
        }

        // The two halves share one group id and split source/dest.
        var gids = moveOps.Select(op => op.GetProperty("moveGroupId").GetInt32()).Distinct().ToList();
        Assert.Single(gids);
        Assert.Contains(moveOps, op => op.GetProperty("isMoveSource").GetBoolean());
        Assert.Contains(moveOps, op => !op.GetProperty("isMoveSource").GetBoolean());
    }
}
