#nullable enable

using System.Collections.Generic;
using System.Linq;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.2 Task 2 tests for <see cref="IrEditScriptBuilder"/>, <see cref="IrEditScriptVerifier"/>, and
/// <see cref="IrEditScriptJson"/>: per-kind alignment→op mapping, move source/destination pairing +
/// group-id determinism, apply-verifier green on every synthetic alignment case, and JSON round-trip.
/// </summary>
/// <remarks>
/// Documents are built via <see cref="IrTestDocuments"/> + <see cref="IrReader"/> (RetainSources=false).
/// The synthetic cases mirror the M2.1 aligner fixtures (identity, edits, insert/delete, pure move,
/// swap, format-only, boilerplate) so the edit script is exercised over the same alignment shapes the
/// aligner tests pin.
/// </remarks>
public class IrEditScriptTests
{
    private static readonly IrReaderOptions NoSources = new() { RetainSources = false };
    private static readonly IrDiffSettings Default = new();

    private static IrDocument Doc(params string[] paragraphTexts) =>
        IrReader.Read(IrTestDocuments.Create(paragraphTexts), NoSources);

    private static IrDocument FromXml(string bodyInnerXml) =>
        IrReader.Read(IrTestDocuments.FromBodyXml(bodyInnerXml), NoSources);

    private static IrEditScript Build(IrDocument l, IrDocument r) =>
        IrEditScriptBuilder.Build(l, r, Default);

    private static int Count(IrEditScript s, IrEditOpKind k) => s.Operations.Count(o => o.Kind == k);

    /// <summary>Build + apply-verify + JSON-round-trip a case in one shot (the full Task 2 triple).</summary>
    private static IrEditScript BuildVerified(IrDocument l, IrDocument r)
    {
        var script = Build(l, r);
        IrEditScriptVerifier.Verify(l, r, script);
        AssertJsonRoundTrips(script);
        return script;
    }

    private static void AssertJsonRoundTrips(IrEditScript script)
    {
        var json = IrEditScriptJson.Write(script);
        var back = IrEditScriptJson.Read(json);
        Assert.Equal(script, back);
        // Write is deterministic: a second write of the round-tripped script is byte-identical.
        Assert.Equal(json, IrEditScriptJson.Write(back));
    }

    // ------------------------------------------------------------------ per-kind mapping

    [Fact]
    public void Identity_all_equal_blocks()
    {
        var l = Doc("alpha", "beta", "gamma");
        var r = Doc("alpha", "beta", "gamma");
        var s = BuildVerified(l, r);

        Assert.Equal(3, s.Operations.Count);
        Assert.All(s.Operations, o => Assert.Equal(IrEditOpKind.EqualBlock, o.Kind));
        Assert.All(s.Operations, o =>
        {
            Assert.NotNull(o.LeftAnchor);
            Assert.NotNull(o.RightAnchor);
            Assert.Null(o.TokenDiff);
        });
    }

    [Fact]
    public void Single_edit_is_modify_block_with_token_diff()
    {
        var l = Doc("alpha", "beta", "gamma");
        var r = Doc("alpha", "beta edited here", "gamma");
        var s = BuildVerified(l, r);

        Assert.Equal(2, Count(s, IrEditOpKind.EqualBlock));
        Assert.Equal(1, Count(s, IrEditOpKind.ModifyBlock));
        var modify = s.Operations.Single(o => o.Kind == IrEditOpKind.ModifyBlock);
        Assert.NotNull(modify.TokenDiff);
        Assert.NotNull(modify.LeftAnchor);
        Assert.NotNull(modify.RightAnchor);
        // The token diff covers the paragraph and includes at least one Insert (added words).
        Assert.Contains(modify.TokenDiff!.Ops, o => o.Kind == IrTokenOpKind.Insert);
    }

    [Fact]
    public void Interior_full_rewrite_carries_matching_body_group_through_json()
    {
        // The preceding paragraph pairs as a normal edit and the following paragraph is equal, so
        // the zero-lexical middle residue is the body shape Word renders as two physical paragraphs.
        var l = Doc("Anchor title blue", "obsolete amber stanza", "shared trailing paragraph");
        var r = Doc("Anchor title bold", "fresh quantum clause", "shared trailing paragraph");
        var s = BuildVerified(l, r);

        var rewriteOps = s.Operations
            .Where(o => o.Kind is IrEditOpKind.InsertBlock or IrEditOpKind.DeleteBlock).ToList();
        Assert.Equal(2, rewriteOps.Count);
        Assert.NotNull(rewriteOps[0].BodyFullRewriteGroupId);
        Assert.Equal(rewriteOps[0].BodyFullRewriteGroupId, rewriteOps[1].BodyFullRewriteGroupId);
        Assert.Contains("\"bodyFullRewriteGroupId\"", IrEditScriptJson.Write(s));
    }

    [Fact]
    public void Insert_block_has_right_anchor_only()
    {
        var l = Doc("alpha", "beta");
        var r = Doc("alpha", "NEW", "beta");
        var s = BuildVerified(l, r);

        Assert.Equal(1, Count(s, IrEditOpKind.InsertBlock));
        var insert = s.Operations.Single(o => o.Kind == IrEditOpKind.InsertBlock);
        Assert.Null(insert.LeftAnchor);
        Assert.NotNull(insert.RightAnchor);
        Assert.Equal(IrEditOpKind.InsertBlock, s.Operations[1].Kind); // emitted at the right position
    }

    [Fact]
    public void Delete_block_has_left_anchor_only()
    {
        var l = Doc("alpha", "beta", "gamma");
        var r = Doc("alpha", "gamma");
        var s = BuildVerified(l, r);

        Assert.Equal(1, Count(s, IrEditOpKind.DeleteBlock));
        var del = s.Operations.Single(o => o.Kind == IrEditOpKind.DeleteBlock);
        Assert.NotNull(del.LeftAnchor);
        Assert.Null(del.RightAnchor);
        // Left-anchored interleave: deletion of "beta" trails "alpha"'s equal op, before "gamma".
        Assert.Equal(IrEditOpKind.EqualBlock, s.Operations[0].Kind);
        Assert.Equal(IrEditOpKind.DeleteBlock, s.Operations[1].Kind);
        Assert.Equal(IrEditOpKind.EqualBlock, s.Operations[2].Kind);
    }

    [Fact]
    public void Format_only_block()
    {
        var l = FromXml(
            "<w:p><w:r><w:t>alpha</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>beta</w:t></w:r></w:p>");
        var r = FromXml(
            "<w:p><w:r><w:t>alpha</w:t></w:r></w:p>" +
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>beta</w:t></w:r></w:p>");
        var s = BuildVerified(l, r);

        Assert.Equal(1, Count(s, IrEditOpKind.FormatOnlyBlock));
        Assert.Equal(1, Count(s, IrEditOpKind.EqualBlock));
        var fmt = s.Operations.Single(o => o.Kind == IrEditOpKind.FormatOnlyBlock);
        Assert.NotNull(fmt.LeftAnchor);
        Assert.NotNull(fmt.RightAnchor);
        Assert.Null(fmt.TokenDiff);
    }

    // ------------------------------------------------------------------ move source/destination pairing

    [Fact]
    public void Pure_move_emits_source_and_destination_sharing_a_group()
    {
        // "gamma" relocated from the end to the front.
        var l = Doc("alpha", "beta", "gamma", "delta");
        var r = Doc("gamma", "alpha", "beta", "delta");
        var s = BuildVerified(l, r);

        var moves = s.Operations.Where(o => o.Kind == IrEditOpKind.MoveBlock).ToList();
        Assert.Equal(2, moves.Count); // exactly source + destination

        var source = moves.Single(o => o.IsMoveSource == true);
        var dest = moves.Single(o => o.IsMoveSource == false);
        Assert.Equal(source.MoveGroupId, dest.MoveGroupId);
        Assert.Equal(1, source.MoveGroupId); // first (only) group → id 1

        Assert.NotNull(source.LeftAnchor);
        Assert.Null(source.RightAnchor);
        Assert.Null(dest.LeftAnchor);
        Assert.NotNull(dest.RightAnchor);
        Assert.Null(dest.TokenDiff); // plain move: exact content, no diff

        // Destination is emitted at the front (gamma's new right position); the rest are equal.
        Assert.Equal(IrEditOpKind.MoveBlock, s.Operations[0].Kind);
        Assert.False(s.Operations[0].IsMoveSource);
    }

    [Fact]
    public void Move_source_interleaves_at_left_anchored_position()
    {
        // "gamma" moves to the front. On the LEFT, gamma sits after beta; its SOURCE op must trail
        // beta's equal op (left-anchored convention), i.e. after the alpha,beta equal ops.
        var l = Doc("alpha", "beta", "gamma", "delta");
        var r = Doc("gamma", "alpha", "beta", "delta");
        var s = BuildVerified(l, r);

        // Sequence: [dest gamma], [equal alpha], [equal beta], [source gamma], [equal delta].
        var kinds = s.Operations.Select(o => (o.Kind, o.IsMoveSource)).ToList();
        int sourceIdx = kinds.FindIndex(t => t.Kind == IrEditOpKind.MoveBlock && t.IsMoveSource == true);
        int betaIdx = s.Operations.ToList().FindIndex(o =>
            o.Kind == IrEditOpKind.EqualBlock && o.LeftAnchor == l.Body.Blocks[1].Anchor.ToString());
        int deltaIdx = s.Operations.ToList().FindIndex(o =>
            o.Kind == IrEditOpKind.EqualBlock && o.LeftAnchor == l.Body.Blocks[3].Anchor.ToString());
        Assert.True(sourceIdx > betaIdx, "move source must trail the preceding paired-left (beta).");
        Assert.True(sourceIdx < deltaIdx, "move source must precede the next paired-left (delta).");
    }

    [Fact]
    public void Two_moves_get_group_ids_in_destination_order()
    {
        // Two distinct blocks relocate to the front (reversed order at destination): the group ids are
        // assigned 1,2 in destination (right) order, deterministically.
        var l = Doc("alpha", "beta", "gamma", "delta", "epsilon");
        var r = Doc("delta", "gamma", "alpha", "beta", "epsilon");
        var s = BuildVerified(l, r);

        var destinations = s.Operations
            .Where(o => o.Kind == IrEditOpKind.MoveBlock && o.IsMoveSource == false)
            .ToList();
        Assert.True(destinations.Count >= 1);
        // Group ids appear 1,2,… in destination order.
        var ids = destinations.Select(o => o.MoveGroupId!.Value).ToList();
        Assert.Equal(Enumerable.Range(1, ids.Count).ToList(), ids);
    }

    [Fact]
    public void Adjacent_swap_one_move()
    {
        var l = Doc("alpha", "beta", "gamma");
        var r = Doc("beta", "alpha", "gamma");
        var s = BuildVerified(l, r);

        var moves = s.Operations.Where(o => o.Kind == IrEditOpKind.MoveBlock).ToList();
        Assert.Equal(2, moves.Count); // one move = source + destination
        Assert.Equal(1, moves.Single(o => o.IsMoveSource == true).MoveGroupId);
    }

    // ------------------------------------------------------------------ M2.2 Task 3: MoveModifyBlock (headline)

    [Fact]
    public void Moved_and_edited_paragraph_is_move_modify_block_with_nested_token_ops()
    {
        // THE headline M2.2 capability: a multi-word paragraph relocates from the tail to the front AND is
        // edited in the same revision. The aligner classifies it MovedModified; the edit script emits a
        // MoveModifyBlock source + destination, the DESTINATION carrying the in-move token diff (source
        // tokens vs destination tokens). Apply-verified (the destination reconstructs from the SOURCE
        // block's tokens via the token diff) and JSON-round-tripped.
        var l = Doc(
            "alpha", "beta", "gamma", "delta",
            "the quick brown fox jumps over hounds");
        var r = Doc(
            "the quick brown fox jumps over dogs",
            "alpha", "beta", "gamma", "delta");
        var s = BuildVerified(l, r); // apply-verifier + JSON round-trip baked in

        var moves = s.Operations.Where(o => o.Kind == IrEditOpKind.MoveModifyBlock).ToList();
        Assert.Equal(2, moves.Count); // source + destination

        var source = moves.Single(o => o.IsMoveSource == true);
        var dest = moves.Single(o => o.IsMoveSource == false);
        Assert.Equal(source.MoveGroupId, dest.MoveGroupId);

        // Source carries the left anchor and no diff; destination carries the right anchor and the diff.
        Assert.NotNull(source.LeftAnchor);
        Assert.Null(source.RightAnchor);
        Assert.Null(source.TokenDiff);
        Assert.Null(dest.LeftAnchor);
        Assert.NotNull(dest.RightAnchor);
        Assert.NotNull(dest.TokenDiff);

        // The nested token diff is a real edit: an Equal prefix (the shared words) plus a Delete and an
        // Insert (hounds → dogs).
        var ops = dest.TokenDiff!.Ops;
        Assert.Contains(ops, o => o.Kind == IrTokenOpKind.Equal);
        Assert.Contains(ops, o => o.Kind == IrTokenOpKind.Delete);
        Assert.Contains(ops, o => o.Kind == IrTokenOpKind.Insert);

        // The destination op is at the front (the moved block's new right position).
        Assert.Equal(IrEditOpKind.MoveModifyBlock, s.Operations[0].Kind);
        Assert.False(s.Operations[0].IsMoveSource);

        // Explicit JSON round-trip of the nested token diff (covers the compact-array encoding under a move).
        var back = IrEditScriptJson.Read(IrEditScriptJson.Write(s));
        Assert.Equal(s, back);
        var destBack = back.Operations.Single(o => o.Kind == IrEditOpKind.MoveModifyBlock && o.IsMoveSource == false);
        Assert.Equal(dest.TokenDiff, destBack.TokenDiff);
    }

    // ------------------------------------------------------------------ boilerplate

    [Fact]
    public void Boilerplate_delete_one_no_false_moves()
    {
        var l = Doc(Enumerable.Repeat("boilerplate", 10).ToArray());
        var r = Doc(Enumerable.Repeat("boilerplate", 9).ToArray());
        var s = BuildVerified(l, r);

        Assert.Equal(9, Count(s, IrEditOpKind.EqualBlock));
        Assert.Equal(1, Count(s, IrEditOpKind.DeleteBlock));
        Assert.Equal(0, Count(s, IrEditOpKind.MoveBlock));
    }

    // ------------------------------------------------------------------ non-paragraph Modified

    [Fact]
    public void Modified_table_has_null_token_diff()
    {
        const string tbl =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"100\"/></w:tblGrid>" +
            "<w:tr><w:tc><w:p><w:r><w:t>{0}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        var l = FromXml("<w:p><w:r><w:t>intro</w:t></w:r></w:p>" + string.Format(tbl, "cell-old"));
        var r = FromXml("<w:p><w:r><w:t>intro</w:t></w:r></w:p>" + string.Format(tbl, "cell-new"));
        var s = BuildVerified(l, r);

        var modify = s.Operations.Single(o => o.Kind == IrEditOpKind.ModifyBlock);
        Assert.Null(modify.TokenDiff); // table granularity is Task 4
    }

    // ------------------------------------------------------------------ empty docs

    [Fact]
    public void Empty_left_all_inserts()
    {
        var l = FromXml(string.Empty);
        var r = Doc("alpha", "beta");
        var s = BuildVerified(l, r);
        Assert.Equal(2, Count(s, IrEditOpKind.InsertBlock));
    }

    [Fact]
    public void Empty_right_all_deletes()
    {
        var l = Doc("alpha", "beta");
        var r = FromXml(string.Empty);
        var s = BuildVerified(l, r);
        Assert.Equal(2, Count(s, IrEditOpKind.DeleteBlock));
    }

    [Fact]
    public void Both_empty_no_ops()
    {
        var l = FromXml(string.Empty);
        var r = FromXml(string.Empty);
        var s = BuildVerified(l, r);
        Assert.Empty(s.Operations);
    }

    // ------------------------------------------------------------------ determinism

    [Fact]
    public void Two_builds_are_record_equal()
    {
        var l = Doc("alpha", "beta", "gamma", "delta", "boilerplate", "boilerplate");
        var r = Doc("gamma", "alpha", "beta edited", "boilerplate", "delta", "NEW");

        var a = Build(l, r);
        var b = Build(l, r);
        Assert.Equal(a, b); // record equality over the whole script
        IrEditScriptVerifier.Verify(l, r, a);
        AssertJsonRoundTrips(a);
    }

    // ------------------------------------------------------------------ JSON round-trip (every kind reachable)

    [Fact]
    public void Json_round_trips_a_script_with_every_reachable_kind()
    {
        // A composite case producing Equal, FormatOnly, Modify, Insert, Delete, and Move (src+dest).
        // Each non-equal block is isolated in its OWN spine gap by stable "anchorX"/"anchorY"/"tail"
        // anchors so the M2.2 similarity pairing classifies each unambiguously: the modify sits alone in
        // a 1×1 gap (the unambiguous-residue fallback pairs it as Modified even though "beta" → "beta
        // edited" scores below the 0.5 threshold), the delete sits alone (1 free left, 0 free right →
        // Deleted), and the insert sits alone (0 free left, 1 free right → Inserted).
        var l = FromXml(
            "<w:p><w:r><w:t>alpha</w:t></w:r></w:p>" +          // stays (equal)
            "<w:p><w:r><w:t>fmt</w:t></w:r></w:p>" +            // bolded → format-only
            "<w:p><w:r><w:t>beta</w:t></w:r></w:p>" +          // edited → modify (alone in a 1×1 gap)
            "<w:p><w:r><w:t>anchorX</w:t></w:r></w:p>" +       // stable anchor (equal)
            "<w:p><w:r><w:t>todelete</w:t></w:r></w:p>" +      // deleted (alone in its gap)
            "<w:p><w:r><w:t>anchorY</w:t></w:r></w:p>" +       // stable anchor (equal)
            "<w:p><w:r><w:t>mover</w:t></w:r></w:p>" +         // moved to front
            "<w:p><w:r><w:t>tail</w:t></w:r></w:p>");          // stays (equal)
        var r = FromXml(
            "<w:p><w:r><w:t>mover</w:t></w:r></w:p>" +         // moved here (destination)
            "<w:p><w:r><w:t>alpha</w:t></w:r></w:p>" +
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>fmt</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>beta edited</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>anchorX</w:t></w:r></w:p>" +       // stable anchor (equal)
            "<w:p><w:r><w:t>anchorY</w:t></w:r></w:p>" +       // stable anchor (equal)
            "<w:p><w:r><w:t>inserted</w:t></w:r></w:p>" +      // inserted (own gap, no free left)
            "<w:p><w:r><w:t>tail</w:t></w:r></w:p>");
        var s = BuildVerified(l, r);

        // Confirm the script genuinely exercises the full set we can reach today.
        var kinds = s.Operations.Select(o => o.Kind).ToHashSet();
        Assert.Contains(IrEditOpKind.EqualBlock, kinds);
        Assert.Contains(IrEditOpKind.FormatOnlyBlock, kinds);
        Assert.Contains(IrEditOpKind.ModifyBlock, kinds);
        Assert.Contains(IrEditOpKind.InsertBlock, kinds);
        Assert.Contains(IrEditOpKind.DeleteBlock, kinds);
        Assert.Contains(IrEditOpKind.MoveBlock, kinds);

        // Round-trip equality already asserted by BuildVerified; re-assert explicitly for clarity.
        var back = IrEditScriptJson.Read(IrEditScriptJson.Write(s));
        Assert.Equal(s, back);
    }

    [Fact]
    public void Json_round_trips_a_modify_token_diff_with_all_token_op_kinds()
    {
        // Edit a paragraph so its token diff has Equal, Delete, Insert, and FormatChanged ops, then
        // round-trip the whole script and assert the token diff survives byte-for-byte.
        var l = FromXml(
            "<w:p><w:r><w:t>the </w:t></w:r><w:r><w:t>quick brown fox</w:t></w:r></w:p>");
        var r = FromXml(
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>the </w:t></w:r>" +
            "<w:r><w:t>slow brown wolf</w:t></w:r></w:p>");
        var s = BuildVerified(l, r);

        var modify = s.Operations.Single(o => o.Kind == IrEditOpKind.ModifyBlock);
        Assert.NotNull(modify.TokenDiff);
        var back = IrEditScriptJson.Read(IrEditScriptJson.Write(s));
        Assert.Equal(s, back);
        // The round-tripped token diff is record-equal (covers the compact-array encoding).
        var modifyBack = back.Operations.Single(o => o.Kind == IrEditOpKind.ModifyBlock);
        Assert.Equal(modify.TokenDiff, modifyBack.TokenDiff);
    }

    /// <summary>
    /// Prelim (b): explicit JSON round-trip of a HAND-BUILT script carrying <see cref="IrTextboxDiff"/>s — a
    /// ModifyBlock whose paragraph holds two textboxes (one edited via a nested token diff, one whose surplus
    /// block is all-Deleted). Record equality after Write→Read proves the textboxDiffs branch of
    /// <see cref="IrEditScriptJson"/> serializes and parses losslessly, and a second Write is byte-identical
    /// (determinism). Hand-built rather than diff-derived so the textbox-diff shape is pinned independently of
    /// how the builder happens to produce it.
    /// </summary>
    [Fact]
    public void Json_round_trips_a_hand_built_script_with_textbox_diffs()
    {
        // Textbox 1: an interior edit — a token diff with Equal + Delete + Insert spans.
        var editedBox = new IrTextboxDiff(IrNodeList.From(new[]
        {
            new IrEditOp(IrEditOpKind.ModifyBlock, "p:tbx:box1para", "p:tbx:box1para2",
                new IrTokenDiff(IrNodeList.From(new[]
                {
                    new IrTokenOp(IrTokenOpKind.Equal, 0, 1, 0, 1),
                    new IrTokenOp(IrTokenOpKind.Delete, 1, 2, 1, 1),
                    new IrTokenOp(IrTokenOpKind.Insert, 2, 2, 1, 2),
                })),
                null, null),
        }));

        // Textbox 2: a removed textbox — its lone inner block is all-Deleted.
        var removedBox = new IrTextboxDiff(IrNodeList.From(new[]
        {
            new IrEditOp(IrEditOpKind.DeleteBlock, "p:tbx:box2para", null, null, null, null),
        }));

        var op = new IrEditOp(
            IrEditOpKind.ModifyBlock, "p:body:host", "p:body:host2",
            new IrTokenDiff(IrNodeList.From(new[] { new IrTokenOp(IrTokenOpKind.Equal, 0, 1, 0, 1) })),
            null, null, null,
            IrNodeList.From(new[] { editedBox, removedBox }));

        var script = new IrEditScript(IrNodeList.From(new[] { op }));

        var json = IrEditScriptJson.Write(script);
        var back = IrEditScriptJson.Read(json);

        Assert.Equal(script, back); // full record equality, including the nested textboxDiffs
        Assert.Equal(json, IrEditScriptJson.Write(back)); // deterministic re-write

        // Spot-check the round-tripped textbox diffs survived structurally.
        var backOp = Assert.Single(back.Operations);
        Assert.NotNull(backOp.TextboxDiffs);
        Assert.Equal(2, backOp.TextboxDiffs!.Count);
        Assert.Equal(IrEditOpKind.ModifyBlock, backOp.TextboxDiffs[0].Ops.Single().Kind);
        Assert.Equal(IrEditOpKind.DeleteBlock, backOp.TextboxDiffs[1].Ops.Single().Kind);
    }
}
