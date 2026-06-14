#nullable enable

using System.Collections.Generic;
using System.Linq;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// N-way merger that consolidates several pairwise <see cref="IrEditScript"/>s — each computed
/// against the SAME base document, so they share one base anchor space — into a single
/// <see cref="IrCompositeScript"/>.
/// <para>The merge walks the base document in block order. For each base block it gathers the
/// reviewers' ops anchored there (via <see cref="GroupByBaseAnchor"/>) and dispatches to
/// <c>MergeOneBaseBlock</c>; right-only inserts are slotted relative to the preceding base anchor
/// (via <c>GroupInsertsByPrecedingAnchor</c> + <c>EmitInsertsAt</c>). Conflicts — a base block
/// edited differently by 2+ reviewers — are resolved by the supplied
/// <see cref="Docxodus.ConflictResolution"/> policy and recorded in
/// <see cref="IrCompositeScript.Conflicts"/>.</para>
/// <para>See <c>docs/architecture/ir_diff_engine.md</c> for the diff pipeline this composes over.</para>
/// </summary>
internal static class IrCompositeMerger
{
    public static IrCompositeScript Merge(
        IrDocument baseIr,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        Docxodus.ConflictResolution policy,
        IrDiffSettings settings)
    {
        // 1. Raw pairwise scripts, NOT yet lowered — so PlanMoves can inspect every reviewer's move groups
        //    against the shared base anchor space before any move is collapsed to del/ins.
        var rawScripts = reviewers
            .Select(r => IrEditScriptBuilder.Build(baseIr, r.Ir, settings))
            .ToList();

        // 2. Decide, per reviewer move group, NATIVE (non-colliding single-reviewer move → keep
        //    MoveBlock/MoveModifyBlock) vs LOWER (colliding move → del/ins), and assign a GLOBAL move-group
        //    id so two reviewers' independent native moves never share a w:name.
        var plan = PlanMoves(rawScripts, settings);

        // 3. Transform each script per the plan: NATIVE move groups keep their move ops (with the global gid
        //    substituted); LOWERED move groups + ALL Split/Merge collapse to Insert/Delete (LowerStructuralOps).
        var scripts = rawScripts
            .Select((s, reviewer) => ApplyMovePlan(s, reviewer, plan))
            .ToList();
        var baseOrder = BaseBlockAnchors(baseIr);
        var byBase = GroupByBaseAnchor(scripts);
        var insertsAfter = GroupInsertsByPrecedingAnchor(scripts);

        var ops = new List<IrCompositeOp>();
        var conflicts = new List<IrConflict>();
        var nextConflictId = 1;
        EmitInsertsAt(insertsAfter, "", reviewers, ops);
        foreach (var anchor in baseOrder)
        {
            MergeOneBaseBlock(anchor, byBase, reviewers, baseIr, policy, settings, ops, conflicts, ref nextConflictId);
            EmitInsertsAt(insertsAfter, anchor, reviewers, ops);
        }
        return new IrCompositeScript(IrNodeList.From(ops), IrNodeList.From(conflicts));
    }

    // ---- FOLLOW-ON A: native move composition planning ----

    /// <summary>
    /// The verdict for every reviewer move group: whether it renders as a NATIVE move
    /// (<c>w:moveFrom</c>/<c>w:moveTo</c>, authored to the mover) or is LOWERED to del/ins, plus the GLOBAL
    /// move-group id assigned to a native group. Keyed by <c>(reviewer index, the reviewer's LOCAL
    /// <see cref="IrEditOp.MoveGroupId"/>)</c> — both halves of a move share the local gid, so both resolve to
    /// the same verdict and global gid. <see cref="GlobalGidByGroup"/> contains an entry ONLY for native groups
    /// (a group absent from the map is lowered).
    /// </summary>
    internal sealed record MovePlan(IReadOnlyDictionary<(int Reviewer, int LocalGid), int> GlobalGidByGroup);

    /// <summary>
    /// Decide, per reviewer move group, NATIVE vs LOWER, and assign deterministic GLOBAL move-group ids.
    /// <para>A move group (reviewer R, local gid G, source base anchor S) is NATIVE-eligible iff
    /// <see cref="IrDiffSettings.RenderMoves"/> AND R is the ONLY reviewer that touches base block S — i.e.
    /// <c>touchersByBaseAnchor[S] == {R}</c>. This single predicate covers BOTH collision shapes the lowering
    /// must keep: a move-vs-edit on S (the other reviewer's ModifyBlock/DeleteBlock co-anchors at S) AND two
    /// reviewers moving the same block (both move sources co-anchor at S). A move whose source block is
    /// touched only by its mover is non-colliding and renders natively.</para>
    /// <para><b>Global namespacing.</b> Native groups are assigned global ids from a single counter in
    /// deterministic order — reviewers in list order, then each reviewer's groups by ascending local gid —
    /// starting at 1. Both halves (matched by local gid within the reviewer) get the same global id, so the
    /// renderer's <c>w:name</c> allocator pairs them while two reviewers' independent moves stay distinct.</para>
    /// </summary>
    internal static MovePlan PlanMoves(IReadOnlyList<IrEditScript> rawScripts, IrDiffSettings settings)
    {
        // touchersByBaseAnchor: base anchor → reviewers with a non-Equal op anchored (LeftAnchor) there.
        // A move SOURCE (IsMoveSource=true) anchors at its source base block via LeftAnchor; a ModifyBlock/
        // DeleteBlock/FormatOnlyBlock also carries the base block's LeftAnchor; a move DEST / InsertBlock has
        // a null LeftAnchor and so never marks a base block as touched (it is right-positioned content).
        var touchers = new Dictionary<string, HashSet<int>>();
        for (int reviewer = 0; reviewer < rawScripts.Count; reviewer++)
        {
            foreach (var op in rawScripts[reviewer].Operations)
            {
                if (op.Kind == IrEditOpKind.EqualBlock || op.LeftAnchor is not { } anchor)
                    continue;
                if (!touchers.TryGetValue(anchor, out var set))
                    touchers[anchor] = set = new HashSet<int>();
                set.Add(reviewer);
            }
        }

        var globalGid = new Dictionary<(int, int), int>();
        int nextGlobal = 1;
        if (settings.RenderMoves)
        {
            for (int reviewer = 0; reviewer < rawScripts.Count; reviewer++)
            {
                // The reviewer's native-eligible groups, by ascending local gid, each mapped to its source
                // base anchor (the move-SOURCE op's LeftAnchor). Only groups with a resolvable source whose
                // base block is touched solely by this reviewer are native.
                var eligibleLocalGids = new SortedSet<int>();
                foreach (var op in rawScripts[reviewer].Operations)
                {
                    if (op.Kind is not (IrEditOpKind.MoveBlock or IrEditOpKind.MoveModifyBlock)
                        || op.IsMoveSource != true || op.MoveGroupId is not { } localGid
                        || op.LeftAnchor is not { } source)
                        continue;
                    bool soleToucher = touchers.TryGetValue(source, out var set)
                        && set.Count == 1 && set.Contains(reviewer);
                    if (soleToucher)
                        eligibleLocalGids.Add(localGid);
                }
                foreach (var localGid in eligibleLocalGids)
                    globalGid[(reviewer, localGid)] = nextGlobal++;
            }
        }

        return new MovePlan(globalGid);
    }

    /// <summary>
    /// Transform reviewer <paramref name="reviewer"/>'s raw pairwise script per <paramref name="plan"/>:
    /// a NATIVE move group keeps its <see cref="IrEditOpKind.MoveBlock"/>/<see cref="IrEditOpKind.MoveModifyBlock"/>
    /// source AND destination, with <see cref="IrEditOp.MoveGroupId"/> rewritten to the GLOBAL gid; a LOWERED
    /// move group and ALL Split/Merge collapse to Insert/Delete via the per-op lowering rules. Op order is
    /// preserved (the preceding-anchor insert routing depends on it).
    /// </summary>
    internal static IrEditScript ApplyMovePlan(IrEditScript script, int reviewer, MovePlan plan)
    {
        var result = new List<IrEditOp>(script.Operations.Count);
        foreach (var op in script.Operations)
        {
            if (op.Kind is IrEditOpKind.MoveBlock or IrEditOpKind.MoveModifyBlock
                && op.MoveGroupId is { } localGid
                && plan.GlobalGidByGroup.TryGetValue((reviewer, localGid), out var global))
            {
                // NATIVE: keep both halves verbatim, rewriting the local gid to the globally-namespaced one.
                result.Add(op with { MoveGroupId = global });
            }
            else
            {
                // LOWERED move group, or a Split/Merge → del/ins (LowerStructuralOps' per-op rules).
                LowerOneStructuralOp(op, result);
            }
        }

        System.Diagnostics.Debug.Assert(
            result.All(o =>
                o.Kind is not (IrEditOpKind.SplitBlock or IrEditOpKind.MergeBlock)
                && (o.Kind is not (IrEditOpKind.MoveBlock or IrEditOpKind.MoveModifyBlock)
                    || (o.MoveGroupId is { } g && plan.GlobalGidByGroup.Values.Contains(g)))),
            "ApplyMovePlan must leave no Split/Merge and no LOWERED-candidate move op; surviving moves must be native (global gid).");

        return new IrEditScript(IrNodeList.From(result), script.NoteOps);
    }

    /// <summary>
    /// Rewrite a reviewer's pairwise edit script so it contains NO structural op
    /// (<see cref="IrEditOpKind.MoveBlock"/> / <see cref="IrEditOpKind.MoveModifyBlock"/> /
    /// <see cref="IrEditOpKind.SplitBlock"/> / <see cref="IrEditOpKind.MergeBlock"/>) — lowering each to
    /// equivalent <see cref="IrEditOpKind.InsertBlock"/> / <see cref="IrEditOpKind.DeleteBlock"/> ops,
    /// PRESERVING op order (the preceding-anchor insert routing depends on order).
    /// <para><b>Why.</b> The N-way merger groups each reviewer's ops by base anchor
    /// (<see cref="GroupByBaseAnchor"/>, keyed on <see cref="IrEditOp.LeftAnchor"/>) and routes right-only
    /// inserts by preceding anchor (<see cref="GroupInsertsByPrecedingAnchor"/>, only for
    /// <see cref="IrEditOpKind.InsertBlock"/>). A structural op with a NULL left anchor — a move/move-modify
    /// DESTINATION (RightAnchor set, LeftAnchor null) or a <see cref="IrEditOpKind.MergeBlock"/>
    /// (RightAnchor set, <see cref="IrEditOp.SplitMergeAnchors"/> = left anchors, LeftAnchor null) — reaches
    /// NEITHER path and was silently DROPPED, losing the moved/merged content on accept with no conflict
    /// recorded (reject ≡ base still held — no corruption). Lowering to plain Insert/Delete makes the
    /// existing Equal/Modify/FormatOnly/Insert/Delete composition handle them with no loss.</para>
    /// <para><b>Lowering rules</b> (each preserves accept ≡ that reviewer's right body and reject ≡ base):
    /// <list type="bullet">
    /// <item>Move/MoveModify SOURCE (IsMoveSource=true, LeftAnchor=X) → <c>DeleteBlock(X)</c>.</item>
    /// <item>Move/MoveModify DEST (IsMoveSource=false, RightAnchor=Y) → <c>InsertBlock(Y)</c>. A MoveModify
    /// dest's in-move TokenDiff is discarded — inserting the whole edited destination block is correct
    /// (accept shows the moved+edited text; the paired source delete removes the original → net
    /// move-with-edit).</item>
    /// <item>Split (LeftAnchor=X, SplitMergeAnchors=[R1..Rn] RIGHT anchors) → <c>DeleteBlock(X)</c> then
    /// <c>InsertBlock(R1)</c> … <c>InsertBlock(Rn)</c> in order, at the split op's position.</item>
    /// <item>Merge (RightAnchor=Y, SplitMergeAnchors=[L1..Ln] LEFT anchors) → <c>DeleteBlock(L1)</c> …
    /// <c>DeleteBlock(Ln)</c> then <c>InsertBlock(Y)</c> in order. This DELETES the consumed left base
    /// blocks (otherwise they passed through as EqualBlock and the merge was ignored) and INSERTS the
    /// merged result.</item>
    /// </list></para>
    /// <para><b>v1 limitation (documented).</b> A single reviewer's move/split/merge therefore consolidates
    /// as del/ins rather than native <c>w:moveFrom</c>/<c>w:moveTo</c> or split/merge markup; content is
    /// fully preserved and round-trips. Two reviewers moving the SAME base block become competing
    /// deletes+inserts — the source-block delete collides in byBase → a recorded conflict, no loss. Native
    /// cross-reviewer move/split/merge composition is a follow-on.</para>
    /// </summary>
    internal static IrEditScript LowerStructuralOps(IrEditScript script)
    {
        var lowered = new List<IrEditOp>(script.Operations.Count);
        foreach (var op in script.Operations)
            LowerOneStructuralOp(op, lowered);

        System.Diagnostics.Debug.Assert(
            lowered.All(o => o.Kind is not (IrEditOpKind.MoveBlock or IrEditOpKind.MoveModifyBlock
                or IrEditOpKind.SplitBlock or IrEditOpKind.MergeBlock)),
            "LowerStructuralOps must leave no Move/Split/Merge op before composite grouping.");

        // Note scopes (footnotes/endnotes) are not part of body grouping; the composite path does not yet
        // compose them across reviewers, so they pass through unchanged (parity with the pre-lowering shape).
        return new IrEditScript(IrNodeList.From(lowered), script.NoteOps);
    }

    /// <summary>
    /// Append the lowered form of ONE structural op (Move/MoveModify/Split/Merge) — or the op itself for any
    /// other kind — to <paramref name="sink"/>, per the lowering rules documented on
    /// <see cref="LowerStructuralOps"/>. Shared by <see cref="LowerStructuralOps"/> (lowers everything) and
    /// <see cref="ApplyMovePlan"/> (lowers only the move groups the plan did NOT select for native rendering,
    /// plus every Split/Merge).
    /// </summary>
    private static void LowerOneStructuralOp(IrEditOp op, List<IrEditOp> sink)
    {
        switch (op.Kind)
        {
            case IrEditOpKind.MoveBlock:
            case IrEditOpKind.MoveModifyBlock:
                if (op.IsMoveSource == true)
                    // Lower the move SOURCE to a DeleteBlock, but RETAIN MoveGroupId/IsMoveSource as a
                    // marker that this delete is a RELOCATION (the block is reinserted elsewhere by the
                    // same reviewer), not a plain removal. The renderer/verifier/json dispatch on Kind
                    // (DeleteBlock), so the retained fields are inert there; the merger reads them to
                    // distinguish a contested move (2+ reviewers relocate the same base block to
                    // different places) from a consensus removal — see MergeOneBaseBlock.
                    sink.Add(new IrEditOp(IrEditOpKind.DeleteBlock, op.LeftAnchor, null, null,
                        op.MoveGroupId, IsMoveSource: true));
                else
                    sink.Add(new IrEditOp(IrEditOpKind.InsertBlock, null, op.RightAnchor, null, null, null));
                break;

            case IrEditOpKind.SplitBlock:
                // One left base paragraph fans out into N right segments: delete the base block, then
                // insert each right segment in order at this position.
                sink.Add(new IrEditOp(IrEditOpKind.DeleteBlock, op.LeftAnchor, null, null, null, null));
                if (op.SplitMergeAnchors is { } splitRights)
                    foreach (var rightAnchor in splitRights)
                        sink.Add(new IrEditOp(IrEditOpKind.InsertBlock, null, rightAnchor, null, null, null));
                break;

            case IrEditOpKind.MergeBlock:
                // N left base paragraphs fuse into one right paragraph: delete each consumed left block
                // in order, then insert the merged right result.
                if (op.SplitMergeAnchors is { } mergeLefts)
                    foreach (var leftAnchor in mergeLefts)
                        sink.Add(new IrEditOp(IrEditOpKind.DeleteBlock, leftAnchor, null, null, null, null));
                sink.Add(new IrEditOp(IrEditOpKind.InsertBlock, null, op.RightAnchor, null, null, null));
                break;

            default:
                sink.Add(op);
                break;
        }
    }

    /// <summary>
    /// Index every op that carries a non-null <see cref="IrEditOp.LeftAnchor"/> (i.e. it touches a
    /// base block) by that anchor string, tagged with its reviewer index in <paramref name="scripts"/>.
    /// </summary>
    internal static Dictionary<string, List<(int Reviewer, IrEditOp Op)>> GroupByBaseAnchor(
        IReadOnlyList<IrEditScript> scripts)
    {
        var byBase = new Dictionary<string, List<(int Reviewer, IrEditOp Op)>>();
        for (int reviewer = 0; reviewer < scripts.Count; reviewer++)
        {
            foreach (var op in scripts[reviewer].Operations)
            {
                if (op.LeftAnchor is not { } anchor)
                    continue;
                if (!byBase.TryGetValue(anchor, out var list))
                    byBase[anchor] = list = new List<(int, IrEditOp)>();
                list.Add((reviewer, op));
            }
        }
        return byBase;
    }

    /// <summary>Base body blocks in document order, as <c>kind:scope:unid</c> anchor strings.</summary>
    internal static List<string> BaseBlockAnchors(IrDocument baseIr) =>
        baseIr.Body.Blocks.Select(b => b.Anchor.ToString()).ToList();

    // ---- Task 1.3: exact sub-block (token-span) composition ----

    /// <summary>
    /// Compose N reviewers' per-base-paragraph token diffs — all expressed over the SAME base token
    /// stream <c>[0, baseTokenCount)</c> — into a single authored token-op list plus any conflicts
    /// (spec §6 step 4). Walks base positions <c>pos = 0..baseTokenCount</c> inclusive; at each pos it
    /// (1) resolves the reviewers' INSERTs anchored at pos (consensus when ≤1 distinct inserted text,
    /// otherwise a conflict resolved by <paramref name="policy"/>), then (2) for the base token
    /// <c>[pos, pos+1)</c> resolves the reviewers' DELETEs (consensus when 0/1 deleter or all deleters
    /// delete identically, otherwise a conflict). Conflict ids ascend from <paramref name="conflictIdSeed"/>.
    /// </summary>
    /// <remarks>
    /// INVARIANT (span totality): the emitted ops' non-insert left spans tile <c>[0, baseTokenCount)</c>
    /// exactly once — every base token becomes exactly one Equal or one Delete; inserts carry empty
    /// left spans. Verified by <see cref="AssertTilesBase"/> under <c>Debug.Assert</c>.
    /// </remarks>
    internal static List<IrAuthoredTokenOp> ComposeTokenSpans(
        int baseTokenCount,
        IReadOnlyList<(int Reviewer, string Author, IrTokenDiff Diff, IReadOnlyList<IrDiffToken> RightTokens)> reviewerDiffs,
        Docxodus.ConflictResolution policy, string baseAnchor, int conflictIdSeed,
        out List<IrConflict> conflicts)
    {
        var result = new List<IrAuthoredTokenOp>();
        conflicts = new List<IrConflict>();
        int nextConflictId = conflictIdSeed;

        // Replacement convention (matches IrTokenDiffer): the differ emits a replacement as a Delete
        // followed by an Insert whose LeftStart == Delete.LeftEnd (NOT Delete.LeftStart). The
        // immediately-following Insert in the same reviewer's Ops is that Delete's replacement text.
        // Such replacement-Inserts are owned by the delete path (ComposeBaseTokenAt), keyed to the
        // DELETED base span — so we exclude them here from the standalone insert-at-pos handling to
        // avoid double-counting one reviewer's "replace base[a,b) with text" as both a delete-conflict
        // at [a,b) AND an insert-conflict at [b,b). We identify the set by reference identity.
        var replacementInserts = CollectReplacementInserts(reviewerDiffs);

        for (int pos = 0; pos <= baseTokenCount; pos++)
        {
            ComposeInsertsAt(pos, reviewerDiffs, replacementInserts, policy, baseAnchor, ref nextConflictId, result, conflicts);
            if (pos < baseTokenCount)
                ComposeBaseTokenAt(pos, reviewerDiffs, policy, baseAnchor, ref nextConflictId, result, conflicts);
        }

        System.Diagnostics.Debug.Assert(AssertTilesBase(result, baseTokenCount),
            "ComposeTokenSpans: emitted left spans must tile [0, baseTokenCount) exactly once.");
        return result;
    }

    /// <summary>
    /// The set (by reference identity) of every Insert op that is the replacement half of a
    /// Delete+Insert replacement pair, across all reviewers. A replacement Insert is the op that
    /// immediately follows a Delete in the same reviewer's <see cref="IrTokenDiff.Ops"/> and whose
    /// <c>LeftStart == Delete.LeftEnd</c> (the differ's replacement anchoring). These are handled by the
    /// delete path and must NOT separately re-trigger an insert conflict.
    /// </summary>
    private static HashSet<IrTokenOp> CollectReplacementInserts(
        IReadOnlyList<(int Reviewer, string Author, IrTokenDiff Diff, IReadOnlyList<IrDiffToken> RightTokens)> reviewerDiffs)
    {
        var set = new HashSet<IrTokenOp>(ReferenceEqualityComparer.Instance);
        foreach (var rd in reviewerDiffs)
        {
            var ops = rd.Diff.Ops;
            for (int k = 0; k + 1 < ops.Count; k++)
            {
                if (ops[k].Kind == IrTokenOpKind.Delete &&
                    ops[k + 1].Kind == IrTokenOpKind.Insert &&
                    ops[k + 1].LeftStart == ops[k].LeftEnd)
                {
                    set.Add(ops[k + 1]);
                }
            }
        }
        return set;
    }

    /// <summary>Step 1: resolve every reviewer INSERT anchored at base position <paramref name="pos"/>.</summary>
    private static void ComposeInsertsAt(
        int pos,
        IReadOnlyList<(int Reviewer, string Author, IrTokenDiff Diff, IReadOnlyList<IrDiffToken> RightTokens)> reviewerDiffs,
        HashSet<IrTokenOp> replacementInserts,
        Docxodus.ConflictResolution policy, string baseAnchor, ref int nextConflictId,
        List<IrAuthoredTokenOp> result, List<IrConflict> conflicts)
    {
        // Each entry: the inserting reviewer, in reviewer (list) order, with the inserted text.
        // Replacement-Inserts (the Insert half of a Delete+Insert pair) are EXCLUDED — they are owned by
        // the delete path keyed to the deleted base span, so they must not re-trigger an insert conflict.
        var inserters = new List<(int ReviewerIndex, string Author, IrTokenOp Op, string Text)>();
        for (int i = 0; i < reviewerDiffs.Count; i++)
        {
            var rd = reviewerDiffs[i];
            foreach (var op in rd.Diff.Ops)
            {
                if (op.Kind == IrTokenOpKind.Insert && op.LeftStart == pos && !replacementInserts.Contains(op))
                    inserters.Add((rd.Reviewer, rd.Author, op, RightText(op, rd.RightTokens)));
            }
        }
        if (inserters.Count == 0)
            return;

        // Group by inserted right text, preserving first-appearance order (reviewer order).
        var distinctTexts = new List<string>();
        foreach (var ins in inserters)
            if (!distinctTexts.Contains(ins.Text))
                distinctTexts.Add(ins.Text);

        if (distinctTexts.Count <= 1)
        {
            // Consensus (single reviewer, or multiple reviewers inserting the SAME text): one authored
            // Insert per distinct text; the first reviewer in that group is the author/source.
            foreach (var text in distinctTexts)
            {
                var first = inserters.First(x => x.Text == text);
                result.Add(new IrAuthoredTokenOp(first.Op, first.Author, first.ReviewerIndex));
            }
            return;
        }

        // ≥2 distinct text groups → CONFLICT: one competitor per inserting reviewer.
        var competitors = inserters
            .Select(x => new IrConflictCompetitor(x.Author, x.Text))
            .ToList();
        conflicts.Add(new IrConflict(
            nextConflictId++, baseAnchor, pos, pos, policy, IrNodeList.From(competitors)));

        switch (policy)
        {
            case Docxodus.ConflictResolution.BaseWins:
                // Emit NOTHING for the inserts at pos.
                break;
            case Docxodus.ConflictResolution.FirstReviewerWins:
            {
                var first = inserters.OrderBy(x => x.ReviewerIndex).First();
                result.Add(new IrAuthoredTokenOp(first.Op, first.Author, first.ReviewerIndex));
                break;
            }
            case Docxodus.ConflictResolution.StackAll:
                foreach (var ins in inserters.OrderBy(x => x.ReviewerIndex))
                    result.Add(new IrAuthoredTokenOp(ins.Op, ins.Author, ins.ReviewerIndex));
                break;
        }
    }

    /// <summary>Step 2: resolve the fate of base token <c>[pos, pos+1)</c> across reviewer DELETEs.</summary>
    /// <remarks>
    /// <para><b>Per-base-token tiling vs multi-token deletes.</b> A reviewer's Delete may span several
    /// base tokens (e.g. the differ deletes a word AND its trailing separator as <c>Delete(2,4|2,2)</c>).
    /// This step runs once per base token and emits a SINGLE-token Delete <c>[pos, pos+1)</c> for each
    /// deleted base token — never the deleter's whole multi-token span — so the emitted left spans tile
    /// <c>[0, baseTokenCount)</c> exactly once (the totality invariant). A multi-token delete therefore
    /// surfaces as several adjacent single-token Delete ops, authored to the same reviewer.</para>
    /// <para><b>Replacement insert anchoring.</b> A replacement's Insert is logically anchored at
    /// <c>Delete.LeftEnd</c> (one past the last deleted token). To emit it EXACTLY once for a multi-token
    /// replacement, we emit it only when resolving the LAST covered base token (<c>pos == Delete.LeftEnd - 1</c>).
    /// The insert path (<see cref="ComposeInsertsAt"/>) excludes replacement-Inserts, so this is the only
    /// site that emits the replacement text.</para>
    /// </remarks>
    private static void ComposeBaseTokenAt(
        int pos,
        IReadOnlyList<(int Reviewer, string Author, IrTokenDiff Diff, IReadOnlyList<IrDiffToken> RightTokens)> reviewerDiffs,
        Docxodus.ConflictResolution policy, string baseAnchor, ref int nextConflictId,
        List<IrAuthoredTokenOp> result, List<IrConflict> conflicts)
    {
        // Each deleter, in reviewer order, with the covering Delete op and its replacement text ("" for
        // a pure delete). A replacement is the Insert op IMMEDIATELY FOLLOWING the covering Delete in the
        // same reviewer's Ops (the differ anchors it at Delete.LeftEnd, NOT Delete.LeftStart). Built in
        // reviewer (list) order so deleters[0] / deleters.First() == the lowest reviewer index (M1).
        var deleters = new List<(int ReviewerIndex, string Author, IrTokenOp DeleteOp, string Replacement)>();
        for (int i = 0; i < reviewerDiffs.Count; i++)
        {
            var rd = reviewerDiffs[i];
            if (!TryFindCoveringDelete(rd.Diff.Ops, pos, out var delOp, out var replacement, rd.RightTokens))
                continue;
            deleters.Add((rd.Reviewer, rd.Author, delOp!, replacement));
        }

        if (deleters.Count == 0)
        {
            // Base survives.
            result.Add(BaseEqual(pos));
            return;
        }

        // "Identical" deletion: all deleters have byte-equal replacement text (all pure deletes, or
        // all delete-and-replace with the same inserted text). Pure-delete vs delete-with-replacement
        // are NOT identical (one has "" replacement, the other does not).
        bool allIdentical = deleters.All(d => d.Replacement == deleters[0].Replacement);

        if (deleters.Count == 1 || allIdentical)
        {
            // Consensus: delete THIS base token once, authored to the first deleter (lowest reviewer
            // index). If this is a delete-with-replacement, also emit its replacement insert — but only
            // when resolving the deleter's LAST covered token, so a multi-token replacement emits the
            // text exactly once. The insert path excludes replacement-Inserts, so this is the only site.
            var first = deleters[0];
            result.Add(SingleTokenDelete(pos, first.Author, first.ReviewerIndex));
            EmitReplacementInsertIfLast(pos, first.DeleteOp, first.Author, first.ReviewerIndex, reviewerDiffs, result);
            return;
        }

        // ≥2 deleters that are NOT identical → CONFLICT (recorded per base token of the contested span).
        var competitors = deleters
            .Select(d => new IrConflictCompetitor(d.Author, d.Replacement))
            .ToList();
        conflicts.Add(new IrConflict(
            nextConflictId++, baseAnchor, pos, pos + 1, policy, IrNodeList.From(competitors)));

        switch (policy)
        {
            case Docxodus.ConflictResolution.BaseWins:
                // Base survives.
                result.Add(BaseEqual(pos));
                break;
            case Docxodus.ConflictResolution.FirstReviewerWins:
            {
                // First deleter (lowest reviewer index): delete this base token AND emit its replacement
                // insert (once, at the deleter's last covered token) so the winning replacement lands.
                var first = deleters[0];
                result.Add(SingleTokenDelete(pos, first.Author, first.ReviewerIndex));
                EmitReplacementInsertIfLast(pos, first.DeleteOp, first.Author, first.ReviewerIndex, reviewerDiffs, result);
                break;
            }
            case Docxodus.ConflictResolution.StackAll:
            {
                // A base token can only be deleted once: delete THIS base token once (authored to the
                // first deleter). The competing REPLACEMENT inserts are excluded from the insert path, so
                // stack them here (in reviewer order), each emitted once at its own deleter's last token.
                var first = deleters[0];
                result.Add(SingleTokenDelete(pos, first.Author, first.ReviewerIndex));
                foreach (var d in deleters)
                    EmitReplacementInsertIfLast(pos, d.DeleteOp, d.Author, d.ReviewerIndex, reviewerDiffs, result);
                break;
            }
        }
    }

    /// <summary>
    /// Emit <paramref name="reviewerIndex"/>'s replacement Insert for <paramref name="delOp"/> — but only
    /// when <paramref name="pos"/> is the deleter's LAST covered base token (<c>pos == delOp.LeftEnd - 1</c>),
    /// so a multi-token replacement emits its inserted text exactly once.
    /// </summary>
    private static void EmitReplacementInsertIfLast(
        int pos, IrTokenOp delOp, string author, int reviewerIndex,
        IReadOnlyList<(int Reviewer, string Author, IrTokenDiff Diff, IReadOnlyList<IrDiffToken> RightTokens)> reviewerDiffs,
        List<IrAuthoredTokenOp> result)
    {
        if (pos != delOp.LeftEnd - 1)
            return;
        var ins = FindReplacementInsert(reviewerDiffs, reviewerIndex, delOp);
        if (ins is not null)
            result.Add(new IrAuthoredTokenOp(ins, author, reviewerIndex));
    }

    /// <summary>
    /// The replacement Insert op for <paramref name="delOp"/> in reviewer <paramref name="reviewerIndex"/>:
    /// the op immediately following <paramref name="delOp"/> in that reviewer's Ops when it is an Insert
    /// anchored at <c>delOp.LeftEnd</c>, else null. Mirrors <see cref="TryFindCoveringDelete"/>'s pairing
    /// rule; used to RE-emit a replacement's text since the insert path excludes replacement-Inserts.
    /// </summary>
    private static IrTokenOp? FindReplacementInsert(
        IReadOnlyList<(int Reviewer, string Author, IrTokenDiff Diff, IReadOnlyList<IrDiffToken> RightTokens)> reviewerDiffs,
        int reviewerIndex, IrTokenOp delOp)
    {
        foreach (var rd in reviewerDiffs)
        {
            if (rd.Reviewer != reviewerIndex)
                continue;
            var ops = rd.Diff.Ops;
            for (int k = 0; k + 1 < ops.Count; k++)
            {
                if (ReferenceEquals(ops[k], delOp) &&
                    ops[k + 1].Kind == IrTokenOpKind.Insert &&
                    ops[k + 1].LeftStart == delOp.LeftEnd)
                {
                    return ops[k + 1];
                }
            }
            return null;
        }
        return null;
    }

    /// <summary>A base-sourced Equal op covering base token <paramref name="pos"/> (Author "", reviewer -1).</summary>
    private static IrAuthoredTokenOp BaseEqual(int pos) =>
        new(new IrTokenOp(IrTokenOpKind.Equal, pos, pos + 1, pos, pos + 1), "", -1);

    /// <summary>
    /// A single-base-token Delete <c>[pos, pos+1)</c> authored to <paramref name="author"/> /
    /// <paramref name="reviewerIndex"/>. Always single-token (never a reviewer's whole multi-token Delete
    /// span) so the per-base-token loop preserves the <c>[0, baseTokenCount)</c> tiling invariant.
    /// </summary>
    private static IrAuthoredTokenOp SingleTokenDelete(int pos, string author, int reviewerIndex) =>
        new(new IrTokenOp(IrTokenOpKind.Delete, pos, pos + 1, pos, pos), author, reviewerIndex);

    /// <summary>
    /// Find the Delete op in <paramref name="ops"/> whose left span covers base token
    /// <paramref name="pos"/> (<c>LeftStart &lt;= pos &lt; LeftEnd</c>), and resolve its replacement
    /// text. The replacement is the op IMMEDIATELY FOLLOWING that Delete in <paramref name="ops"/> when
    /// that op is an Insert anchored at <c>Delete.LeftEnd</c> (the differ's Delete→Insert replacement
    /// convention) — concatenated right text — else "" (pure delete). Op-list adjacency is the
    /// authoritative pairing rule; the <c>LeftStart == Delete.LeftEnd</c> check additionally guards
    /// against an unrelated standalone Insert that merely happens to follow the Delete in Ops.
    /// </summary>
    private static bool TryFindCoveringDelete(
        IrNodeList<IrTokenOp> ops, int pos, out IrTokenOp? delOp, out string replacement,
        IReadOnlyList<IrDiffToken> rightTokens)
    {
        for (int k = 0; k < ops.Count; k++)
        {
            var op = ops[k];
            if (op.Kind != IrTokenOpKind.Delete || op.LeftStart > pos || pos >= op.LeftEnd)
                continue;

            delOp = op;
            replacement =
                k + 1 < ops.Count &&
                ops[k + 1].Kind == IrTokenOpKind.Insert &&
                ops[k + 1].LeftStart == op.LeftEnd
                    ? RightText(ops[k + 1], rightTokens)
                    : "";
            return true;
        }

        delOp = null;
        replacement = "";
        return false;
    }

    /// <summary>Concatenate the raw <see cref="IrDiffToken.Text"/> of <paramref name="op"/>'s right span.</summary>
    private static string RightText(IrTokenOp op, IReadOnlyList<IrDiffToken> rightTokens)
    {
        if (op.RightLength == 0)
            return "";
        var sb = new System.Text.StringBuilder();
        for (int i = op.RightStart; i < op.RightEnd; i++)
            sb.Append(rightTokens[i].Text);
        return sb.ToString();
    }

    /// <summary>
    /// Debug-only totality check: the concatenation of non-insert (Equal/Delete) left spans, in op
    /// order, equals <c>[0, baseTokenCount)</c> exactly once ascending.
    /// </summary>
    private static bool AssertTilesBase(List<IrAuthoredTokenOp> ops, int baseTokenCount)
    {
        int expected = 0;
        foreach (var authored in ops)
        {
            var op = authored.Op;
            if (op.Kind == IrTokenOpKind.Insert)
                continue;
            if (op.LeftStart != expected)
                return false;
            expected = op.LeftEnd;
        }
        return expected == baseTokenCount;
    }

    // ---- Task 1.4: block-level merge dispatch ----

    /// <summary>
    /// Merge the reviewers' ops anchored at one base block (<paramref name="anchor"/>) into the composite
    /// op stream. Dispatch (spec §6 step 3):
    /// <list type="bullet">
    /// <item><b>Untouched / all-equal</b> → one base-sourced EqualBlock.</item>
    /// <item><b>Single reviewer touched it</b> → that reviewer's op verbatim.</item>
    /// <item><b>Consensus</b> (all touching ops produce the same right result) → one op (first reviewer).</item>
    /// <item><b>All ModifyBlock paragraph edits</b> → token-span composition (<see cref="ComposeTokenSpans"/>):
    /// disjoint word edits compose into one merged ModifyBlock with per-span authorship; overlapping word
    /// edits become token conflicts resolved by <paramref name="policy"/>.</item>
    /// <item><b>Anything else</b> (e.g. delete-vs-modify) → a BLOCK-LEVEL conflict resolved by policy.</item>
    /// </list>
    /// </summary>
    private static void MergeOneBaseBlock(
        string anchor,
        IReadOnlyDictionary<string, List<(int Reviewer, IrEditOp Op)>> byBase,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        IrDocument baseIr,
        Docxodus.ConflictResolution policy,
        IrDiffSettings settings,
        List<IrCompositeOp> ops,
        List<IrConflict> conflicts,
        ref int nextConflictId)
    {
        if (!byBase.TryGetValue(anchor, out var entries) ||
            entries.All(e => e.Op.Kind == IrEditOpKind.EqualBlock))
        {
            ops.Add(EqualOp(anchor));
            return;
        }
        var touched = entries.Where(e => e.Op.Kind != IrEditOpKind.EqualBlock).ToList();

        if (touched.Count == 1)
        {
            var (rev, op) = touched[0];
            ops.Add(EmitOp(op, reviewers[rev].Author, rev));
            return;
        }

        // CONTESTED RELOCATION: 2+ reviewers each RELOCATED this same base block (their move-source ops,
        // lowered to DeleteBlock with IsMoveSource retained, all anchor here) to DIFFERENT places — the
        // reviewers AGREE the block leaves its origin but DISAGREE on its destination. Emit the consensus
        // removal ONCE (so reject ≡ base and the origin is not duplicated) and RECORD a conflict for the
        // placement disagreement; each reviewer's relocating INSERT is routed independently by preceding
        // anchor (EmitInsertsAt), so every destination survives accept — no content loss under any policy.
        // The recorded conflict is deliberately NOT policy-resolved into the op stream (unlike a block
        // EDIT conflict): a BaseWins flip-to-keep would WRONGLY restore a block both reviewers removed.
        if (touched.All(e => e.Op.Kind == IrEditOpKind.DeleteBlock)
            && touched.Count(e => e.Op.IsMoveSource == true) >= 2)
        {
            conflicts.Add(new IrConflict(nextConflictId++, anchor, 0, 0, policy,
                IrNodeList.From(touched.Select(e =>
                    new IrConflictCompetitor(reviewers[e.Reviewer].Author,
                        ContestedBlockText(e.Op, baseIr, settings))))));
            ops.Add(EmitOp(touched[0].Op, reviewers[touched[0].Reviewer].Author, touched[0].Reviewer));
            return;
        }

        if (AllOpsIdentical(touched, reviewers, settings))            // CONSENSUS
        {
            // Tripwire: a multi-reviewer UNCOMPARABLE edit (table, or section-break/opaque modify — anything
            // BlockResultText can't distinguish) must never reach the consensus emit (which keeps only
            // touched[0] and silently drops the rest). v1 routes such edits to the block-level conflict
            // branch; if one ever lands here an edit is being silently dropped (B4 regression).
            System.Diagnostics.Debug.Assert(
                !(touched.Count > 1 && touched.Any(e => IsUncomparableModify(e.Op))),
                "Multi-reviewer uncomparable edit reached consensus emit — a non-paragraph edit is being silently dropped.");
            var (rev, op) = touched[0];
            ops.Add(EmitOp(op, reviewers[rev].Author, rev));
            return;
        }
        if (touched.All(e => e.Op.Kind == IrEditOpKind.ModifyBlock && e.Op.TokenDiff != null)
            && touched.All(e => ParagraphPropsUnchanged(e.Op, baseIr, reviewers[e.Reviewer].Ir)))
        {
            // TOKEN-SPAN COMPOSITION
            var baseTokens = ParagraphBaseTokens(anchor, baseIr, settings);   // IReadOnlyList<IrDiffToken>
            int baseCount = baseTokens.Count;
            var reviewerDiffs = touched.Select(e =>
                (e.Reviewer, reviewers[e.Reviewer].Author, e.Op.TokenDiff!,
                 RightTokensFor(e.Op, reviewers[e.Reviewer].Ir, settings))).ToList();
            var authored = ComposeTokenSpans(baseCount, reviewerDiffs, policy, anchor, nextConflictId, out var blockConflicts);
            nextConflictId += blockConflicts.Count;
            conflicts.AddRange(blockConflicts);
            var merged = new IrTokenDiff(IrNodeList.From(authored.Select(a => a.Op)));
            var structOp = touched[0].Op with { TokenDiff = merged };
            // Each contributing reviewer's INSERT authored-token spans index THAT reviewer's right-token
            // list, so the renderer needs each reviewer's OWN right paragraph anchor (the merged structOp
            // only retains touched[0]'s). Carry them additively per contributing reviewer with a right anchor.
            var sourceRightAnchors = touched
                .Where(e => e.Op.RightAnchor != null)
                .Select(e => new IrSourceRightAnchor(e.Reviewer, e.Op.RightAnchor!))
                .ToList();
            // Only the first conflict id is linked on the op; consumers needing all conflict ids
            // on this op must scan IrCompositeScript.Conflicts independently (by design).
            ops.Add(new IrCompositeOp(structOp, "", touched[0].Reviewer, IrNodeList.From(authored),
                blockConflicts.Count > 0 ? blockConflicts[0].Id : (int?)null,
                IrNodeList.From(sourceRightAnchors)));
            return;
        }
        // BLOCK-LEVEL CONFLICT
        var cid = nextConflictId++;
        conflicts.Add(new IrConflict(cid, anchor, 0, 0, policy,
            IrNodeList.From(touched.Select(e =>
                new IrConflictCompetitor(reviewers[e.Reviewer].Author, BlockResultText(e.Op, reviewers[e.Reviewer].Ir, settings))))));
        switch (policy)
        {
            case Docxodus.ConflictResolution.BaseWins:
                ops.Add(EqualOp(anchor)); break;
            case Docxodus.ConflictResolution.FirstReviewerWins:
                ops.Add(EmitOp(touched[0].Op, reviewers[touched[0].Reviewer].Author, touched[0].Reviewer, cid)); break;
            case Docxodus.ConflictResolution.StackAll:
                EmitStackAllBlockConflict(anchor, touched, reviewers, cid, ops);
                break;
        }
    }

    /// <summary>
    /// Emit the StackAll resolution of a BLOCK-LEVEL conflict so that AT MOST ONE op is anchored to (and
    /// thus consumes/restores on reject) the contested base block <paramref name="anchor"/>.
    /// <para><b>Why not stack every touched op.</b> A delete-vs-edit conflict has <c>touched</c> = a
    /// base-anchored DeleteBlock AND a base-anchored ModifyBlock, BOTH base-restoring on reject. Stacking
    /// both duplicates the base paragraph on reject (and doubles the <c>w:del</c> wrapper). To keep
    /// reject ≡ base, only the FIRST touched op (lowest reviewer index — Modify/Delete/etc.) stays
    /// base-anchored; every other competitor that has RIGHT content is re-emitted as a base-ANCHORLESS
    /// <see cref="IrEditOpKind.InsertBlock"/> (its own reviewer's right block, <c>SourceReviewer</c> set so
    /// the renderer sources from that reviewer), which contributes a <c>w:ins</c>-wrapped block that
    /// vanishes on reject. A pure DeleteBlock competitor (no right content) is dropped — it is already
    /// captured in the recorded conflict, and re-anchoring it would re-introduce the duplication.</para>
    /// </summary>
    private static void EmitStackAllBlockConflict(
        string anchor,
        List<(int Reviewer, IrEditOp Op)> touched,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        int cid,
        List<IrCompositeOp> ops)
    {
        // Capture the pre-emission count so the totality tripwire below scopes to the ops THIS call appends
        // (O(emitted) per call) rather than rescanning the whole accumulated stream (O(N) → O(N²) over the doc).
        int priorCount = ops.Count;

        // touched[0] stays base-anchored verbatim (lowest reviewer index — base-consuming on reject), but
        // CONTRACT-CLEANED: if it is a lowered move-source DeleteBlock its internal relocation marker must
        // not leak into the emitted op (EmitOp strips MoveGroupId/IsMoveSource for Delete/Insert).
        ops.Add(EmitOp(touched[0].Op, reviewers[touched[0].Reviewer].Author, touched[0].Reviewer, cid));

        // Every other competitor that carries right content becomes a base-anchorless InsertBlock (its
        // reviewer's right block), so it adds NOTHING on reject. Pure deletes (no right) are skipped.
        for (int i = 1; i < touched.Count; i++)
        {
            var (rev, op) = touched[i];
            if (op.RightAnchor is not { } rightAnchor)
                continue;
            var insert = new IrEditOp(IrEditOpKind.InsertBlock, null, rightAnchor, null, null, null);
            ops.Add(new IrCompositeOp(insert, reviewers[rev].Author, rev, null, cid));
        }

        // Exactly one op appended by THIS call consumes/restores the base anchor (LeftAnchor == anchor) — the
        // block-level analogue of the token path's AssertTilesBase totality invariant. Scoped to ops appended
        // here (Skip(priorCount)) so the check is O(emitted), not an O(N) rescan of the accumulated stream.
        System.Diagnostics.Debug.Assert(
            ops.Skip(priorCount).Count(o => o.Op.LeftAnchor == anchor) == 1,
            "StackAll block-conflict must emit exactly one base-anchored op (reject ≡ base).");
    }

    /// <summary>A base-sourced EqualBlock op for <paramref name="anchor"/> (Author "", reviewer -1).</summary>
    private static IrCompositeOp EqualOp(string anchor) =>
        new(new IrEditOp(IrEditOpKind.EqualBlock, anchor, anchor, null, null, null), "", -1);

    /// <summary>
    /// Build the FINAL emitted <see cref="IrCompositeOp"/> for a base-block dispatch, stripping the
    /// internal relocation marker (<see cref="IrEditOp.MoveGroupId"/>/<see cref="IrEditOp.IsMoveSource"/>)
    /// that <see cref="LowerStructuralOps"/> retains on a lowered move-source <see cref="IrEditOpKind.DeleteBlock"/>.
    /// <para><b>Why.</b> The marker exists ONLY so <see cref="MergeOneBaseBlock"/>'s contested-relocation
    /// detection can tell a relocation-delete from a plain removal — it is read off the GROUPED (pre-emit)
    /// <c>touched</c>/byBase ops. The documented <see cref="IrEditOp"/> field-presence contract says a
    /// <see cref="IrEditOpKind.DeleteBlock"/>/<see cref="IrEditOpKind.InsertBlock"/> carries NULL move
    /// fields, and the public edit-script JSON (<see cref="IrCompositeScriptJson"/>) is the differentiator
    /// — so the op that actually lands in the emitted <see cref="IrCompositeScript"/> must be contract-clean.
    /// Every block-level emit site that may carry a marked delete routes through here; the detection still
    /// reads the marker off the grouped ops, not the emitted one.</para>
    /// </summary>
    private static IrCompositeOp EmitOp(
        IrEditOp op, string author, int sourceReviewer, int? conflictId = null) =>
        new(StripRelocationMarker(op), author, sourceReviewer, null, conflictId);

    /// <summary>
    /// Return <paramref name="op"/> with the relocation marker cleared when it is a Delete/Insert that
    /// (illegally for the public contract) carries move fields; otherwise return it unchanged. Only a
    /// lowered move-source DeleteBlock ever carries the marker today, but Insert is covered too for
    /// belt-and-suspenders so no future lowering can leak it.
    /// </summary>
    private static IrEditOp StripRelocationMarker(IrEditOp op) =>
        op.Kind is (IrEditOpKind.DeleteBlock or IrEditOpKind.InsertBlock)
            && (op.MoveGroupId is not null || op.IsMoveSource is not null)
            ? op with { MoveGroupId = null, IsMoveSource = null }
            : op;

    /// <summary>
    /// Index every reviewer's RIGHT-POSITIONED op — a right-only <see cref="IrEditOpKind.InsertBlock"/> OR a
    /// NATIVE move DESTINATION (<see cref="IrEditOpKind.MoveBlock"/>/<see cref="IrEditOpKind.MoveModifyBlock"/>
    /// with <see cref="IrEditOp.IsMoveSource"/> false) — by the base anchor it FOLLOWS: the
    /// <see cref="IrEditOp.LeftAnchor"/> of the most recent left-anchored op in that reviewer's script (or ""
    /// for ops at the very top, before any base block). The composite emitter slots each reviewer's
    /// right-positioned ops immediately after that anchor (<see cref="EmitInsertsAt"/>), so two reviewers both
    /// adding after the same base block both appear, attributed, with no conflict. A native move DESTINATION
    /// carries a null left anchor (like an insert), so the existing preceding-anchor tracking already
    /// positions it; <see cref="EmitInsertsAt"/> emits it via <see cref="EmitOp"/> intact (move fields kept).
    /// </summary>
    internal static Dictionary<string, List<(int Reviewer, IrEditOp Op)>> GroupInsertsByPrecedingAnchor(
        IReadOnlyList<IrEditScript> scripts)
    {
        var map = new Dictionary<string, List<(int Reviewer, IrEditOp Op)>>();
        for (int i = 0; i < scripts.Count; i++)
        {
            string preceding = "";
            foreach (var op in scripts[i].Operations)
            {
                bool isRightPositioned = op.Kind == IrEditOpKind.InsertBlock
                    || (op.Kind is IrEditOpKind.MoveBlock or IrEditOpKind.MoveModifyBlock
                        && op.IsMoveSource == false);
                if (isRightPositioned)
                {
                    if (!map.TryGetValue(preceding, out var list)) map[preceding] = list = new();
                    list.Add((i, op));
                }
                else if (op.LeftAnchor != null) preceding = op.LeftAnchor;
            }
        }
        return map;
    }

    /// <summary>
    /// Emit every reviewer's inserts slotted after <paramref name="anchor"/> (in reviewer order, then the
    /// reviewer's own insert order). Disjoint inserts from different reviewers both appear — there is no
    /// insert-vs-insert conflict at the block level (the spec treats block inserts as independent additions).
    /// </summary>
    private static void EmitInsertsAt(
        IReadOnlyDictionary<string, List<(int Reviewer, IrEditOp Op)>> insertsAfter,
        string anchor,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        List<IrCompositeOp> ops)
    {
        if (!insertsAfter.TryGetValue(anchor, out var list)) return;
        foreach (var (rev, op) in list)
            ops.Add(EmitOp(op, reviewers[rev].Author, rev));
    }

    // ---- Task 1.4: result-equivalence + tokenization helpers ----

    /// <summary>
    /// True when every touched op produces the SAME right result, so the reviewers reached the same edit and
    /// there is no conflict. Conservative: any shape we cannot prove equal returns false (falling through to
    /// compose/conflict). A set of ModifyBlock ops is identical iff all resolve to byte-equal right paragraph
    /// text; a set of DeleteBlock ops is identical (all delete the same base block to nothing). Mixed kinds —
    /// including delete-vs-modify — are NOT identical.
    /// </summary>
    private static bool AllOpsIdentical(
        List<(int Reviewer, IrEditOp Op)> touched,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        IrDiffSettings settings)
    {
        var kind = touched[0].Op.Kind;
        if (touched.Any(e => e.Op.Kind != kind))
            return false;

        if (kind == IrEditOpKind.DeleteBlock)
            return true;

        if (kind == IrEditOpKind.ModifyBlock)
        {
            // An UNCOMPARABLE ModifyBlock (no TokenDiff — i.e. a table, or a section-break / opaque modify)
            // must NEVER be treated as identical via the empty-text shortcut: BlockResultText returns "" for
            // any non-paragraph block, so two reviewers editing the SAME base table (even DIFFERENT cells) or
            // the SAME base section break / opaque block — differently — would (mis)compare as ""=="" and
            // short-circuit to a false consensus, silently dropping every reviewer but the first and recording
            // NO conflict (data loss). v1 does NOT compose such blocks across reviewers, so any multi-reviewer
            // edit to them must fall through to the BLOCK-LEVEL conflict branch (BaseWins keeps the base block +
            // records the conflict; the other policies surface a recorded conflict too).
            if (touched.Any(e => IsUncomparableModify(e.Op)))
                return false;

            string first = BlockResultText(touched[0].Op, reviewers[touched[0].Reviewer].Ir, settings);
            return touched.All(e =>
                BlockResultText(e.Op, reviewers[e.Reviewer].Ir, settings) == first);
        }

        return false;
    }

    /// <summary>
    /// True when <paramref name="op"/> is a ModifyBlock whose result text cannot meaningfully distinguish
    /// reviewers — so the empty-text consensus shortcut in <see cref="AllOpsIdentical"/> would falsely fire.
    /// <para>A TABLE modify carries a nested <see cref="IrEditOp.TableDiff"/>; a section-break / opaque modify
    /// carries NEITHER a <see cref="IrEditOp.TokenDiff"/> nor a <see cref="IrEditOp.TableDiff"/>. In every such
    /// case <see cref="BlockResultText"/> returns "" (it only serializes paragraph and table right blocks),
    /// so two reviewers making DIFFERENT edits would (mis)compare as <c>""==""</c> and short-circuit to a
    /// false consensus — silently dropping every reviewer but the first and recording NO conflict (data loss).
    /// A PARAGRAPH modify always carries a non-null <see cref="IrEditOp.TokenDiff"/> and is unaffected; it
    /// composes / reaches genuine text consensus normally.</para>
    /// <para>Because <c>op.TableDiff != null</c> implies <c>op.TokenDiff == null</c>, the single
    /// <c>TokenDiff == null</c> predicate subsumes BOTH tables and section-break/opaque modifies. v1 does not
    /// compose such blocks across reviewers, so multi-reviewer edits to them route to the block-level conflict
    /// branch rather than the phantom-consensus path.</para>
    /// </summary>
    private static bool IsUncomparableModify(IrEditOp op) =>
        op.Kind == IrEditOpKind.ModifyBlock && op.TokenDiff == null;

    /// <summary>Tokenize the BASE paragraph at <paramref name="anchor"/> exactly as
    /// <see cref="IrEditScriptBuilder"/> tokenizes a Modified pair's left side, so token indices line up with
    /// the reviewers' <see cref="IrEditOp.TokenDiff"/> left coordinates. Non-paragraph (or unknown) anchors
    /// yield an empty list.</summary>
    private static IReadOnlyList<IrDiffToken> ParagraphBaseTokens(
        string anchor, IrDocument baseIr, IrDiffSettings settings) =>
        TokenizeAnchor(anchor, baseIr, settings);

    /// <summary>Tokenize the reviewer's RIGHT paragraph (<paramref name="op"/>'s <see cref="IrEditOp.RightAnchor"/>)
    /// exactly as <see cref="IrEditScriptBuilder"/> tokenizes a Modified pair's right side, so token indices
    /// line up with <paramref name="op"/>'s <see cref="IrEditOp.TokenDiff"/> right coordinates. This supplies
    /// the inserted-text source <see cref="ComposeTokenSpans"/> needs.</summary>
    private static IReadOnlyList<IrDiffToken> RightTokensFor(
        IrEditOp op, IrDocument reviewerIr, IrDiffSettings settings) =>
        op.RightAnchor is { } ra ? TokenizeAnchor(ra, reviewerIr, settings) : System.Array.Empty<IrDiffToken>();

    /// <summary>
    /// True when a ModifyBlock reviewer op changed ONLY paragraph text (and/or run-level formatting),
    /// leaving the paragraph's own <c>w:pPr</c> (style id, justification, indentation, spacing, outline
    /// level, keep/break flags, and the unmodeled-pPr digest) byte-identical to the base. This is the
    /// gate for TOKEN-SPAN COMPOSITION: that path clones the BASE paragraph's pPr, so a reviewer who
    /// ALSO changed pPr would have that change silently dropped — such ops fall through to the
    /// block-level CONFLICT branch (where the policy preserves/surfaces the reviewer's full edit).
    /// <para>The comparison reuses the IR's modeled paragraph format (<see cref="IrParaFormat"/>), the
    /// same per-paragraph formatting the reader records from <c>w:pPr</c> and that the aligner's
    /// FormatOnly classification keys on — its record value-equality covers every modeled property plus
    /// the unmodeled-pPr digest, and is deliberately INDEPENDENT of run/token formatting (which the
    /// compose path handles correctly). A reviewer op with no resolvable right paragraph, or a
    /// non-paragraph anchor, conservatively reports "changed" so it is NOT composed.</para>
    /// </summary>
    private static bool ParagraphPropsUnchanged(IrEditOp op, IrDocument baseIr, IrDocument reviewerIr)
    {
        if (op.LeftAnchor is not { } leftAnchor
            || !baseIr.AnchorIndex.TryGetValue(leftAnchor, out var baseBlock)
            || baseBlock is not IrParagraph basePara)
            return false;
        if (op.RightAnchor is not { } rightAnchor
            || !reviewerIr.AnchorIndex.TryGetValue(rightAnchor, out var rightBlock)
            || rightBlock is not IrParagraph rightPara)
            return false;
        return basePara.Format == rightPara.Format;
    }

    /// <summary>Resolve <paramref name="anchor"/> to a block in <paramref name="ir"/> and tokenize it with the
    /// SAME entry the diff builder used (<see cref="IrDiffTokenizer.Tokenize"/>). Empty for an unknown anchor
    /// or a non-paragraph block.</summary>
    private static IReadOnlyList<IrDiffToken> TokenizeAnchor(string anchor, IrDocument ir, IrDiffSettings settings) =>
        ir.AnchorIndex.TryGetValue(anchor, out var block) && block is IrParagraph p
            ? IrDiffTokenizer.Tokenize(p, settings)
            : System.Array.Empty<IrDiffToken>();

    /// <summary>
    /// The accepted (right) text the reviewer's op would produce: the concatenated token text of the right
    /// paragraph for a ModifyBlock (or any op carrying a <see cref="IrEditOp.RightAnchor"/>), and "" for a
    /// DeleteBlock (no right anchor — the block is removed). Used for conflict-competitor reporting and the
    /// consensus check. Tokenizing for text uses the diff tokenizer so the text matches what the diff sees.
    /// </summary>
    private static string BlockResultText(IrEditOp op, IrDocument reviewerIr, IrDiffSettings settings)
    {
        if (op.RightAnchor is not { } ra)
            return "";
        if (!reviewerIr.AnchorIndex.TryGetValue(ra, out var block))
            return "";
        var sb = new System.Text.StringBuilder();
        switch (block)
        {
            case IrParagraph p:
                foreach (var t in IrDiffTokenizer.Tokenize(p, settings))
                    sb.Append(t.Text);
                break;
            case IrTable tbl:
                // Serialize the table's cell text so a table conflict's competitor ResultText is meaningful
                // (rather than ""): concatenate each cell's paragraph text, cells separated by '␟' (unit
                // separator) and rows by '␞' (record separator) so two differently-edited tables produce
                // distinguishable strings. v1 does not compose table cells; this text is for conflict reporting.
                AppendTableText(tbl, sb, settings);
                break;
        }
        return sb.ToString();
    }

    /// <summary>
    /// The conflict-competitor text for a CONTESTED RELOCATION (2+ reviewers each relocated the same base
    /// block to different places). The competing ops are lowered move-source DeleteBlocks with a NULL
    /// RightAnchor, so <see cref="BlockResultText"/> would return "" for every competitor — leaving the
    /// recorded conflict unable to identify WHICH block is contested. Source the text from the contested
    /// block's LEFT content instead (the base block being relocated, resolved via the op's
    /// <see cref="IrEditOp.LeftAnchor"/>), so the conflict names the relocated block. All competitors share
    /// the same base block, so they all report the same (identifying) text — the disagreement is the
    /// destination, recorded by the per-reviewer routing of the relocating inserts, not the text here.
    /// </summary>
    private static string ContestedBlockText(IrEditOp op, IrDocument baseIr, IrDiffSettings settings)
    {
        if (op.LeftAnchor is not { } la || !baseIr.AnchorIndex.TryGetValue(la, out var block))
            return "";
        var sb = new System.Text.StringBuilder();
        switch (block)
        {
            case IrParagraph p:
                foreach (var t in IrDiffTokenizer.Tokenize(p, settings))
                    sb.Append(t.Text);
                break;
            case IrTable tbl:
                AppendTableText(tbl, sb, settings);
                break;
        }
        return sb.ToString();
    }

    /// <summary>Append <paramref name="tbl"/>'s cell text (row/cell-delimited) to <paramref name="sb"/> for
    /// conflict-competitor reporting. Nested paragraphs in a cell are tokenized with the diff tokenizer so the
    /// text matches what the diff sees; nested tables recurse.</summary>
    private static void AppendTableText(IrTable tbl, System.Text.StringBuilder sb, IrDiffSettings settings)
    {
        bool firstRow = true;
        foreach (var row in tbl.Rows)
        {
            if (!firstRow) sb.Append('␞');
            firstRow = false;
            bool firstCell = true;
            foreach (var cell in row.Cells)
            {
                if (!firstCell) sb.Append('␟');
                firstCell = false;
                foreach (var b in cell.Blocks)
                {
                    if (b is IrParagraph cp)
                        foreach (var t in IrDiffTokenizer.Tokenize(cp, settings))
                            sb.Append(t.Text);
                    else if (b is IrTable nested)
                        AppendTableText(nested, sb, settings);
                }
            }
        }
    }
}
