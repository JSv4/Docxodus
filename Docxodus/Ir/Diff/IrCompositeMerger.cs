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
        //    id so two reviewers' independent native moves never share a w:name. The same touchers map then
        //    decides, per reviewer Split/Merge, NATIVE (sole toucher of every consumed base anchor → keep the
        //    structural op, rendered as native split/merge markup) vs LOWER (colliding → del/ins).
        var touchers = BuildTouchers(rawScripts);
        var plan = PlanMoves(rawScripts, settings, touchers);

        // 3. Transform each script per the plans: NATIVE move groups keep their move ops (with the global gid
        //    substituted) and NATIVE Split/Merge ops pass through; LOWERED move groups and COLLIDING
        //    Split/Merge collapse to Insert/Delete.
        var scripts = rawScripts
            .Select((s, reviewer) => ApplySplitMergePlan(ApplyMovePlan(s, reviewer, plan), reviewer, touchers))
            .ToList();
        var baseOrder = BaseBlockAnchors(baseIr);
        var byBase = GroupByBaseAnchor(scripts);
        var insertsAfter = GroupInsertsByPrecedingAnchor(scripts);
        var nativeMerges = IndexNativeMerges(scripts);

        var ops = new List<IrCompositeOp>();
        var conflicts = new List<IrConflict>();
        var nextConflictId = 1;
        MergeBlockStream(baseOrder, byBase, insertsAfter, nativeMerges, reviewers, baseIr, policy, settings,
            ops, conflicts, ref nextConflictId);

        // 4. NOTE SCOPES: merge the reviewers' footnote/endnote diffs against the shared base note store —
        //    a base-matched note's blocks are themselves a base-anchored block stream, so the SAME grouping +
        //    per-block dispatch runs over each note; reviewer-inserted notes pass through authored. Also
        //    collect every reviewer's note-id correspondence for the renderer's reference rewrite.
        var (noteOps, noteIdMaps) = MergeNoteScopes(
            baseIr, reviewers, rawScripts, policy, settings, conflicts, ref nextConflictId);

        return new IrCompositeScript(IrNodeList.From(ops), IrNodeList.From(conflicts), noteOps, noteIdMaps);
    }

    // ---- N-way note-scope merge ----

    /// <summary>
    /// Merge the reviewers' footnote/endnote diffs — each computed against the SHARED base note store — into
    /// composed per-note op streams. Per kind (footnotes then endnotes):
    /// <list type="bullet">
    /// <item><b>Base-matched notes</b> (a reviewer's <see cref="IrNoteDiff.LeftNoteId"/> non-null) group by
    /// base note id; each touched note runs the SAME <see cref="MergeBlockStream"/> dispatch over the note's
    /// base blocks (reviewers absent from the group are untouched), so disjoint note edits compose, identical
    /// ones reach consensus, and contested ones become recorded conflicts resolved by policy — exactly the
    /// body semantics. Structural ops inside a note (split/merge/move of note paragraphs) are conservatively
    /// LOWERED to del/ins (content-preserving; native in-note structural composition is a follow-on).
    /// Pure id-shift entries (all-EqualBlock) contribute an id-map entry but no composed note diff.</item>
    /// <item><b>Reviewer-inserted notes</b> (null <see cref="IrNoteDiff.LeftNoteId"/>) pass through authored
    /// to their reviewer, ordered by (reviewer, numeric note id) for determinism.</item>
    /// </list>
    /// Every reviewer note diff also yields an <see cref="IrReviewerNoteIdMap"/> entry so the renderer can
    /// rewrite reviewer-sourced body references into the base-anchored output id space.
    /// </summary>
    private static (IrNodeList<IrCompositeNoteDiff>? NoteOps, IrNodeList<IrReviewerNoteIdMap>? NoteIdMaps)
        MergeNoteScopes(
            IrDocument baseIr,
            IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
            IReadOnlyList<IrEditScript> rawScripts,
            Docxodus.ConflictResolution policy,
            IrDiffSettings settings,
            List<IrConflict> conflicts,
            ref int nextConflictId)
    {
        var noteOps = new List<IrCompositeNoteDiff>();
        var idMaps = new List<IrReviewerNoteIdMap>();

        foreach (var kind in new[] { IrNoteKind.Footnote, IrNoteKind.Endnote })
        {
            var baseStore = kind == IrNoteKind.Footnote ? baseIr.Footnotes : baseIr.Endnotes;

            var byBaseNote = new Dictionary<string, List<(int Reviewer, IrNoteDiff Diff)>>();
            var insertedNotes = new List<(int Reviewer, IrNoteDiff Diff)>();
            for (int r = 0; r < rawScripts.Count; r++)
            {
                if (rawScripts[r].NoteOps is not { } nds)
                    continue;
                foreach (var nd in nds)
                {
                    if (nd.Kind != kind)
                        continue;
                    idMaps.Add(new IrReviewerNoteIdMap(r, kind, nd.NoteId, nd.LeftNoteId));
                    if (nd.LeftNoteId is { } baseId)
                    {
                        if (!byBaseNote.TryGetValue(baseId, out var list))
                            byBaseNote[baseId] = list = new List<(int, IrNoteDiff)>();
                        list.Add((r, nd));
                    }
                    else
                    {
                        insertedNotes.Add((r, nd));
                    }
                }
            }

            // Base-matched notes, numeric id ascending (mirrors the two-way NoteOps ordering).
            foreach (var baseId in byBaseNote.Keys.OrderBy(NoteIdSortKey).ThenBy(id => id, StringComparer.Ordinal))
            {
                if (!baseStore.Notes.TryGetValue(baseId, out var baseScope))
                    continue;   // defensive: a LeftNoteId always names a base-store note

                var baseOrder = baseScope.Blocks.Select(b => b.Anchor.ToString()).ToList();

                // Full-length per-reviewer scripts (empty = untouched), with in-note structural ops lowered.
                var scripts = new List<IrEditScript>(reviewers.Count);
                for (int r = 0; r < reviewers.Count; r++)
                    scripts.Add(new IrEditScript(IrNodeList.From(System.Array.Empty<IrEditOp>())));
                foreach (var (r, nd) in byBaseNote[baseId])
                    scripts[r] = LowerStructuralOps(new IrEditScript(nd.Ops));

                var byBase = GroupByBaseAnchor(scripts);
                var insertsAfter = GroupInsertsByPrecedingAnchor(scripts);

                var composed = new List<IrCompositeOp>();
                MergeBlockStream(baseOrder, byBase, insertsAfter, NativeMergeIndex.Empty, reviewers, baseIr,
                    policy, settings, composed, conflicts, ref nextConflictId);

                // A pure id-shift note (every reviewer op EqualBlock) needs NO composed entry: the composite
                // output keeps the BASE id space, so there is nothing to reconcile in the part itself — the
                // id map above already lets the renderer rewrite that reviewer's cloned references.
                if (composed.Any(o => o.Op.Kind is not IrEditOpKind.EqualBlock))
                    noteOps.Add(new IrCompositeNoteDiff(kind, baseId, IrNodeList.From(composed)));
            }

            // Reviewer-inserted notes: authored passthrough, ordered (reviewer, numeric id).
            foreach (var (r, nd) in insertedNotes
                         .OrderBy(e => e.Reviewer)
                         .ThenBy(e => NoteIdSortKey(e.Diff.NoteId))
                         .ThenBy(e => e.Diff.NoteId, StringComparer.Ordinal))
            {
                var authored = nd.Ops.Select(o => new IrCompositeOp(o, reviewers[r].Author, r)).ToList();
                noteOps.Add(new IrCompositeNoteDiff(kind, null, IrNodeList.From(authored), r, nd.NoteId));
            }
        }

        return noteOps.Count == 0 && idMaps.Count == 0
            ? (null, null)
            : (noteOps.Count > 0 ? IrNodeList.From(noteOps) : null,
               idMaps.Count > 0 ? IrNodeList.From(idMaps) : null);
    }

    /// <summary>Numeric sort key for a note id (<c>w:id</c> is an integer in practice; non-numeric ids sort
    /// after all numeric ones, tie-broken by ordinal string comparison at the call sites).</summary>
    private static long NoteIdSortKey(string id) =>
        long.TryParse(id, out var n) ? n : long.MaxValue;

    /// <summary>
    /// A load-bearing composite invariant. Unlike <see cref="System.Diagnostics.Debug.Assert(bool, string)"/>
    /// — which is <c>[Conditional("DEBUG")]</c> and therefore stripped from Release builds (the configuration
    /// this library ships in AND the one CI runs its tests under, <c>dotnet test -c Release</c>) — this is
    /// enforced at runtime in EVERY build. These guard content-integrity invariants (token/table tiling and
    /// totality, no un-lowered structural op, single base-anchored op), where a violation would otherwise
    /// silently emit a CORRUPT or LOSSY consolidated document; the O(n)-per-merged-block check is paid
    /// deliberately to fail fast instead of losing content. A throw here is an engine bug, not caller error.
    /// </summary>
    private static void Invariant(bool condition, string message)
    {
        if (!condition)
            throw new System.InvalidOperationException(
                "DocxDiff consolidate invariant violated (engine bug): " + message);
    }

    /// <summary>
    /// The index of NATIVE (surviving, sole-toucher) <see cref="IrEditOpKind.MergeBlock"/> ops for a block
    /// stream. A MergeBlock has a null <see cref="IrEditOp.LeftAnchor"/> and consumes N base anchors via
    /// <see cref="IrEditOp.SplitMergeAnchors"/>, so the base-anchor-keyed grouping cannot dispatch it —
    /// <see cref="MergeBlockStream"/> emits it when the walk reaches its FIRST consumed anchor
    /// (<see cref="ByFirstAnchor"/>) and suppresses the block emission (but keeps insert slotting) for the
    /// remaining consumed anchors (<see cref="Consumed"/>).
    /// </summary>
    internal sealed record NativeMergeIndex(
        IReadOnlyDictionary<string, (int Reviewer, IrEditOp Op)> ByFirstAnchor,
        IReadOnlySet<string> Consumed)
    {
        public static readonly NativeMergeIndex Empty = new(
            new Dictionary<string, (int, IrEditOp)>(), new HashSet<string>());
    }

    /// <summary>
    /// Index every surviving (native) <see cref="IrEditOpKind.MergeBlock"/> across the transformed scripts
    /// by its FIRST consumed base anchor, with the remaining consumed anchors in <see cref="NativeMergeIndex.Consumed"/>.
    /// Eligibility (sole toucher of every consumed anchor) was already enforced by
    /// <see cref="ApplySplitMergePlan"/>, so at most one reviewer's merge claims any base anchor here.
    /// </summary>
    internal static NativeMergeIndex IndexNativeMerges(IReadOnlyList<IrEditScript> scripts)
    {
        Dictionary<string, (int, IrEditOp)>? byFirst = null;
        HashSet<string>? consumed = null;
        for (int reviewer = 0; reviewer < scripts.Count; reviewer++)
        {
            foreach (var op in scripts[reviewer].Operations)
            {
                if (op.Kind != IrEditOpKind.MergeBlock || op.SplitMergeAnchors is not { Count: > 0 } lefts)
                    continue;
                byFirst ??= new Dictionary<string, (int, IrEditOp)>(StringComparer.Ordinal);
                consumed ??= new HashSet<string>(StringComparer.Ordinal);
                byFirst[lefts[0]] = (reviewer, op);
                for (int i = 1; i < lefts.Count; i++)
                    consumed.Add(lefts[i]);
            }
        }
        return byFirst == null
            ? NativeMergeIndex.Empty
            : new NativeMergeIndex(byFirst, consumed!);
    }

    /// <summary>
    /// The reusable per-block dispatch loop: walk <paramref name="baseOrder"/> (the base block anchors in
    /// document order), merging each via <see cref="MergeOneBaseBlock"/> and slotting each reviewer's
    /// right-only inserts after the preceding base anchor (<see cref="EmitInsertsAt"/>). Factored out of
    /// <see cref="Merge"/> so the SAME grouping+dispatch runs over a cell's mini-body during per-cell table
    /// composition (FOLLOW-ON B): a cell's paragraph blocks are themselves a base-anchored block stream.
    /// <para><b>Native merges.</b> A surviving <see cref="IrEditOpKind.MergeBlock"/> (sole-toucher, kept by
    /// <see cref="ApplySplitMergePlan"/>) is emitted when the walk reaches its FIRST consumed base anchor;
    /// the remaining consumed anchors emit no block op (the merge op subsumes them) but still slot pending
    /// right-only inserts, which therefore land AFTER the whole merge markup — positionally coarse but safe
    /// (an insert between merge members would corrupt the deleted-pilcrow reject reconstruction).</para>
    /// </summary>
    private static void MergeBlockStream(
        IReadOnlyList<string> baseOrder,
        IReadOnlyDictionary<string, List<(int Reviewer, IrEditOp Op)>> byBase,
        IReadOnlyDictionary<string, List<(int Reviewer, IrEditOp Op)>> insertsAfter,
        NativeMergeIndex nativeMerges,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        IrDocument baseIr,
        Docxodus.ConflictResolution policy,
        IrDiffSettings settings,
        List<IrCompositeOp> ops,
        List<IrConflict> conflicts,
        ref int nextConflictId)
    {
        EmitInsertsAt(insertsAfter, "", reviewers, ops);
        foreach (var anchor in baseOrder)
        {
            if (nativeMerges.ByFirstAnchor.TryGetValue(anchor, out var merge))
            {
                // A native N:1 merge claims this anchor (its first consumed base paragraph): emit the merge
                // op verbatim, authored to the merging reviewer. The other consumed anchors are suppressed
                // below as the walk reaches them.
                ops.Add(EmitOp(merge.Op, reviewers[merge.Reviewer].Author, merge.Reviewer));
            }
            else if (!nativeMerges.Consumed.Contains(anchor))
            {
                MergeOneBaseBlock(anchor, byBase, reviewers, baseIr, policy, settings, ops, conflicts, ref nextConflictId);
            }
            // A merge-consumed anchor emits no block op, but its pending inserts still slot here.
            EmitInsertsAt(insertsAfter, anchor, reviewers, ops);
        }
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
    /// Build the cross-reviewer TOUCHERS map: base anchor → the set of reviewers with a non-Equal op
    /// anchored at that base block. The single eligibility source for every native-structural decision
    /// (moves, splits, merges): an op whose consumed base block(s) are touched ONLY by its own reviewer is
    /// non-colliding and may render natively.
    /// <para>A move SOURCE (IsMoveSource=true) anchors at its source base block via LeftAnchor; a ModifyBlock/
    /// DeleteBlock/FormatOnlyBlock/SplitBlock also carries the base block's LeftAnchor; a move DEST /
    /// InsertBlock has a null LeftAnchor and so never marks a base block as touched (it is right-positioned
    /// content). A MergeBlock is the exception the LeftAnchor path MISSES: its own LeftAnchor is null
    /// (RightAnchor carries the merged result) and it CONSUMES N base paragraphs carried in
    /// SplitMergeAnchors (left anchors). Those consumed anchors must be registered as touched too —
    /// otherwise a reviewer who MERGES L1+L2 is invisible at L1/L2, and another reviewer's MOVE of L1
    /// would be misclassified as a sole-toucher native move, emitting a w:moveTo with no coherent paired
    /// w:moveFrom once the merge's lowered deletes resolve (an orphaned half-move). SplitBlock needs no
    /// such branch: its LeftAnchor IS its single consumed base paragraph, so the main path already counts
    /// it (a SplitBlock's SplitMergeAnchors are RIGHT segment anchors, not base blocks).</para>
    /// </summary>
    internal static Dictionary<string, HashSet<int>> BuildTouchers(IReadOnlyList<IrEditScript> rawScripts)
    {
        var touchers = new Dictionary<string, HashSet<int>>();
        for (int reviewer = 0; reviewer < rawScripts.Count; reviewer++)
        {
            foreach (var op in rawScripts[reviewer].Operations)
            {
                if (op.Kind == IrEditOpKind.MergeBlock && op.SplitMergeAnchors is { } mergeLefts)
                {
                    foreach (var la in mergeLefts)
                    {
                        if (!touchers.TryGetValue(la, out var ms))
                            touchers[la] = ms = new HashSet<int>();
                        ms.Add(reviewer);
                    }
                    continue;
                }
                if (op.Kind == IrEditOpKind.EqualBlock || op.LeftAnchor is not { } anchor)
                    continue;
                if (!touchers.TryGetValue(anchor, out var set))
                    touchers[anchor] = set = new HashSet<int>();
                set.Add(reviewer);
            }
        }
        return touchers;
    }

    /// <summary>True when base anchor <paramref name="anchor"/> is touched by reviewer
    /// <paramref name="reviewer"/> ALONE (see <see cref="BuildTouchers"/>).</summary>
    private static bool SoleToucher(
        IReadOnlyDictionary<string, HashSet<int>> touchers, string anchor, int reviewer) =>
        touchers.TryGetValue(anchor, out var set) && set.Count == 1 && set.Contains(reviewer);

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
    internal static MovePlan PlanMoves(
        IReadOnlyList<IrEditScript> rawScripts, IrDiffSettings settings,
        IReadOnlyDictionary<string, HashSet<int>>? touchersByBaseAnchor = null)
    {
        var touchers = touchersByBaseAnchor ?? BuildTouchers(rawScripts);

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
                    if (SoleToucher(touchers, source, reviewer))
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
    /// move group collapses to Insert/Delete via the per-op lowering rules. Split/Merge ops pass through
    /// UNCHANGED — their native-vs-lower decision belongs to <see cref="ApplySplitMergePlan"/>. Op order is
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
            else if (op.Kind is IrEditOpKind.SplitBlock or IrEditOpKind.MergeBlock)
            {
                // Split/Merge are decided by ApplySplitMergePlan (sole-toucher → native, else lowered).
                result.Add(op);
            }
            else
            {
                // LOWERED move group → del/ins (LowerStructuralOps' per-op rules); all other kinds verbatim.
                LowerOneStructuralOp(op, result);
            }
        }

        Invariant(
            result.All(o =>
                o.Kind is not (IrEditOpKind.MoveBlock or IrEditOpKind.MoveModifyBlock)
                || (o.MoveGroupId is { } g && plan.GlobalGidByGroup.Values.Contains(g))),
            "ApplyMovePlan must leave no LOWERED-candidate move op; surviving moves must be native (global gid).");

        return new IrEditScript(IrNodeList.From(result), script.NoteOps);
    }

    /// <summary>
    /// Decide each of reviewer <paramref name="reviewer"/>'s Split/Merge ops NATIVE vs LOWER, mirroring the
    /// move plan's single eligibility predicate: the op is NATIVE iff the reviewer is the SOLE toucher of
    /// EVERY base paragraph it consumes — a <see cref="IrEditOpKind.SplitBlock"/> consumes its
    /// <see cref="IrEditOp.LeftAnchor"/>; a <see cref="IrEditOpKind.MergeBlock"/> consumes each of its
    /// <see cref="IrEditOp.SplitMergeAnchors"/> (left anchors). A NATIVE split/merge passes through verbatim
    /// and renders as native split/merge markup (inserted/deleted pilcrows + per-segment token spans — the
    /// SAME shape the two-way renderer produces); a COLLIDING one lowers to del/ins so the existing
    /// composition/conflict machinery handles the contention with no loss. Op order preserved.
    /// </summary>
    internal static IrEditScript ApplySplitMergePlan(
        IrEditScript script, int reviewer, IReadOnlyDictionary<string, HashSet<int>> touchers)
    {
        var result = new List<IrEditOp>(script.Operations.Count);
        foreach (var op in script.Operations)
        {
            switch (op.Kind)
            {
                case IrEditOpKind.SplitBlock when op.LeftAnchor is { } splitLeft
                    && SoleToucher(touchers, splitLeft, reviewer):
                case IrEditOpKind.MergeBlock when op.SplitMergeAnchors is { Count: > 0 } mergeLefts
                    && mergeLefts.All(la => SoleToucher(touchers, la, reviewer)):
                    result.Add(op);   // NATIVE
                    break;
                case IrEditOpKind.SplitBlock:
                case IrEditOpKind.MergeBlock:
                    LowerOneStructuralOp(op, result);   // COLLIDING → del/ins
                    break;
                default:
                    result.Add(op);
                    break;
            }
        }
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
    /// <para><b>When lowering applies.</b> The consolidate pipeline keeps a NON-COLLIDING (sole-toucher)
    /// reviewer move/split/merge NATIVE (see <see cref="PlanMoves"/>/<see cref="ApplySplitMergePlan"/>) and
    /// lowers only the COLLIDING ones — two reviewers moving/merging the same base block become competing
    /// deletes+inserts, whose source-block deletes collide in byBase → a recorded conflict, no loss. This
    /// method lowers EVERYTHING unconditionally (the conservative transform, kept for tests and as the
    /// documentation of the per-op rules the planners reuse via <see cref="LowerOneStructuralOp"/>).</para>
    /// </summary>
    internal static IrEditScript LowerStructuralOps(IrEditScript script)
    {
        var lowered = new List<IrEditOp>(script.Operations.Count);
        foreach (var op in script.Operations)
            LowerOneStructuralOp(op, lowered);

        Invariant(
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

        Invariant(AssertTilesBase(result, baseTokenCount),
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
            Invariant(
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
        // TABLE COMPOSITION (FOLLOW-ON B): 2+ reviewers edited the SAME base table. DISJOINT cross-reviewer
        // cell edits compose inline (each lands, authored); SAME-cell edits by 2+ reviewers conflict per
        // policy. Column adds/removes compose too (anchor-based cell pairing + authored InsertCell/DeleteCell
        // with native w:cellIns/w:cellDel markup). Fires only when every toucher is a table ModifyBlock AND
        // row moves are uncontested — otherwise we FALL BACK to the block-level conflict (no silent loss; the
        // base table is kept under BaseWins and the disagreement is recorded under every policy). An
        // UNCONTESTED reviewer MovedRow composes (lowered to del+ins rows, the two-way shape).
        if (touched.All(e => e.Op.Kind == IrEditOpKind.ModifyBlock && e.Op.TableDiff != null)
            && baseIr.AnchorIndex.TryGetValue(anchor, out var baseBlk) && baseBlk is IrTable baseTable
            && MovedRowsComposable(touched))
        {
            var reviewerTableDiffs = touched
                .Select(e => (e.Reviewer, reviewers[e.Reviewer].Author, e.Op.TableDiff!))
                .ToList();
            var merged = ComposeTableDiffs(
                anchor, baseTable, reviewerTableDiffs, reviewers, baseIr, policy, settings,
                ref nextConflictId, out var tableConflicts, out var authoredRows);
            conflicts.AddRange(tableConflicts);

            Invariant(AssertTilesBaseTable(authoredRows, baseTable),
                "ComposeTableDiffs: authored rows/cells must tile the base table's rows/cells exactly once.");

            var structOp = touched[0].Op with { TableDiff = merged };
            ops.Add(new IrCompositeOp(structOp, "", touched[0].Reviewer,
                AuthoredTokens: null,
                ConflictId: tableConflicts.Count > 0 ? tableConflicts[0].Id : (int?)null,
                SourceRightAnchors: null,
                AuthoredRows: IrNodeList.From(authoredRows)));
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
        Invariant(
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
                else if (op.Kind == IrEditOpKind.MergeBlock && op.SplitMergeAnchors is { Count: > 0 } lefts)
                {
                    // A native MergeBlock consumes N base paragraphs but carries a null LeftAnchor; an insert
                    // FOLLOWING it in the reviewer's script must slot after the LAST consumed anchor (i.e.
                    // after the whole merged region), not before it.
                    preceding = lefts[lefts.Count - 1];
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

    // ---- FOLLOW-ON B: per-cell table composition ----

    /// <summary>
    /// Row-move composability gate (the row analogue of the block move plan's sole-toucher predicate): true
    /// when every reviewer <see cref="IrRowOpKind.MovedRow"/>'s SOURCE base row is touched (non-Equal,
    /// non-Insert row op) by that reviewer ALONE. A composable MovedRow is then LOWERED during row indexing
    /// (source → DeleteRow at the base row, destination → InsertRow at the new position — exactly the del+ins
    /// shape the two-way renderer itself uses for MovedRow), so the existing row dispatch composes it with the
    /// other reviewers' row edits. A CONTESTED row move (another reviewer modifies/deletes/moves the same base
    /// row) forces the FALL BACK to the whole-table block conflict — no silent loss: the block-conflict branch
    /// keeps the base table under BaseWins and records the disagreement under every policy.
    /// </summary>
    private static bool MovedRowsComposable(List<(int Reviewer, IrEditOp Op)> touched)
    {
        // Left row anchor → reviewers with a base-row-consuming row op there (Modify/Delete/MovedRow-source).
        var rowTouchers = new Dictionary<string, HashSet<int>>();
        foreach (var (reviewer, op) in touched)
            if (op.TableDiff is { } td)
                foreach (var rowOp in td.RowOps)
                    if (rowOp.Kind is not (IrRowOpKind.EqualRow or IrRowOpKind.InsertRow)
                        && rowOp.LeftRowAnchor is { } la)
                    {
                        if (!rowTouchers.TryGetValue(la, out var set))
                            rowTouchers[la] = set = new HashSet<int>();
                        set.Add(reviewer);
                    }

        foreach (var (reviewer, op) in touched)
            if (op.TableDiff is { } td)
                foreach (var rowOp in td.RowOps)
                    if (rowOp.Kind == IrRowOpKind.MovedRow && rowOp.IsMoveSource == true
                        && rowOp.LeftRowAnchor is { } src
                        && !(rowTouchers.TryGetValue(src, out var set) && set.Count == 1 && set.Contains(reviewer)))
                        return false;
        return true;
    }

    /// <summary>
    /// Compose N reviewers' table diffs (all against the SAME base table) into a MERGED
    /// <see cref="IrTableDiff"/> (the apply/JSON truth) plus a parallel <paramref name="authoredRows"/> view
    /// (renderer/revision attribution). Rows align by BASE row anchor; per base row 0 touch → EqualRow,
    /// 1 ModifyRow → that reviewer's cells (authored), 1 DeleteRow (others equal) → DeleteRow, ≥2 ModifyRow →
    /// per-cell compose (the recursion), delete-vs-modify same row → row-level conflict (policy-resolved),
    /// ≥2 deletes same row → consensus delete. InsertRows by different reviewers route by preceding base-row
    /// anchor (both appear, authored, no conflict). Cell conflict ids ascend from
    /// <paramref name="nextConflictId"/>.
    /// </summary>
    private static IrTableDiff ComposeTableDiffs(
        string tableAnchor,
        IrTable baseTable,
        List<(int Reviewer, string Author, IrTableDiff Diff)> reviewerDiffs,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        IrDocument baseIr,
        Docxodus.ConflictResolution policy,
        IrDiffSettings settings,
        ref int nextConflictId,
        out List<IrConflict> conflicts,
        out List<IrAuthoredRowOp> authoredRows)
    {
        conflicts = new List<IrConflict>();
        authoredRows = new List<IrAuthoredRowOp>();
        var mergedRowOps = new List<IrRowOp>();

        // Index each reviewer's row ops by BASE row anchor, and collect their right-only InsertRows keyed by
        // the preceding base row anchor (or "" before the first base row).
        var byBaseRow = new Dictionary<string, List<(int Reviewer, string Author, IrRowOp Op)>>();
        var insertRowsAfter = new Dictionary<string, List<(int Reviewer, string Author, IrRowOp Op)>>();
        foreach (var (reviewer, author, diff) in reviewerDiffs)
        {
            string preceding = "";
            foreach (var rowOp in diff.RowOps)
            {
                if (rowOp.Kind == IrRowOpKind.InsertRow)
                {
                    if (!insertRowsAfter.TryGetValue(preceding, out var il)) insertRowsAfter[preceding] = il = new();
                    il.Add((reviewer, author, rowOp));
                    continue;
                }
                if (rowOp.Kind == IrRowOpKind.MovedRow)
                {
                    // An UNCONTESTED row move (guaranteed by the MovedRowsComposable gate) lowers to the
                    // del+ins row shape the two-way renderer itself produces for MovedRow: the SOURCE half
                    // becomes a DeleteRow at the base row, the DESTINATION half an InsertRow slotted after
                    // the preceding base row — so the move composes with other reviewers' row edits and
                    // round-trips (reject restores the row at its old position, accept shows the new one).
                    if (rowOp.IsMoveSource == true && rowOp.LeftRowAnchor is { } src)
                    {
                        if (!byBaseRow.TryGetValue(src, out var ml)) byBaseRow[src] = ml = new();
                        ml.Add((reviewer, author, new IrRowOp(IrRowOpKind.DeleteRow, src, null, null)));
                        preceding = src;
                    }
                    else if (rowOp.RightRowAnchor is { } dst)
                    {
                        if (!insertRowsAfter.TryGetValue(preceding, out var dl)) insertRowsAfter[preceding] = dl = new();
                        dl.Add((reviewer, author, new IrRowOp(IrRowOpKind.InsertRow, null, dst, null)));
                    }
                    continue;
                }
                if (rowOp.LeftRowAnchor is { } la)
                {
                    if (!byBaseRow.TryGetValue(la, out var bl)) byBaseRow[la] = bl = new();
                    bl.Add((reviewer, author, rowOp));
                    preceding = la;
                }
            }
        }

        EmitComposedInsertRows(insertRowsAfter, "", mergedRowOps, authoredRows);
        foreach (var baseRow in baseTable.Rows)
        {
            string rowAnchor = baseRow.Anchor.ToString();
            ComposeOneBaseRow(rowAnchor, baseRow, byBaseRow, reviewers, baseIr, policy, settings,
                ref nextConflictId, conflicts, mergedRowOps, authoredRows);
            EmitComposedInsertRows(insertRowsAfter, rowAnchor, mergedRowOps, authoredRows);
        }

        return new IrTableDiff(IrNodeList.From(mergedRowOps));
    }

    /// <summary>Emit every reviewer's InsertRows slotted after <paramref name="anchor"/> (reviewer order):
    /// both the merged row op and its authored view. Disjoint inserted rows from different reviewers all
    /// appear — no insert-vs-insert row conflict (block-insert convention).</summary>
    private static void EmitComposedInsertRows(
        IReadOnlyDictionary<string, List<(int Reviewer, string Author, IrRowOp Op)>> insertRowsAfter,
        string anchor, List<IrRowOp> mergedRowOps, List<IrAuthoredRowOp> authoredRows)
    {
        if (!insertRowsAfter.TryGetValue(anchor, out var list)) return;
        foreach (var (reviewer, author, op) in list)
        {
            mergedRowOps.Add(op);
            authoredRows.Add(new IrAuthoredRowOp(
                IrRowOpKind.InsertRow, null, reviewer, author, null, op.RightRowAnchor));
        }
    }

    /// <summary>Dispatch one base row across reviewers (see <see cref="ComposeTableDiffs"/>).</summary>
    private static void ComposeOneBaseRow(
        string rowAnchor, IrRow baseRow,
        IReadOnlyDictionary<string, List<(int Reviewer, string Author, IrRowOp Op)>> byBaseRow,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        IrDocument baseIr, Docxodus.ConflictResolution policy, IrDiffSettings settings,
        ref int nextConflictId, List<IrConflict> conflicts,
        List<IrRowOp> mergedRowOps, List<IrAuthoredRowOp> authoredRows)
    {
        // The reviewers who did something non-Equal to this base row.
        var touched = byBaseRow.TryGetValue(rowAnchor, out var entries)
            ? entries.Where(e => e.Op.Kind != IrRowOpKind.EqualRow).ToList()
            : new List<(int Reviewer, string Author, IrRowOp Op)>();

        if (touched.Count == 0)
        {
            // Untouched (every reviewer Equal / absent) → base row verbatim.
            mergedRowOps.Add(new IrRowOp(IrRowOpKind.EqualRow, rowAnchor, rowAnchor, null));
            authoredRows.Add(new IrAuthoredRowOp(IrRowOpKind.EqualRow, rowAnchor, -1, "", null));
            return;
        }

        var modifiers = touched.Where(e => e.Op.Kind == IrRowOpKind.ModifyRow).ToList();
        var deleters = touched.Where(e => e.Op.Kind == IrRowOpKind.DeleteRow).ToList();

        // delete-vs-modify SAME row → row-level conflict (policy-resolved like the block conflict).
        if (deleters.Count > 0 && modifiers.Count > 0)
        {
            int cid = nextConflictId++;
            conflicts.Add(new IrConflict(cid, rowAnchor, 0, 0, policy,
                IrNodeList.From(touched.Select(e => new IrConflictCompetitor(
                    e.Author, RowResultText(baseRow, settings))))));
            switch (policy)
            {
                case Docxodus.ConflictResolution.BaseWins:
                    mergedRowOps.Add(new IrRowOp(IrRowOpKind.EqualRow, rowAnchor, rowAnchor, null));
                    authoredRows.Add(new IrAuthoredRowOp(IrRowOpKind.EqualRow, rowAnchor, -1, "", null));
                    break;
                case Docxodus.ConflictResolution.FirstReviewerWins:
                {
                    // Lowest reviewer index wins: emit that reviewer's row op (modify or delete).
                    var winner = touched.OrderBy(e => e.Reviewer).First();
                    EmitWinnerRow(winner, rowAnchor, baseRow, reviewers, baseIr, policy, settings,
                        ref nextConflictId, conflicts, mergedRowOps, authoredRows);
                    break;
                }
                case Docxodus.ConflictResolution.StackAll:
                {
                    // A row can be consumed/restored on reject by only ONE op: emit the lowest-index toucher's
                    // row (modify or delete). The competing edit is captured in the recorded conflict; emitting
                    // a second base-anchored row op would duplicate the base row on reject.
                    var winner = touched.OrderBy(e => e.Reviewer).First();
                    EmitWinnerRow(winner, rowAnchor, baseRow, reviewers, baseIr, policy, settings,
                        ref nextConflictId, conflicts, mergedRowOps, authoredRows);
                    break;
                }
            }
            return;
        }

        // ≥2 deletes (no modify) → consensus delete (delete the base row once, authored to the first deleter).
        if (deleters.Count >= 1 && modifiers.Count == 0)
        {
            var first = deleters.OrderBy(e => e.Reviewer).First();
            mergedRowOps.Add(new IrRowOp(IrRowOpKind.DeleteRow, rowAnchor, null, null));
            authoredRows.Add(new IrAuthoredRowOp(IrRowOpKind.DeleteRow, rowAnchor, first.Reviewer, first.Author, null));
            return;
        }

        // 1 ModifyRow → that reviewer's cells, authored (single-source row).
        if (modifiers.Count == 1)
        {
            var (reviewer, author, op) = modifiers[0];
            mergedRowOps.Add(op);
            var authoredCells = AuthoredCellsForSingleReviewer(op, baseRow, reviewer, author);
            authoredRows.Add(new IrAuthoredRowOp(IrRowOpKind.ModifyRow, rowAnchor, reviewer, author, IrNodeList.From(authoredCells)));
            return;
        }

        // ≥2 ModifyRow → PER-CELL compose (the recursion).
        var composedCells = new List<IrAuthoredCellOp>();
        var mergedCellOps = new List<IrCellOp>();
        ComposeRowCells(rowAnchor, baseRow, modifiers, reviewers, baseIr, policy, settings,
            ref nextConflictId, conflicts, mergedCellOps, composedCells);
        mergedRowOps.Add(new IrRowOp(IrRowOpKind.ModifyRow, rowAnchor, rowAnchor, IrNodeList.From(mergedCellOps)));
        authoredRows.Add(new IrAuthoredRowOp(IrRowOpKind.ModifyRow, rowAnchor, -1, "", IrNodeList.From(composedCells)));
    }

    /// <summary>Emit the winning row of a row-level conflict (its merged row op + authored view).</summary>
    private static void EmitWinnerRow(
        (int Reviewer, string Author, IrRowOp Op) winner, string rowAnchor, IrRow baseRow,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers, IrDocument baseIr,
        Docxodus.ConflictResolution policy, IrDiffSettings settings,
        ref int nextConflictId, List<IrConflict> conflicts,
        List<IrRowOp> mergedRowOps, List<IrAuthoredRowOp> authoredRows)
    {
        if (winner.Op.Kind == IrRowOpKind.DeleteRow)
        {
            mergedRowOps.Add(new IrRowOp(IrRowOpKind.DeleteRow, rowAnchor, null, null));
            authoredRows.Add(new IrAuthoredRowOp(IrRowOpKind.DeleteRow, rowAnchor, winner.Reviewer, winner.Author, null));
        }
        else // ModifyRow winner
        {
            mergedRowOps.Add(winner.Op);
            var cells = AuthoredCellsForSingleReviewer(winner.Op, baseRow, winner.Reviewer, winner.Author);
            authoredRows.Add(new IrAuthoredRowOp(IrRowOpKind.ModifyRow, rowAnchor, winner.Reviewer, winner.Author,
                IrNodeList.From(cells)));
        }
    }

    /// <summary>
    /// The authored-cell view of ONE reviewer's ModifyRow: each changed cell (BlockOps non-null) becomes an
    /// <see cref="IrAuthoredCellOp"/> whose composed block ops are that reviewer's BlockOps wrapped as
    /// single-source composite ops (author = the reviewer); an unchanged paired cell (BlockOps null) is a
    /// base passthrough (ComposedBlockOps null); a LEFT-only cell op is the reviewer's column REMOVE
    /// (<see cref="IrAuthoredCellKind.DeleteCell"/>); a RIGHT-only cell op is the reviewer's column ADD
    /// (<see cref="IrAuthoredCellKind.InsertCell"/>). Walks the reviewer's cell ops in THEIR order (the
    /// differ emits every base cell exactly once, paired or left-only, plus inserted cells), keyed by BASE
    /// cell anchor. A changed cell's SHELL (<c>w:tcPr</c>) is sourced from the reviewer's right cell (the
    /// shell digest participates in the cell ContentHash, so a shell-only edit IS a changed cell — sourcing
    /// the shell from the reviewer is what makes a width/merge-only edit land in the composed output instead
    /// of silently reverting to the base shell; for a content-only edit the two shells are canonically
    /// identical, so this is inert).
    /// </summary>
    private static List<IrAuthoredCellOp> AuthoredCellsForSingleReviewer(
        IrRowOp op, IrRow baseRow, int reviewer, string author)
    {
        var result = new List<IrAuthoredCellOp>();
        if (op.CellOps is not { } cellOps)
        {
            // No per-cell op list (a row-property-only modify): every base cell passes through.
            foreach (var cell in baseRow.Cells)
                result.Add(new IrAuthoredCellOp(cell.Anchor.ToString(), null));
            return result;
        }

        foreach (var cellOp in cellOps)
        {
            if (cellOp.LeftCellAnchor is { } baseCellAnchor && cellOp.RightCellAnchor != null)
            {
                // Paired cell: composed content when changed, base passthrough otherwise.
                if (cellOp.BlockOps is { } blockOps)
                {
                    var composed = blockOps.Select(b => EmitOp(b, author, reviewer)).ToList();
                    result.Add(new IrAuthoredCellOp(baseCellAnchor, IrNodeList.From(composed),
                        ShellSourceReviewer: reviewer,
                        ShellRightCellAnchor: cellOp.RightCellAnchor));
                }
                else
                {
                    result.Add(new IrAuthoredCellOp(baseCellAnchor, null));
                }
            }
            else if (cellOp.LeftCellAnchor is { } deletedAnchor)
            {
                // LEFT-only: the reviewer removed this base cell (column remove).
                result.Add(new IrAuthoredCellOp(deletedAnchor, null,
                    ShellSourceReviewer: reviewer, ShellRightCellAnchor: null,
                    Kind: IrAuthoredCellKind.DeleteCell, Author: author));
            }
            else if (cellOp.RightCellAnchor is { } insertedAnchor)
            {
                // RIGHT-only: the reviewer added this cell (column add).
                result.Add(new IrAuthoredCellOp(null, null,
                    ShellSourceReviewer: reviewer, ShellRightCellAnchor: insertedAnchor,
                    Kind: IrAuthoredCellKind.InsertCell, Author: author));
            }
        }
        return result;
    }

    /// <summary>
    /// PER-CELL compose (the recursion): reviewers' cell ops pair by BASE cell ANCHOR (not position, so one
    /// reviewer's column add/remove cannot shift another reviewer's edits). Per base cell: 0 reviewers changed
    /// it → base passthrough; a reviewer's LEFT-only cell op (column REMOVE) → an authored
    /// <see cref="IrAuthoredCellKind.DeleteCell"/> (consensus when no editor contests; a delete-vs-edit cell
    /// is a policy-resolved conflict); 1 editor → that reviewer's BlockOps authored; ≥2 editors → RECURSE
    /// into the body block/token composition over the cell's mini-body. A reviewer's RIGHT-only cell op
    /// (column ADD) slots after the preceding base cell as an authored
    /// <see cref="IrAuthoredCellKind.InsertCell"/> — inserted cells from different reviewers all appear
    /// (block-insert convention). Cell conflicts ascend from <paramref name="nextConflictId"/> with
    /// BaseAnchor = the base cell anchor.
    /// </summary>
    private static void ComposeRowCells(
        string rowAnchor, IrRow baseRow,
        List<(int Reviewer, string Author, IrRowOp Op)> modifiers,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        IrDocument baseIr, Docxodus.ConflictResolution policy, IrDiffSettings settings,
        ref int nextConflictId, List<IrConflict> conflicts,
        List<IrCellOp> mergedCellOps, List<IrAuthoredCellOp> composedCells)
    {
        // Index each modifier's cell ops by BASE cell anchor; RIGHT-only (inserted) cells key by the
        // preceding base cell anchor ("" = before the first cell).
        var byBaseCell = new Dictionary<string, List<(int Reviewer, string Author, IrCellOp CellOp)>>();
        var insertCellsAfter = new Dictionary<string, List<(int Reviewer, string Author, IrCellOp CellOp)>>();
        foreach (var (reviewer, author, rowOp) in modifiers)
        {
            if (rowOp.CellOps is not { } cops)
                continue;
            string preceding = "";
            foreach (var cellOp in cops)
            {
                if (cellOp.LeftCellAnchor is { } la)
                {
                    if (!byBaseCell.TryGetValue(la, out var list)) byBaseCell[la] = list = new();
                    list.Add((reviewer, author, cellOp));
                    preceding = la;
                }
                else if (cellOp.RightCellAnchor != null)
                {
                    if (!insertCellsAfter.TryGetValue(preceding, out var list)) insertCellsAfter[preceding] = list = new();
                    list.Add((reviewer, author, cellOp));
                }
            }
        }

        EmitComposedInsertCells(insertCellsAfter, "", mergedCellOps, composedCells);
        foreach (var baseCell in baseRow.Cells)
        {
            string baseCellAnchor = baseCell.Anchor.ToString();
            ComposeOneBaseCell(baseCell, baseCellAnchor, byBaseCell, reviewers, baseIr, policy, settings,
                ref nextConflictId, conflicts, mergedCellOps, composedCells);
            EmitComposedInsertCells(insertCellsAfter, baseCellAnchor, mergedCellOps, composedCells);
        }
    }

    /// <summary>Emit every reviewer's INSERTED cells slotted after <paramref name="anchor"/> (reviewer order):
    /// the merged right-only cell op plus its authored <see cref="IrAuthoredCellKind.InsertCell"/> view.
    /// Disjoint inserted cells from different reviewers all appear — no insert-vs-insert cell conflict
    /// (block-insert convention).</summary>
    private static void EmitComposedInsertCells(
        IReadOnlyDictionary<string, List<(int Reviewer, string Author, IrCellOp CellOp)>> insertCellsAfter,
        string anchor, List<IrCellOp> mergedCellOps, List<IrAuthoredCellOp> composedCells)
    {
        if (!insertCellsAfter.TryGetValue(anchor, out var list)) return;
        foreach (var (reviewer, author, cellOp) in list)
        {
            mergedCellOps.Add(cellOp);
            composedCells.Add(new IrAuthoredCellOp(null, null,
                ShellSourceReviewer: reviewer, ShellRightCellAnchor: cellOp.RightCellAnchor,
                Kind: IrAuthoredCellKind.InsertCell, Author: author));
        }
    }

    /// <summary>Dispatch one base cell across reviewers (see <see cref="ComposeRowCells"/>).</summary>
    private static void ComposeOneBaseCell(
        IrCell baseCell, string baseCellAnchor,
        IReadOnlyDictionary<string, List<(int Reviewer, string Author, IrCellOp CellOp)>> byBaseCell,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        IrDocument baseIr, Docxodus.ConflictResolution policy, IrDiffSettings settings,
        ref int nextConflictId, List<IrConflict> conflicts,
        List<IrCellOp> mergedCellOps, List<IrAuthoredCellOp> composedCells)
    {
        var entries = byBaseCell.TryGetValue(baseCellAnchor, out var list)
            ? list
            : new List<(int Reviewer, string Author, IrCellOp CellOp)>();

        // A LEFT-only cell op = the reviewer REMOVED this cell (column remove). An editor changed its
        // content (BlockOps non-null). A paired op with null BlockOps is an untouched pairing.
        var deleters = entries.Where(e => e.CellOp.RightCellAnchor == null).ToList();
        var cellEditors = entries.Where(e => e.CellOp.RightCellAnchor != null && e.CellOp.BlockOps != null).ToList();

        if (deleters.Count > 0 && cellEditors.Count > 0)
        {
            // delete-vs-edit CELL conflict (the cell analogue of the row-level conflict).
            int cid = nextConflictId++;
            conflicts.Add(new IrConflict(cid, baseCellAnchor, 0, 0, policy,
                IrNodeList.From(deleters.Concat(cellEditors).OrderBy(e => e.Reviewer).Select(e =>
                    new IrConflictCompetitor(e.Author, CellResultText(baseCell, settings))))));
            if (policy == Docxodus.ConflictResolution.BaseWins)
            {
                mergedCellOps.Add(new IrCellOp(baseCellAnchor, baseCellAnchor, null));
                composedCells.Add(new IrAuthoredCellOp(baseCellAnchor, null));
                return;
            }
            // FirstReviewerWins / StackAll: the lowest-index toucher wins (a cell can be consumed once).
            var winner = deleters.Concat(cellEditors).OrderBy(e => e.Reviewer).First();
            if (winner.CellOp.RightCellAnchor == null)
            {
                EmitDeletedCell(baseCellAnchor, winner, mergedCellOps, composedCells);
            }
            else
            {
                EmitSingleEditorCell(baseCellAnchor, winner, mergedCellOps, composedCells);
            }
            return;
        }

        if (deleters.Count > 0)
        {
            // Consensus delete (1+ deleters, no editor): remove the cell once, authored to the first deleter.
            var first = deleters.OrderBy(e => e.Reviewer).First();
            EmitDeletedCell(baseCellAnchor, first, mergedCellOps, composedCells);
            return;
        }

        if (cellEditors.Count == 0)
        {
            // Base passthrough.
            mergedCellOps.Add(new IrCellOp(baseCellAnchor, baseCellAnchor, null));
            composedCells.Add(new IrAuthoredCellOp(baseCellAnchor, null));
            return;
        }

        if (cellEditors.Count == 1)
        {
            EmitSingleEditorCell(baseCellAnchor, cellEditors[0], mergedCellOps, composedCells);
            return;
        }

        // ≥2 reviewers changed this cell — recurse into block/token composition over the cell mini-body,
        // and compose the cell SHELL (w:tcPr) separately: 0 shell-changers → base shell; all changers
        // agree → that shell (consensus); ≥2 distinct shells → a recorded conflict resolved by policy.
        var (shellReviewer, shellAnchor) = ComposeCellShell(
            baseCell, baseCellAnchor, cellEditors, reviewers, policy, settings, ref nextConflictId, conflicts);
        var cellOps = new List<IrCompositeOp>();
        var cellConflicts = new List<IrConflict>();
        ComposeCellMiniBody(baseCell, cellEditors, reviewers, baseIr, policy, settings,
            ref nextConflictId, cellConflicts, cellOps);
        conflicts.AddRange(cellConflicts);
        mergedCellOps.Add(new IrCellOp(baseCellAnchor, baseCellAnchor,
            IrNodeList.From(cellOps.Select(c => c.Op))));
        composedCells.Add(new IrAuthoredCellOp(baseCellAnchor, IrNodeList.From(cellOps),
            shellReviewer, shellAnchor));
    }

    /// <summary>Emit one reviewer's single-editor cell (merged op verbatim + authored Content view with the
    /// reviewer-sourced shell).</summary>
    private static void EmitSingleEditorCell(
        string baseCellAnchor, (int Reviewer, string Author, IrCellOp CellOp) editor,
        List<IrCellOp> mergedCellOps, List<IrAuthoredCellOp> composedCells)
    {
        mergedCellOps.Add(editor.CellOp);
        var composed = editor.CellOp.BlockOps!.Select(b => EmitOp(b, editor.Author, editor.Reviewer)).ToList();
        composedCells.Add(new IrAuthoredCellOp(baseCellAnchor, IrNodeList.From(composed),
            ShellSourceReviewer: editor.CellOp.RightCellAnchor != null ? editor.Reviewer : -1,
            ShellRightCellAnchor: editor.CellOp.RightCellAnchor));
    }

    /// <summary>Emit one reviewer's deleted base cell (merged left-only op + authored
    /// <see cref="IrAuthoredCellKind.DeleteCell"/> view).</summary>
    private static void EmitDeletedCell(
        string baseCellAnchor, (int Reviewer, string Author, IrCellOp CellOp) deleter,
        List<IrCellOp> mergedCellOps, List<IrAuthoredCellOp> composedCells)
    {
        mergedCellOps.Add(new IrCellOp(baseCellAnchor, null, null));
        composedCells.Add(new IrAuthoredCellOp(baseCellAnchor, null,
            ShellSourceReviewer: deleter.Reviewer, ShellRightCellAnchor: null,
            Kind: IrAuthoredCellKind.DeleteCell, Author: deleter.Author));
    }

    /// <summary>
    /// Compose the SHELL (<c>w:tcPr</c>) of one multi-editor cell across its editing reviewers, via the
    /// <see cref="IrCell.ShellDigest"/> each reviewer's right cell carries. Editors whose right-cell shell
    /// canonically equals the base cell's are not shell-changers (their edit was content-only). 0 changers →
    /// the base shell (-1/null). All changers sharing ONE digest → consensus: the first (lowest reviewer
    /// index) changer's shell. ≥2 distinct digests → a recorded <see cref="IrConflict"/> (anchored at the
    /// base cell) resolved by <paramref name="policy"/>: BaseWins keeps the base shell; FirstReviewerWins and
    /// StackAll take the first changer's shell (shells cannot stack — the disagreement is recorded either
    /// way, so no reviewer's shell edit is ever silently dropped without a conflict).
    /// </summary>
    private static (int ShellReviewer, string? ShellAnchor) ComposeCellShell(
        IrCell baseCell, string baseCellAnchor,
        List<(int Reviewer, string Author, IrCellOp CellOp)> cellEditors,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        Docxodus.ConflictResolution policy, IrDiffSettings settings,
        ref int nextConflictId, List<IrConflict> conflicts)
    {
        var changers = new List<(int Reviewer, string Author, string RightAnchor, IrHash Digest)>();
        foreach (var (reviewer, author, cellOp) in cellEditors)
        {
            if (cellOp.RightCellAnchor is not { } rightAnchor)
                continue;
            var rightCell = FindCell(reviewers[reviewer].Ir, rightAnchor);
            if (rightCell != null && !rightCell.ShellDigest.Equals(baseCell.ShellDigest))
                changers.Add((reviewer, author, rightAnchor, rightCell.ShellDigest));
        }
        if (changers.Count == 0)
            return (-1, null);
        changers.Sort((a, b) => a.Reviewer.CompareTo(b.Reviewer));

        bool allAgree = changers.All(c => c.Digest.Equals(changers[0].Digest));
        if (!allAgree)
        {
            // The competitors name the contested cell by its BASE content (the disagreement — differing
            // tcPr — is structural, recorded by the conflict's existence; mirrors RowResultText's rule).
            conflicts.Add(new IrConflict(nextConflictId++, baseCellAnchor, 0, 0, policy,
                IrNodeList.From(changers.Select(c =>
                    new IrConflictCompetitor(c.Author, CellResultText(baseCell, settings))))));
            if (policy == Docxodus.ConflictResolution.BaseWins)
                return (-1, null);
        }
        return (changers[0].Reviewer, changers[0].RightAnchor);
    }

    /// <summary>Resolve a <c>tc:</c> cell anchor to its <see cref="IrCell"/> in <paramref name="ir"/> (cells
    /// are not in <see cref="IrDocument.AnchorIndex"/>): walk every indexed table's rows/cells, recursing
    /// into nested tables. Null for an unknown anchor. Anchors are unique, so the walk order is immaterial.</summary>
    private static IrCell? FindCell(IrDocument ir, string cellAnchor)
    {
        foreach (var block in ir.AnchorIndex.Values)
            if (block is IrTable tbl && FindCellInTable(tbl, cellAnchor) is { } found)
                return found;
        return null;
    }

    private static IrCell? FindCellInTable(IrTable tbl, string cellAnchor)
    {
        foreach (var row in tbl.Rows)
            foreach (var cell in row.Cells)
            {
                if (cell.Anchor.ToString() == cellAnchor)
                    return cell;
                foreach (var b in cell.Blocks)
                    if (b is IrTable nested && FindCellInTable(nested, cellAnchor) is { } found)
                        return found;
            }
        return null;
    }

    /// <summary>The flat text of one cell (its paragraphs tokenized with the diff tokenizer, nested tables
    /// recursed), for shell-conflict competitor reporting — the cell analogue of <see cref="RowResultText"/>.</summary>
    private static string CellResultText(IrCell cell, IrDiffSettings settings)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var b in cell.Blocks)
        {
            if (b is IrParagraph cp)
                foreach (var t in IrDiffTokenizer.Tokenize(cp, settings))
                    sb.Append(t.Text);
            else if (b is IrTable nested)
                AppendTableText(nested, sb, settings);
        }
        return sb.ToString();
    }

    /// <summary>
    /// Recurse the body merge over ONE cell's paragraph blocks: build a per-reviewer cell script from each
    /// editor's <see cref="IrCellOp.BlockOps"/>, group by base anchor + preceding-anchor inserts exactly as
    /// the body does, and run the shared <see cref="MergeBlockStream"/> dispatch over the base cell's block
    /// anchors. The cell's paragraph anchors live in the SAME global <see cref="IrDocument.AnchorIndex"/>, so
    /// the body's tokenization/format gates apply unchanged. Cell-local conflicts carry the base cell-paragraph
    /// anchor (which resolves in <paramref name="baseIr"/>'s index).
    /// </summary>
    private static void ComposeCellMiniBody(
        IrCell baseCell,
        List<(int Reviewer, string Author, IrCellOp CellOp)> cellEditors,
        IReadOnlyList<(string Author, IrDocument Ir)> reviewers,
        IrDocument baseIr, Docxodus.ConflictResolution policy, IrDiffSettings settings,
        ref int nextConflictId, List<IrConflict> conflicts, List<IrCompositeOp> ops)
    {
        // The base cell's block anchors in document order = the mini-body's base order.
        var baseOrder = baseCell.Blocks.Select(b => b.Anchor.ToString()).ToList();

        // Build one IrEditScript per editing reviewer from its cell BlockOps. Reviewer indices must match the
        // outer reviewers list (so MergeOneBaseBlock's reviewers[e.Reviewer] resolves correctly); a reviewer
        // that did NOT edit this cell contributes an empty script (all base blocks Equal — represented by
        // absence, which MergeOneBaseBlock treats as untouched). We therefore build a FULL-length script list
        // indexed by reviewer, with empty scripts for non-editors.
        var scripts = new List<IrEditScript>(reviewers.Count);
        for (int r = 0; r < reviewers.Count; r++)
            scripts.Add(new IrEditScript(IrNodeList.From(System.Array.Empty<IrEditOp>())));
        foreach (var (reviewer, _, cellOp) in cellEditors)
            scripts[reviewer] = new IrEditScript(cellOp.BlockOps!);

        // A cell mini-body can itself carry Split/Merge ops (a reviewer split/merged paragraphs INSIDE the
        // cell). Run the same sole-toucher plan the body runs: native when uncontested, lowered to del/ins
        // when another editor touches a consumed cell paragraph. Without this, a MergeBlock (null LeftAnchor)
        // reached NEITHER grouping map and the merge was silently dropped from the composed cell.
        var cellTouchers = BuildTouchers(scripts);
        for (int r = 0; r < scripts.Count; r++)
            scripts[r] = ApplySplitMergePlan(scripts[r], r, cellTouchers);

        var byBase = GroupByBaseAnchor(scripts);
        var insertsAfter = GroupInsertsByPrecedingAnchor(scripts);
        var nativeMerges = IndexNativeMerges(scripts);

        MergeBlockStream(baseOrder, byBase, insertsAfter, nativeMerges, reviewers, baseIr, policy, settings,
            ops, conflicts, ref nextConflictId);
    }

    /// <summary>The flat text a row op's competitor carries, for delete-vs-modify ROW conflict reporting. A
    /// row is not an <see cref="IrBlock"/> (so it is not in <see cref="IrDocument.AnchorIndex"/>) and the
    /// merger reads without source elements, so a reviewer's right-row text cannot be resolved by anchor here;
    /// the competitor names the CONTESTED row by its BASE cell text instead (all competitors share the same
    /// base row, so the disagreement — delete vs modify — is recorded by the conflict's existence, not the
    /// text). Mirrors <see cref="ContestedBlockText"/>'s "name the contested block by its left content" rule.</summary>
    private static string RowResultText(IrRow baseRow, IrDiffSettings settings)
    {
        var sb = new System.Text.StringBuilder();
        AppendRowText(baseRow, sb, settings);
        return sb.ToString();
    }

    /// <summary>Append a row's cell text (cell-delimited) to <paramref name="sb"/>.</summary>
    private static void AppendRowText(IrRow row, System.Text.StringBuilder sb, IrDiffSettings settings)
    {
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

    /// <summary>
    /// Debug-only totality guard for a composed table (analogous to <see cref="AssertTilesBase"/>): the
    /// authored rows must tile the base table's rows (every base row consumed exactly once — by an EqualRow,
    /// a ModifyRow, or a DeleteRow — plus any whole InsertRows which carry a null base anchor), and each
    /// composed ModifyRow's cells must tile the base row's cells exactly once.
    /// </summary>
    private static bool AssertTilesBaseTable(List<IrAuthoredRowOp> authoredRows, IrTable baseTable)
    {
        var baseRowAnchors = baseTable.Rows.Select(r => r.Anchor.ToString()).ToList();
        var consumed = new List<string>();
        foreach (var rowOp in authoredRows)
        {
            if (rowOp.Kind == IrRowOpKind.InsertRow)
            {
                if (rowOp.BaseRowAnchor != null) return false; // an inserted row has no base counterpart
                continue;
            }
            if (rowOp.BaseRowAnchor is not { } ra) return false;
            consumed.Add(ra);

            if (rowOp.Kind == IrRowOpKind.ModifyRow && rowOp.ComposedCells is { } cells)
            {
                var baseRow = baseTable.Rows.FirstOrDefault(r => r.Anchor.ToString() == ra);
                if (baseRow == null) return false;
                var baseCells = baseRow.Cells.Select(c => c.Anchor.ToString()).ToList();
                var cellAnchors = cells.Select(c => c.BaseCellAnchor).Where(a => a != null).ToList();
                if (cellAnchors.Count != baseCells.Count) return false;
                for (int i = 0; i < baseCells.Count; i++)
                    if (cellAnchors[i] != baseCells[i]) return false;
            }
        }
        // Every base row consumed exactly once, in order.
        if (consumed.Count != baseRowAnchors.Count) return false;
        for (int i = 0; i < baseRowAnchors.Count; i++)
            if (consumed[i] != baseRowAnchors[i]) return false;
        return true;
    }
}
