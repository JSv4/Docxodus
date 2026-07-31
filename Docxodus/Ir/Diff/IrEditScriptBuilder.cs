#nullable enable

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Builds an <see cref="IrEditScript"/> from two documents (M2.2 Task 2): runs the
/// <see cref="IrBlockAligner"/>, then projects each alignment entry to one or two block-level edit ops,
/// token-diffing Modified paragraph pairs along the way.
/// </summary>
/// <remarks>
/// <para><b>Move source-interleave rule (deterministic, documented, apply-verifier-proven).</b> The
/// aligner emits ONE entry per <see cref="IrAlignmentKind.Moved"/> pair, at the moved block's RIGHT
/// position. The edit script needs TWO ops — a source (delete-from-old-position) and a destination
/// (insert-at-new-position) — so the script reads as a unified diff. We place them thus:</para>
/// <list type="number">
/// <item>The DESTINATION op (<c>IsMoveSource=false</c>, <c>RightAnchor</c> set) is emitted IN PLACE,
/// at the moved entry's position in the aligner's right-ordered entry list — exactly where the aligner
/// put the entry.</item>
/// <item>The SOURCE op (<c>IsMoveSource=true</c>, <c>LeftAnchor</c> set) is interleaved using the SAME
/// left-anchored unified-diff convention the aligner uses for <see cref="IrAlignmentKind.Deleted"/>
/// entries: the source op trails the op of the nearest PAIRED-IN-PLACE left block preceding the moved
/// left block on the LEFT side; sources preceding every such left block go at the very front, in left
/// order. We reconstruct that adjacency from the alignment entries (which carry the left block of every
/// paired entry) plus the left document's block order, so the rule reuses the aligner's published
/// convention rather than duplicating its private interleave helper.</item>
/// </list>
/// <para><b>MoveGroupId allocation.</b> Ascending starting at 1, assigned in DESTINATION order — i.e.
/// the order moved entries appear in the aligner's right-ordered entry list. Deterministic because the
/// entry order is.</para>
/// <para><b>Determinism.</b> Every step is a pure function of the (deterministic) alignment entries and
/// the left block order; no dictionary iteration order is observed for output.</para>
/// </remarks>
internal static class IrEditScriptBuilder
{
    /// <summary>The left side of a move (source), keyed by the moved left block's body index.</summary>
    private readonly record struct MoveInfo(
        int GroupId, IrBlock LeftBlock, IrEditOpKind OpKind, bool RequiresWholeParagraphReplace);

    public static IrEditScript Build(IrDocument left, IrDocument right, IrDiffSettings settings)
    {
        ArgumentNullException.ThrowIfNull(left);
        ArgumentNullException.ThrowIfNull(right);
        ArgumentNullException.ThrowIfNull(settings);

        var alignment = IrBlockAligner.Align(left, right, settings);
        // Cross-paragraph token-stream fusion (IrDiffSettings.CrossParagraphTokenDiff) is enabled ONLY for
        // the top-level BODY projection — the sole scope where it is round-trip-verified. Every OTHER caller
        // of ProjectAlignment (notes, headers/footers, table cells via IrTableDiffer, textbox interiors) uses
        // the default allowCrossParagraph=false, so a word-matched run in those scopes keeps today's exact
        // per-pair behavior, and RenderCrossParagraphRun only ever runs on body blocks.
        var bodyOps = ProjectAlignment(left.Body.Blocks, alignment, settings, allowCrossParagraph: true);
        var noteOps = BuildNoteOps(left, right, settings);
        var headerFooterOps = BuildHeaderFooterOps(left, right, settings);
        return new IrEditScript(IrNodeList.From(bodyOps),
            noteOps.Count == 0 ? null : IrNodeList.From(noteOps),
            headerFooterOps);
    }

    // ------------------------------------------------------------------ header/footer scopes (2026-07-03)

    /// <summary>
    /// Diff the header and footer STORIES of <paramref name="left"/> vs <paramref name="right"/> —
    /// the capability the WmlComparer oracle lacks (it ignores header/footer differences), mirroring
    /// Word Compare's default-on "Headers and footers" granularity. Gated by
    /// <see cref="IrDiffSettings.CompareHeadersFooters"/>.
    ///
    /// <para><b>Pairing (the Word-faithfulness core).</b> Word binds a story to a (section, occurrence
    /// kind) cell, inheriting the previous section's story when a section carries no reference of that
    /// kind. Scope names (<c>hdr1</c>…) are per-document positional labels over part-enumeration order —
    /// NOT stable across two documents — so pairing walks the effective story grid instead: for each
    /// section ordinal and kind, resolve each side's effective story (explicit reference, else inherited
    /// carry-forward), then de-duplicate to distinct (left part, right part) pairs. An inherited story used
    /// by several cells is therefore diffed once. When one LEFT part maps to several different RIGHT stories,
    /// one plan retains the original left part and every additional plan starts from a cloned left-part graph;
    /// explicit per-section bindings select the intended output part and may include a topology-only rejoin
    /// back to the primary part. This preserves an independently reversible story for every output branch.</para>
    ///
    /// <para><b>Ops.</b> A matched pair runs the block aligner over the two stories' block lists exactly
    /// like a note or a cell (a header-text edit surfaces as a ModifyBlock token diff); an all-Equal pair
    /// emits NOTHING (the unchanged story keeps today's verbatim part carry-over); a right-only story is
    /// all InsertBlock; a left-only story all DeleteBlock. Deterministic order: headers (section
    /// ascending, kind Default&lt;First&lt;Even), then footers.</para>
    /// </summary>
    private static IrNodeList<IrHeaderFooterDiff>? BuildHeaderFooterOps(
        IrDocument left, IrDocument right, IrDiffSettings settings)
    {
        if (!settings.CompareHeadersFooters)
            return null;

        var diffs = new List<IrHeaderFooterDiff>();
        BuildOneHeaderFooterScope(left, right, left.Headers, right.Headers, isHeader: true, settings, diffs);
        BuildOneHeaderFooterScope(left, right, left.Footers, right.Footers, isHeader: false, settings, diffs);
        return diffs.Count == 0 ? null : IrNodeList.From(diffs);
    }

    private static readonly IrHeaderFooterKind[] HeaderFooterKinds =
        { IrHeaderFooterKind.Default, IrHeaderFooterKind.First, IrHeaderFooterKind.Even };

    private static void BuildOneHeaderFooterScope(
        IrDocument leftDocument, IrDocument rightDocument,
        IrNodeList<IrHeaderFooter> leftParts, IrNodeList<IrHeaderFooter> rightParts,
        bool isHeader, IrDiffSettings settings, List<IrHeaderFooterDiff> diffs)
    {
        // Pair stories on the effective section × kind grid, but retain the explicit-left grid too: a static
        // rebind must compensate for an explicit left reference that would otherwise override a preceding clone.
        // The body supplies the authoritative section topology; refs only extend that count for malformed input.
        int sectionCount = Math.Max(SectionCount(leftDocument), SectionCount(rightDocument));
        foreach (var hf in leftParts.Concat(rightParts))
            foreach (var reference in hf.References)
                sectionCount = Math.Max(sectionCount, reference.SectionIndex + 1);

        var leftExplicit = ExplicitStoryGrid(leftParts);
        var rightExplicit = ExplicitStoryGrid(rightParts);
        var leftGrid = EffectiveStoryGrid(leftExplicit, sectionCount);
        var rightGrid = EffectiveStoryGrid(rightExplicit, sectionCount);

        // Distinct physical story pairs are rendered once. Unlike the former "first pairing wins" rule, retain
        // every pair: a later pair of the same LEFT part receives a cloned left part, letting its redline reject
        // independently without changing an earlier section that still points at the original.
        var plans = new List<HeaderFooterPairPlan>();
        var plansByPair = new Dictionary<(Uri? Left, Uri? Right), HeaderFooterPairPlan>();
        var plansByCell = new Dictionary<(int Section, IrHeaderFooterKind Kind), HeaderFooterPairPlan>();
        for (int s = 0; s < sectionCount; s++)
        {
            foreach (var kind in HeaderFooterKinds)
            {
                leftGrid.TryGetValue((s, kind), out var leftStory);
                rightGrid.TryGetValue((s, kind), out var rightStory);
                if (leftStory is null && rightStory is null)
                    continue;

                var pair = (leftStory?.Scope.PartUri, rightStory?.Scope.PartUri);
                if (!plansByPair.TryGetValue(pair, out var plan))
                {
                    plan = new HeaderFooterPairPlan(leftStory, rightStory);
                    plansByPair.Add(pair, plan);
                    plans.Add(plan);
                }
                plan.Cells.Add((s, kind));
                plansByCell.Add((s, kind), plan);
            }
        }

        foreach (var plan in plans)
            plan.Ops = BuildHeaderFooterStoryOps(plan.Left, plan.Right, settings);

        // One physical left part can safely be rewritten once. Prefer an unchanged pair as its primary output
        // (preserving the left part byte-for-byte), otherwise use the first effective cell deterministically;
        // every other right target gets a separate clone of that same left source.
        foreach (var group in plans.Where(p => p.LeftPartUri is not null).GroupBy(p => p.LeftPartUri!))
        {
            var primary = group
                .OrderBy(p => p.Ops.Count == 0 ? 0 : 1)
                .ThenBy(p => p.FirstSectionIndex)
                .ThenBy(p => p.FirstKind)
                .First();
            foreach (var plan in group)
                plan.CloneLeftPart = !ReferenceEquals(plan, primary);
        }

        // Plan static references over each kind's output timeline. An explicitly referenced RIGHT cell whose plan
        // actually changes content or output identity gets an exact binding: the body renderer can carry its raw
        // right rId, and a right part may map to different left-derived outputs in different sections. Pure A/A
        // unchanged stories remain absent as before. We also bind at output-part transitions and when an explicit
        // left reference would override an inherited clone even though the desired output part stays the same.
        // This produces the crucial clone→primary rejoin record with empty Ops when needed.
        foreach (var kind in HeaderFooterKinds)
        {
            object? previousOutput = null;
            bool hasPreviousOutput = false;
            for (int s = 0; s < sectionCount; s++)
            {
                if (!plansByCell.TryGetValue((s, kind), out var plan))
                {
                    previousOutput = null;
                    hasPreviousOutput = false;
                    continue;
                }

                object desiredOutput = plan.OutputIdentity;
                leftExplicit.TryGetValue((s, kind), out var explicitLeftStory);
                rightExplicit.TryGetValue((s, kind), out var explicitRightStory);
                bool explicitAlreadySelectsDesired = desiredOutput is Uri desiredLeftUri &&
                    explicitLeftStory?.Scope.PartUri == desiredLeftUri;
                bool explicitWouldOverrideDesired = explicitLeftStory is not null && !explicitAlreadySelectsDesired;
                bool rightExplicitNeedsBinding = explicitRightStory is not null &&
                    (plan.Ops.Count > 0 || plan.CloneLeftPart || plan.LeftPartUri is null);
                bool needsBinding = rightExplicitNeedsBinding || (hasPreviousOutput
                    ? !Equals(previousOutput, desiredOutput) || explicitWouldOverrideDesired
                    : !explicitAlreadySelectsDesired);
                if (needsBinding)
                    plan.ReferenceBindings.Add(new IrHeaderFooterBinding(s, kind));

                previousOutput = desiredOutput;
                hasPreviousOutput = true;
            }
        }

        foreach (var plan in plans.OrderBy(p => p.FirstSectionIndex).ThenBy(p => p.FirstKind))
        {
            // A topology-only record is meaningful: it rebinds a later section back to the original primary part
            // after an earlier cloned story. Purely unchanged/no-topology pairs still remain absent as before.
            if (plan.Ops.Count == 0 && plan.ReferenceBindings.Count == 0)
                continue;

            var (sectionIndex, kind) = plan.Cells[0];
            diffs.Add(new IrHeaderFooterDiff(isHeader, kind, sectionIndex,
                ScopeName: (plan.Right ?? plan.Left)!.ScopeName, LeftScopeName: plan.Left?.ScopeName,
                LeftPartUri: plan.LeftPartUri, RightPartUri: plan.RightPartUri,
                IrNodeList.From(plan.Ops), plan.CloneLeftPart,
                plan.ReferenceBindings.Count == 0 ? null : IrNodeList.From(plan.ReferenceBindings)));
        }
    }

    private sealed class HeaderFooterPairPlan
    {
        public HeaderFooterPairPlan(IrHeaderFooter? left, IrHeaderFooter? right)
        {
            Left = left;
            Right = right;
        }

        public IrHeaderFooter? Left { get; }
        public IrHeaderFooter? Right { get; }
        public Uri? LeftPartUri => Left?.Scope.PartUri;
        public Uri? RightPartUri => Right?.Scope.PartUri;
        public List<(int SectionIndex, IrHeaderFooterKind Kind)> Cells { get; } = new();
        public List<IrEditOp> Ops { get; set; } = new();
        public bool CloneLeftPart { get; set; }
        public List<IrHeaderFooterBinding> ReferenceBindings { get; } = new();
        public int FirstSectionIndex => Cells[0].SectionIndex;
        public IrHeaderFooterKind FirstKind => Cells[0].Kind;

        // Original left parts use their URI as a shared identity. Clones and right-only stories use the plan
        // object itself, so physically distinct output parts never collapse merely because they share a right URI.
        public object OutputIdentity => !CloneLeftPart && LeftPartUri is { } leftUri ? leftUri : this;
    }

    private static int SectionCount(IrDocument document) =>
        Math.Max(1, document.Body.Blocks.OfType<IrParagraph>()
            .Count(paragraph => paragraph.InlineSectionBreakAnchor is not null) + 1);

    private static List<IrEditOp> BuildHeaderFooterStoryOps(
        IrHeaderFooter? left, IrHeaderFooter? right, IrDiffSettings settings)
    {
        if (left is not null && right is not null)
        {
            var alignment = IrBlockAligner.AlignBlocks(left.Scope.Blocks, right.Scope.Blocks, settings);
            var ops = ProjectAlignment(left.Scope.Blocks, alignment, settings);
            return ops.Any(op => op.Kind is not IrEditOpKind.EqualBlock) ? ops : new List<IrEditOp>();
        }
        if (right is not null)
            return right.Scope.Blocks
                .Select(block => new IrEditOp(IrEditOpKind.InsertBlock, null, block.Anchor.ToString(), null, null, null))
                .ToList();
        return left!.Scope.Blocks
            .Select(block => new IrEditOp(IrEditOpKind.DeleteBlock, block.Anchor.ToString(), null, null, null, null))
            .ToList();
    }

    /// <summary>
    /// Resolve one side's effective story per (section ordinal, kind) cell: an explicit body reference
    /// wins (the FIRST reference when a section carries duplicates), else the previous section's
    /// effective story carries forward (Word's inheritance rule); the first section with no reference
    /// has no story. Deterministic — built from the reader's document-ordered reference lists.
    /// </summary>
    private static Dictionary<(int Section, IrHeaderFooterKind Kind), IrHeaderFooter> ExplicitStoryGrid(
        IrNodeList<IrHeaderFooter> parts)
    {
        var explicitCells = new Dictionary<(int, IrHeaderFooterKind), IrHeaderFooter>();
        foreach (var hf in parts)
            foreach (var r in hf.References)
                if (!explicitCells.ContainsKey((r.SectionIndex, r.Kind)))
                    explicitCells[(r.SectionIndex, r.Kind)] = hf;

        return explicitCells;
    }

    private static Dictionary<(int Section, IrHeaderFooterKind Kind), IrHeaderFooter> EffectiveStoryGrid(
        IReadOnlyDictionary<(int Section, IrHeaderFooterKind Kind), IrHeaderFooter> explicitCells,
        int sectionCount)
    {
        var grid = new Dictionary<(int, IrHeaderFooterKind), IrHeaderFooter>();
        foreach (var kind in HeaderFooterKinds)
        {
            IrHeaderFooter? current = null;
            for (int s = 0; s < sectionCount; s++)
            {
                if (explicitCells.TryGetValue((s, kind), out var hf))
                    current = hf;
                if (current is not null)
                    grid[(s, kind)] = current;
            }
        }
        return grid;
    }

    // ------------------------------------------------------------------ note scopes (M2.4 Task 1)

    /// <summary>
    /// Diff the footnote and endnote stores of <paramref name="left"/> vs <paramref name="right"/>, in the
    /// DETERMINISTIC document order <see cref="IrEditScript.NoteOps"/> documents: footnotes (by note id,
    /// numeric ascending) then endnotes (by note id, numeric ascending). For each note id present in either
    /// store: a matched note aligns its left/right block lists with the body block aligner and projects the
    /// alignment to block ops (so a footnote-text edit surfaces as a ModifyBlock token diff inside the note,
    /// exactly like a body paragraph); an only-left note becomes all-Deleted blocks; an only-right note
    /// all-Inserted blocks. Mirrors <see cref="WmlComparer.GetRevisions"/>'s footnote+endnote coverage —
    /// header/footer scopes are deliberately NOT diffed (the oracle does not diff them either).
    /// </summary>
    private static List<IrNoteDiff> BuildNoteOps(IrDocument left, IrDocument right, IrDiffSettings settings)
    {
        var result = new List<IrNoteDiff>();
        result.AddRange(BuildOneStore(left, right, IrNoteKind.Footnote, settings));
        result.AddRange(BuildOneStore(left, right, IrNoteKind.Endnote, settings));
        return result;
    }

    /// <summary>
    /// Diff one note store (footnotes OR endnotes) under the oracle's NOTE CORRESPONDENCE semantics.
    ///
    /// <para><b>Why not by raw <c>w:id</c>.</b> <see cref="WmlComparer"/> does NOT pair notes by their stored
    /// <c>w:id</c>. <c>WmlComparer.ChangeFootnoteEndnoteReferencesToUniqueRange</c> RENUMBERS every note id to a
    /// per-document range in BODY-REFERENCE ORDER (the n-th <c>w:footnoteReference</c>/<c>w:endnoteReference</c>
    /// encountered walking <c>document.xml</c> gets id <c>base+n</c>, and its note definition is renumbered to
    /// match); then <c>WmlComparer.ProcessFootnoteEndnote</c> pairs a note with another note IFF their body
    /// REFERENCES correlate Equal in the body diff. So a note's correspondence is driven by its reference's
    /// position in the body, never by the original id. When a reference is inserted (WC034-After3 relocates an
    /// endnote ref INTO the middle of <c>Video</c>, shifting the ids), by-id pairing cross-matches unrelated
    /// notes and over-reports; reference-order/content pairing matches the oracle.</para>
    ///
    /// <para><b>The IR equivalent.</b> Collect each side's referenced note ids in body document order
    /// (<see cref="CollectNoteReferenceOrder"/>), then align the two reference sequences
    /// (<see cref="AlignNoteReferences"/>): exact-content references anchor an order-preserving spine first, the
    /// residue pairs leftover references by best note-content similarity (the lone-left/lone-right case pairs
    /// UNCONDITIONALLY — the same 1×1-residue rule the block aligner uses, so a single edited note on each side
    /// is always a content modify, never a delete+insert), and surplus references fall out as whole-note
    /// insert/delete. <b>Invariant:</b> when the reference sequences have equal length and pair up in order
    /// (no inserted/deleted reference shifted them — the overwhelmingly common case), this reduces EXACTLY to
    /// the former by-id pairing, so unrelated fixtures (WC-1600/1660/1750/…) are byte-identical.</para>
    ///
    /// <para>A matched pair aligns its note blocks (a footnote-text edit surfaces as a ModifyBlock token diff
    /// inside the note, like a body paragraph); an only-left reference becomes all-Deleted blocks, an only-right
    /// all-Inserted. The per-scope op stream is ordered by RIGHT reference order (inserts interleaved at their
    /// reference position), then deleted-only notes, then any unreferenced-on-both-sides notes — deterministic.</para>
    /// </summary>
    private static List<IrNoteDiff> BuildOneStore(
        IrDocument left, IrDocument right, IrNoteKind kind, IrDiffSettings settings)
    {
        var leftStore = kind == IrNoteKind.Footnote ? left.Footnotes : left.Endnotes;
        var rightStore = kind == IrNoteKind.Footnote ? right.Footnotes : right.Endnotes;

        // Referenced note ids in body document order (the oracle's correspondence axis). Defensively de-dup:
        // a note referenced twice corresponds once (first reference wins) — the oracle renumbers per reference
        // but compares each note's content once.
        var refsLeft = DistinctInOrder(CollectNoteReferenceOrder(left, kind), leftStore);
        var refsRight = DistinctInOrder(CollectNoteReferenceOrder(right, kind), rightStore);

        var correspondence = AlignNoteReferences(refsLeft, refsRight, leftStore, rightStore, settings);

        // Notes that exist in a store but are NOT referenced from the body (orphans — uncommon but legal).
        // Pair common ids, otherwise treat as whole-note ins/del; appended after the referenced stream in
        // numeric-id order for determinism.
        AppendUnreferencedNotes(correspondence, refsLeft, refsRight, leftStore, rightStore);

        var diffs = new List<IrNoteDiff>();
        foreach (var (leftId, rightId) in correspondence)
        {
            bool hasLeft = leftId is not null && leftStore.Notes.TryGetValue(leftId, out _);
            bool hasRight = rightId is not null && rightStore.Notes.TryGetValue(rightId, out _);

            // The note id used to scope the diff: the right id for matched/inserted notes (the produced
            // document is right-shaped), the left id for a deleted-only note.
            string scopeId = rightId ?? leftId!;

            List<IrEditOp> ops;
            if (hasLeft && hasRight)
            {
                var leftScope = leftStore.Notes[leftId!];
                var rightScope = rightStore.Notes[rightId!];
                var alignment = IrBlockAligner.AlignBlocks(leftScope.Blocks, rightScope.Blocks, settings);
                ops = ProjectAlignment(leftScope.Blocks, alignment, settings);
            }
            else if (hasRight)
            {
                ops = rightStore.Notes[rightId!].Blocks
                    .Select(b => new IrEditOp(IrEditOpKind.InsertBlock, null, b.Anchor.ToString(), null, null, null))
                    .ToList();
            }
            else
            {
                ops = leftStore.Notes[leftId!].Blocks
                    .Select(b => new IrEditOp(IrEditOpKind.DeleteBlock, b.Anchor.ToString(), null, null, null, null))
                    .ToList();
            }

            // A matched note whose alignment is entirely EqualBlock/FormatOnly carries no real change; normally we
            // emit no note diff so an unedited note produces zero revisions. EXCEPTION: a matched note whose left
            // and right ids DIFFER (an inserted reference shifted the numbering) still needs an entry so the markup
            // renderer can reconcile the produced definition's id to the right id space — without it the renumber
            // pass cannot link the equal body reference (right id) to its definition (still at the left id), and
            // accept/reject would dangle. The all-EqualBlock ops project to NO revisions (EqualBlock is a no-op in
            // IrRevisionRenderer), so the revisions surface and its counts are unchanged; only the markup renderer
            // acts on the entry (re-id + content passthrough).
            bool hasRealChange = ops.Any(o => o.Kind is not IrEditOpKind.EqualBlock);
            bool idShifted = hasLeft && hasRight && leftId != rightId;
            if (hasRealChange || idShifted)
                diffs.Add(new IrNoteDiff(kind, scopeId, IrNodeList.From(ops), hasLeft ? leftId : null));
        }
        return diffs;
    }

    // ------------------------------------------------------------------ note correspondence (M2.5 Task 3)

    /// <summary>Walk the body in document order (recursing into block SDTs, tables, textboxes, fields, and hyperlinks the
    /// same way the reader/tokenizer do) and return the <see cref="IrNoteRef.NoteId"/> of every reference of
    /// <paramref name="kind"/>, in encounter order — the oracle's note-correspondence axis.</summary>
    private static List<string> CollectNoteReferenceOrder(IrDocument doc, IrNoteKind kind)
    {
        var ids = new List<string>();
        WalkBlocksForNoteRefs(doc.Body.Blocks, kind, ids);
        return ids;
    }

    private static void WalkBlocksForNoteRefs(IReadOnlyList<IrBlock> blocks, IrNoteKind kind, List<string> sink)
    {
        foreach (var block in blocks)
        {
            switch (block)
            {
                case IrParagraph p:
                    WalkInlinesForNoteRefs(p.Inlines, kind, sink);
                    break;
                case IrTable t:
                    foreach (var row in t.Rows)
                        foreach (var cell in row.Cells)
                            WalkBlocksForNoteRefs(cell.Blocks, kind, sink);
                    break;
                case IrSdtBlock sdt:
                    WalkBlocksForNoteRefs(sdt.Blocks, kind, sink);
                    break;
            }
        }
    }

    private static void WalkInlinesForNoteRefs(IReadOnlyList<IrInline> inlines, IrNoteKind kind, List<string> sink)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case IrNoteRef note when note.Kind == kind:
                    sink.Add(note.NoteId);
                    break;
                case IrFieldRun field:
                    WalkInlinesForNoteRefs(field.CachedResult, kind, sink);
                    break;
                case IrHyperlink link:
                    WalkInlinesForNoteRefs(link.Inlines, kind, sink);
                    break;
                case IrTextbox tbx:
                    WalkBlocksForNoteRefs(tbx.Blocks, kind, sink);
                    break;
            }
        }
    }

    /// <summary>Keep the first occurrence of each id (a note referenced twice corresponds once) and drop ids
    /// with no note definition in the store (a dangling reference contributes no note diff).</summary>
    private static List<string> DistinctInOrder(List<string> ids, IrNoteStore store)
    {
        var seen = new HashSet<string>();
        var result = new List<string>(ids.Count);
        foreach (var id in ids)
            if (store.Notes.ContainsKey(id) && seen.Add(id))
                result.Add(id);
        return result;
    }

    /// <summary>
    /// Align two body-reference-ordered note-id sequences into an ordered list of <c>(leftId?, rightId?)</c>
    /// correspondence pairs, mirroring the oracle's body-reference correlation.
    ///
    /// <para><b>Pass 1 — exact-content spine.</b> References whose notes are <see cref="IrBlock.ContentHash"/>-equal
    /// are matched along the longest order-preserving (LCS) subsequence, so an unchanged note pairs with its
    /// reference-order counterpart even when other references were inserted around it.</para>
    /// <para><b>Pass 2 — similarity residue.</b> Leftover left references pair with leftover right references
    /// (order-respecting) by best note similarity, highest score first; a lone-left/lone-right residue pairs
    /// UNCONDITIONALLY (the 1×1-residue rule — a single edited note on each side is one modify, not del+ins).</para>
    /// <para><b>Surplus.</b> Unpaired left references → <c>(id, null)</c> (whole-note delete); unpaired right →
    /// <c>(null, id)</c> (whole-note insert). The result is ordered by right reference position with deletes
    /// interleaved before the right reference they precede, so the op stream reads as a unified note diff.</para>
    /// </summary>
    private static List<(string? Left, string? Right)> AlignNoteReferences(
        List<string> refsLeft, List<string> refsRight,
        IrNoteStore leftStore, IrNoteStore rightStore, IrDiffSettings settings)
    {
        int nLeft = refsLeft.Count;
        int nRight = refsRight.Count;
        var leftPartner = new int[nLeft];  // right index this left ref paired with, or -1
        var rightPartner = new int[nRight]; // left index this right ref paired with, or -1
        Array.Fill(leftPartner, -1);
        Array.Fill(rightPartner, -1);

        // Pass 1: exact-content LCS spine.
        bool ContentEqual(int li, int rj) =>
            NoteContentEqual(leftStore.Notes[refsLeft[li]], rightStore.Notes[refsRight[rj]]);
        var spine = LongestCommonSubsequence(nLeft, nRight, ContentEqual);
        foreach (var (li, rj) in spine)
        {
            leftPartner[li] = rj;
            rightPartner[rj] = li;
        }

        // Pass 2: similarity residue over the still-free references, preserving order. Greedy highest-score
        // first; the lone-left/lone-right pairing falls out as the single remaining candidate (forced pair).
        // Score on the note's WHOLE-content token multiset (Jaccard) so a multi-paragraph note is not penalized
        // by per-block averaging — the most content-overlapping note wins, matching the oracle's body-LCS choice.
        var bagCache = new Dictionary<string, Dictionary<string, int>>();
        Dictionary<string, int> LeftBag(string id) => NoteTokenBag(bagCache, "L:" + id, leftStore.Notes[id], settings);
        Dictionary<string, int> RightBag(string id) => NoteTokenBag(bagCache, "R:" + id, rightStore.Notes[id], settings);
        double Score(int li, int rj) => BagJaccard(LeftBag(refsLeft[li]), RightBag(refsRight[rj]));
        GreedyResiduePair(nLeft, nRight, leftPartner, rightPartner, Score);

        // Emit in right order with deletes interleaved before the right ref they precede (by left order).
        var result = new List<(string?, string?)>();
        int nextLeft = 0;
        for (int rj = 0; rj < nRight; rj++)
        {
            int li = rightPartner[rj];
            if (li >= 0)
            {
                // Flush any unpaired left refs that precede this matched left ref (deletes in left order).
                while (nextLeft < li)
                {
                    if (leftPartner[nextLeft] < 0)
                        result.Add((refsLeft[nextLeft], null));
                    nextLeft++;
                }
                nextLeft = li + 1;
                result.Add((refsLeft[li], refsRight[rj]));
            }
            else
            {
                result.Add((null, refsRight[rj]));
            }
        }
        // Trailing unpaired left refs (deletes after the last matched left ref).
        for (; nextLeft < nLeft; nextLeft++)
            if (leftPartner[nextLeft] < 0)
                result.Add((refsLeft[nextLeft], null));

        return result;
    }

    /// <summary>Longest common subsequence of <c>[0,nLeft)</c> × <c>[0,nRight)</c> under the boolean
    /// <paramref name="match"/> predicate, returned as ascending (leftIndex, rightIndex) pairs. Standard
    /// O(nLeft·nRight) DP — note-reference counts are tiny, so cost is negligible.</summary>
    private static List<(int Left, int Right)> LongestCommonSubsequence(
        int nLeft, int nRight, Func<int, int, bool> match)
    {
        var dp = new int[nLeft + 1, nRight + 1];
        for (int i = nLeft - 1; i >= 0; i--)
            for (int j = nRight - 1; j >= 0; j--)
                dp[i, j] = match(i, j)
                    ? dp[i + 1, j + 1] + 1
                    : Math.Max(dp[i + 1, j], dp[i, j + 1]);

        var pairs = new List<(int, int)>();
        for (int i = 0, j = 0; i < nLeft && j < nRight;)
        {
            if (match(i, j)) { pairs.Add((i, j)); i++; j++; }
            else if (dp[i + 1, j] >= dp[i, j + 1]) i++;
            else j++;
        }
        return pairs;
    }

    /// <summary>Greedily pair still-free left/right references (those with partner == -1) by descending
    /// <paramref name="score"/>, preserving order monotonicity: a chosen pair must not cross an already-chosen
    /// pair. Ties break by smallest left then right index. The lone-left/lone-right case yields the single
    /// candidate, pairing it unconditionally regardless of score.</summary>
    private static void GreedyResiduePair(
        int nLeft, int nRight, int[] leftPartner, int[] rightPartner, Func<int, int, double> score)
    {
        while (true)
        {
            double best = double.NegativeInfinity;
            int bestLi = -1, bestRj = -1;
            for (int li = 0; li < nLeft; li++)
            {
                if (leftPartner[li] >= 0) continue;
                for (int rj = 0; rj < nRight; rj++)
                {
                    if (rightPartner[rj] >= 0) continue;
                    if (CrossesExistingPair(li, rj, leftPartner)) continue;
                    double s = score(li, rj);
                    if (s > best) { best = s; bestLi = li; bestRj = rj; }
                }
            }
            if (bestLi < 0) return; // no order-compatible free pair remains
            leftPartner[bestLi] = bestRj;
            rightPartner[bestRj] = bestLi;
        }
    }

    /// <summary>True if pairing left ref <paramref name="li"/> with right ref <paramref name="rj"/> would cross
    /// an already-established pair (order violation): some matched left ref &lt; li paired with a right ref &gt; rj,
    /// or some matched left ref &gt; li paired with a right ref &lt; rj.</summary>
    private static bool CrossesExistingPair(int li, int rj, int[] leftPartner)
    {
        for (int k = 0; k < leftPartner.Length; k++)
        {
            int p = leftPartner[k];
            if (p < 0) continue;
            if ((k < li && p > rj) || (k > li && p < rj)) return true;
        }
        return false;
    }

    /// <summary>Two note scopes are "exact-content equal" iff they have the same number of blocks and every
    /// block's <see cref="IrBlock.ContentHash"/> matches in order (a structural content digest — the same key
    /// the block aligner anchors on). Used only to seed the LCS spine of unchanged references.</summary>
    private static bool NoteContentEqual(IrScope a, IrScope b)
    {
        if (a.Blocks.Count != b.Blocks.Count) return false;
        for (int i = 0; i < a.Blocks.Count; i++)
            if (!a.Blocks[i].ContentHash.Equals(b.Blocks[i].ContentHash)) return false;
        return true;
    }

    /// <summary>Build (and cache) a note scope's whole-content WORD multiset — the <see cref="IrDiffTokenKind.Word"/>
    /// tokens of direct paragraph blocks and paragraphs nested beneath block SDTs, concatenated in document order,
    /// keyed by <see cref="IrDiffToken.MatchKey"/>.
    /// Whitespace, separators, note-ref/opaque markers and other non-word tokens are EXCLUDED: they are shared
    /// near-uniformly across notes (every note opens with the same note-ref marker and is mostly spaces) and
    /// would otherwise dominate a Jaccard score, pairing notes by length rather than by content. Block SDTs
    /// contribute their descendant paragraphs; other non-paragraph blocks contribute nothing. The residue
    /// similarity is a coarse content-overlap heuristic — it disambiguates
    /// which reference a leftover note corresponds to, never gating a forced lone-left/lone-right pair.</summary>
    private static Dictionary<string, int> NoteTokenBag(
        Dictionary<string, Dictionary<string, int>> cache, string cacheKey, IrScope scope, IrDiffSettings settings)
    {
        if (cache.TryGetValue(cacheKey, out var bag)) return bag;
        bag = new Dictionary<string, int>();
        AddNoteTokens(scope.Blocks, bag, settings);
        cache[cacheKey] = bag;
        return bag;
    }

    /// <summary>Add word tokens from direct paragraph blocks and paragraph descendants of block SDTs in a note.
    /// This is deliberately a bookkeeping projection only: it does not make a block SDT token-diffable or alter
    /// block alignment, which continues to treat its envelope as atomic.</summary>
    private static void AddNoteTokens(
        IReadOnlyList<IrBlock> blocks, Dictionary<string, int> bag, IrDiffSettings settings)
    {
        foreach (var block in blocks)
        {
            switch (block)
            {
                case IrParagraph p:
                    foreach (var t in IrDiffTokenizer.Tokenize(p, settings))
                        if (t.Kind == IrDiffTokenKind.Word)
                            bag[t.MatchKey] = bag.TryGetValue(t.MatchKey, out int c) ? c + 1 : 1;
                    break;
                case IrSdtBlock sdt:
                    AddNoteTokens(sdt.Blocks, bag, settings);
                    break;
            }
        }
    }

    /// <summary>Jaccard index over two token multisets (sum of per-key min counts / sum of per-key max counts).
    /// Two empty bags score 1.0; an empty-vs-nonempty pair scores 0. Used only to disambiguate the residue.</summary>
    private static double BagJaccard(Dictionary<string, int> a, Dictionary<string, int> b)
    {
        int totalA = a.Values.Sum();
        int totalB = b.Values.Sum();
        if (totalA == 0 && totalB == 0) return 1.0;
        if (totalA == 0 || totalB == 0) return 0.0;

        int intersection = 0;
        var (small, large) = a.Count <= b.Count ? (a, b) : (b, a);
        foreach (var kv in small)
            if (large.TryGetValue(kv.Key, out int other))
                intersection += Math.Min(kv.Value, other);
        int union = totalA + totalB - intersection;
        return union == 0 ? 1.0 : (double)intersection / union;
    }

    /// <summary>Append correspondence entries for notes present in a store but NOT referenced from the body
    /// (orphans). Common ids pair (defensive — keeps a previously-by-id-matched orphan note matched); the rest
    /// become whole-note ins/del. Ordered numeric-id-ascending for determinism.</summary>
    private static void AppendUnreferencedNotes(
        List<(string? Left, string? Right)> correspondence,
        List<string> refsLeft, List<string> refsRight,
        IrNoteStore leftStore, IrNoteStore rightStore)
    {
        var referencedLeft = new HashSet<string>(refsLeft);
        var referencedRight = new HashSet<string>(refsRight);
        var orphanLeft = leftStore.Notes.Keys.Where(id => !referencedLeft.Contains(id));
        var orphanRight = rightStore.Notes.Keys.Where(id => !referencedRight.Contains(id));
        var ids = new SortedSet<string>(orphanLeft.Concat(orphanRight), NoteIdComparer.Instance);
        foreach (var id in ids)
        {
            bool l = leftStore.Notes.ContainsKey(id) && !referencedLeft.Contains(id);
            bool r = rightStore.Notes.ContainsKey(id) && !referencedRight.Contains(id);

            // A footnote's identity is its w:id, so each id must correspond EXACTLY ONCE. When a note's
            // REFERENCE was deleted on one side but its DEFINITION still exists on the other (the def lingers,
            // unreferenced), the reference-driven AlignNoteReferences pass already emitted a one-sided surplus
            // entry for the side that kept its reference: a kept-left reference → (id, null); a kept-right
            // reference → (null, id). The lingering definition is the SAME note, so reconcile that surplus into a
            // matched (id, id) pair instead of appending a SECOND entry — appending here would emit a second
            // w:footnote definition with the same id (one inserted, one deleted): the duplicate-id corruption.
            // A matched pair re-diffs the (usually unchanged) definition: a no-op when content is equal, a content
            // diff otherwise — never a duplicate.
            if (r && !l)
            {
                int idx = correspondence.FindIndex(c => c.Left == id && c.Right == null);
                if (idx >= 0) { correspondence[idx] = (id, id); continue; }
            }
            else if (l && !r)
            {
                int idx = correspondence.FindIndex(c => c.Right == id && c.Left == null);
                if (idx >= 0) { correspondence[idx] = (id, id); continue; }
            }

            if (l && r) correspondence.Add((id, id));
            else if (r) correspondence.Add((null, id));
            else if (l) correspondence.Add((id, null));
        }
    }

    /// <summary>Numeric-ascending note-id order (id is a <c>w:id</c> integer string); non-numeric ids sort
    /// after all numeric ids by ordinal string, so the order is total and deterministic.</summary>
    private sealed class NoteIdComparer : IComparer<string>
    {
        public static readonly NoteIdComparer Instance = new();

        public int Compare(string? x, string? y)
        {
            bool xn = int.TryParse(x, NumberStyles.Integer, CultureInfo.InvariantCulture, out int xi);
            bool yn = int.TryParse(y, NumberStyles.Integer, CultureInfo.InvariantCulture, out int yi);
            if (xn && yn) return xi.CompareTo(yi);
            if (xn) return -1;
            if (yn) return 1;
            return string.CompareOrdinal(x, y);
        }
    }

    /// <summary>
    /// Project an alignment over <paramref name="leftBlocks"/> into the ordered block edit-op list
    /// (right order, with move/delete sources interleaved). Shared by <see cref="Build"/> (body) and
    /// <see cref="IrTableDiffer"/> (cell block lists) so both produce identical op shapes. Move group ids
    /// are LOCAL to this projection (1..N in destination order) — for cell projections that means ids are
    /// scoped to the cell, which is exactly right since a row/cell move never crosses cells in M2.2.
    /// </summary>
    public static List<IrEditOp> ProjectAlignment(
        IrNodeList<IrBlock> leftBlocks, IrBlockAlignment alignment, IrDiffSettings settings,
        bool allowCrossParagraph = false)
    {
        // Left block index by reference identity → used to order move-source interleaving by left position.
        var leftIndex = BuildLeftIndexMap(leftBlocks);

        // Pass 1: assign MoveGroupIds in destination (right-entry) order, ascending from 1, capturing
        // each move's source block + the op kind (MoveBlock vs MoveModifyBlock), keyed by left index.
        var moves = new Dictionary<int, MoveInfo>(); // left-block index → move info
        int nextGroup = 1;
        foreach (var entry in alignment.Entries)
        {
            if (entry.Kind is IrAlignmentKind.Moved or IrAlignmentKind.MovedModified)
            {
                int li = leftIndex[entry.Left!];
                var opKind = entry.Kind == IrAlignmentKind.MovedModified
                    ? IrEditOpKind.MoveModifyBlock
                    : IrEditOpKind.MoveBlock;
                moves[li] = new MoveInfo(nextGroup++, entry.Left!, opKind,
                    StructuralCarrierDiffers(entry.Left!, entry.Right!) ||
                    (entry.Kind == IrAlignmentKind.MovedModified &&
                     entry.Left is IrParagraph movedLeft && entry.Right is IrParagraph movedRight &&
                     HasUnsliceableFieldCarrier(movedLeft, movedRight)));
            }
        }

        // Bucket move-source ops by the left index of the nearest preceding paired-in-place left block
        // (left-anchored convention; -1 = front), walking the LEFT document order.
        var sourcesAfterLeft = BuildSourceInterleave(leftBlocks, alignment, leftIndex, moves);

        var ops = new List<IrEditOp>();

        // Move-source ops are STAGED when their anchor's entry is emitted and flushed lazily so they
        // interleave with the anchor's trailing DELETED entries by LEFT index — the op stream then
        // reads in left-document order (a deleted block at left index 7 precedes a moved-away block
        // at left index 8 even though both trail the same in-place anchor; M2.6 fuzz seed-16
        // reject-order regression).
        var pendingSources = new List<int>();
        StageSources(sourcesAfterLeft, -1, pendingSources); // sources preceding every in-place anchor

        var alignmentEntries = alignment.Entries;

        // Story-final pair detection for the weak-pair flush (decoded 2026-07-27): the Modified pair
        // whose left block is the story's LAST left block and whose entry carries the story's LAST
        // right content is the tail-fusion construct of the per-pair path. Word flushes it to a whole
        // del+ins when its only retention is a lone short word (see ApplyStoryFinalWeakPairFlush);
        // interior pairs never flush (Word retains even a lone shared function word mid-story).
        // Trailing section-break blocks are TRANSPARENT here, exactly as in the renderer's story-end
        // test (the trailing section-break op renders nothing).
        int lastRightEntryIndex = -1;
        for (int i = 0; i < alignmentEntries.Count; i++)
        {
            var e = alignmentEntries[i];
            bool hasRight = (e.Right is not null && e.Right is not IrSectionBreak) ||
                (e.Kind == IrAlignmentKind.Split && e.MultiBlocks is { Count: > 0 });
            if (hasRight)
                lastRightEntryIndex = i;
        }
        IrBlock? lastLeftBlock = null;
        for (int i = leftBlocks.Count - 1; i >= 0; i--)
        {
            if (leftBlocks[i] is not IrSectionBreak)
            {
                lastLeftBlock = leftBlocks[i];
                break;
            }
        }
        for (int entryIndex = 0; entryIndex < alignmentEntries.Count; entryIndex++)
        {
            var entry = alignmentEntries[entryIndex];
            if (pendingSources.Count > 0)
            {
                // A Deleted entry releases only the staged sources whose left index PRECEDES it; any
                // other entry ends the anchor's deletion run and releases everything staged.
                int limit = entry.Kind == IrAlignmentKind.Deleted ? leftIndex[entry.Left!] : int.MaxValue;
                FlushPendingSources(pendingSources, limit, moves, ops);
            }

            // Cross-paragraph token-stream fusion (markup-only, gated by IrDiffSettings.CrossParagraphTokenDiff;
            // default off ⇒ byte-identical). At a Modified paragraph entry, try to fuse the maximal streak of
            // ADJACENT streamable Modified pairs starting here into ONE CrossParagraphRunBlock op — the flat
            // word+pilcrow stream decoded from Word's compare output over word-matched regions. Runs of <2
            // pairs, non-streamable members, and segmenter bails all fall straight through to the ordinary
            // per-entry emission below. (Every pending move source was flushed above — a Modified entry flushes
            // everything — and the streak collector stops before any member with sources staged UNDER it, so
            // the source-op interleave order is untouched.)
            if (allowCrossParagraph && settings.CrossParagraphTokenDiff &&
                entry.Kind is IrAlignmentKind.Modified or IrAlignmentKind.Inserted or IrAlignmentKind.Deleted)
            {
                // At a Modified entry, the word-matched pair-run collector goes first (its behavior is
                // pinned); the story-final MIXED-region collector (2026-07-27) then covers the decoded
                // shapes the run path cannot own — zero-pair replace regions whose one-sided paragraphs
                // share words, and regions where one-sided members PRECEDE the word-matched pair.
                int consumed = 0;
                List<IrEditOp>? trailingSectionOps = null;
                var fused = entry.Kind == IrAlignmentKind.Modified
                    ? TryBuildCrossParagraphRunOp(
                        alignmentEntries, entryIndex, settings, sourcesAfterLeft, leftIndex, out consumed)
                    : null;
                if (fused is null && pendingSources.Count == 0)
                    fused = TryBuildStoryFinalMixedRegionOp(
                        alignmentEntries, entryIndex, settings, sourcesAfterLeft, leftIndex,
                        out consumed, out trailingSectionOps);
                if (fused is not null)
                {
                    ops.Add(fused);
                    if (trailingSectionOps is not null)
                        ops.AddRange(trailingSectionOps);
                    // Stage the move-sources anchored under every consumed member exactly as the per-entry
                    // path would (paired-in-place lefts, plus a consumed Merge entry's member lefts; only
                    // the LAST consumed pair may carry any — the streak collector enforces it, and tail
                    // absorption is refused outright when it does — so the staged list stays sorted
                    // exactly as the per-entry path would leave it).
                    for (int k = 0; k < consumed; k++)
                    {
                        var consumedEntry = alignmentEntries[entryIndex + k];
                        if (consumedEntry.Left is not null && IsPairedInPlace(consumedEntry.Kind))
                            StageSources(sourcesAfterLeft, leftIndex[consumedEntry.Left], pendingSources);
                        if (consumedEntry.Kind == IrAlignmentKind.Merge && consumedEntry.MultiBlocks is { } consumedLefts)
                            foreach (var lb in consumedLefts)
                                StageSources(sourcesAfterLeft, leftIndex[lb], pendingSources);
                    }
                    entryIndex += consumed - 1; // -1: the for-loop's ++ advances past the last consumed entry.
                    continue;
                }
            }

            switch (entry.Kind)
            {
                case IrAlignmentKind.Unchanged:
                    if (StructuralCarrierDiffers(entry.Left!, entry.Right!))
                        ops.Add(MakeParagraphModifyOp((IrParagraph)entry.Left!, (IrParagraph)entry.Right!, settings));
                    else
                        ops.Add(new IrEditOp(IrEditOpKind.EqualBlock,
                            entry.Left!.Anchor.ToString(), entry.Right!.Anchor.ToString(),
                            null, null, null));
                    break;

                case IrAlignmentKind.FormatOnly:
                    if (StructuralCarrierDiffers(entry.Left!, entry.Right!))
                        ops.Add(MakeParagraphModifyOp((IrParagraph)entry.Left!, (IrParagraph)entry.Right!, settings));
                    else
                        ops.Add(new IrEditOp(IrEditOpKind.FormatOnlyBlock,
                            entry.Left!.Anchor.ToString(), entry.Right!.Anchor.ToString(),
                            null, null, null));
                    break;

                case IrAlignmentKind.Modified:
                    // A story-final pair directly preceded by unpaired gap content (an Inserted/Deleted
                    // entry) is part of a LARGER tail construct in the reference output — its window
                    // carries the gap's matter too, and the reference retains the pair's lone word
                    // there — so the flush applies only to a cleanly-paired tail.
                    ops.Add(MakeModifyOp(entry.Left!, entry.Right!, settings,
                        storyFinalPair: entryIndex == lastRightEntryIndex &&
                                        ReferenceEquals(entry.Left, lastLeftBlock) &&
                                        (entryIndex == 0 ||
                                         alignmentEntries[entryIndex - 1].Kind is not
                                             (IrAlignmentKind.Inserted or IrAlignmentKind.Deleted))));
                    break;

                case IrAlignmentKind.Inserted:
                    ops.Add(new IrEditOp(IrEditOpKind.InsertBlock,
                        null, entry.Right!.Anchor.ToString(), null, null, null,
                        BodyFullRewriteGroupId: entry.BodyFullRewriteGroupId));
                    break;

                case IrAlignmentKind.Deleted:
                    ops.Add(new IrEditOp(IrEditOpKind.DeleteBlock,
                        entry.Left!.Anchor.ToString(), null, null, null, null,
                        BodyFullRewriteGroupId: entry.BodyFullRewriteGroupId));
                    break;

                // The detection gate (IrBlockAligner split/merge scan) only ever groups PARAGRAPH
                // blocks into a Split/Merge entry, so the IrParagraph casts below are safe.
                case IrAlignmentKind.Split:
                {
                    var lp = (IrParagraph)entry.Left!;
                    var members = entry.MultiBlocks!.Cast<IrParagraph>().ToList();
                    if (HasStructuralCarrier(lp) || members.Any(HasStructuralCarrier))
                    {
                        // A split has no one-to-one paragraph shell in which to place an atomic inline carrier
                        // or a non-tokenizable field carrier. Lower it to the existing whole-block pair.
                        ops.Add(new IrEditOp(IrEditOpKind.DeleteBlock,
                            lp.Anchor.ToString(), null, null, null, null));
                        foreach (var member in members)
                            ops.Add(new IrEditOp(IrEditOpKind.InsertBlock,
                                null, member.Anchor.ToString(), null, null, null));
                        break;
                    }
                    ops.Add(new IrEditOp(IrEditOpKind.SplitBlock,
                        lp.Anchor.ToString(), null, null, null, null, null, null,
                        IrNodeList.From(members.Select(m => m.Anchor.ToString()).ToList()),
                        IrSplitSegmenter.ComputeSegmentDiffs(lp, members, settings, singularIsLeft: true)));
                    break;
                }

                case IrAlignmentKind.Merge:
                {
                    var rp = (IrParagraph)entry.Right!;
                    var members = entry.MultiBlocks!.Cast<IrParagraph>().ToList();
                    if (HasStructuralCarrier(rp) || members.Any(HasStructuralCarrier))
                    {
                        foreach (var member in members)
                            ops.Add(new IrEditOp(IrEditOpKind.DeleteBlock,
                                member.Anchor.ToString(), null, null, null, null));
                        ops.Add(new IrEditOp(IrEditOpKind.InsertBlock,
                            null, rp.Anchor.ToString(), null, null, null));
                        break;
                    }
                    // Segment diffs are computed singular-vs-members (rp sliced against each left
                    // member) then MIRRORED so each stored diff reads left-member → right-slice,
                    // keeping the universal "left = left document" orientation for every consumer.
                    var sliced = IrSplitSegmenter.ComputeSegmentDiffs(rp, members, settings, singularIsLeft: false);
                    ops.Add(new IrEditOp(IrEditOpKind.MergeBlock,
                        null, rp.Anchor.ToString(), null, null, null, null, null,
                        IrNodeList.From(members.Select(m => m.Anchor.ToString()).ToList()),
                        IrNodeList.From(sliced.Select(IrSplitSegmenter.MirrorDiff).ToList())));
                    break;
                }

                case IrAlignmentKind.Moved:
                case IrAlignmentKind.MovedModified:
                {
                    // Emit the DESTINATION op in place; the SOURCE op was interleaved separately.
                    var move = moves[leftIndex[entry.Left!]];
                    // MoveModifyBlock (from a MovedModified alignment, M2.2 Task 3) carries the in-move
                    // token diff on its destination — tokenize source (left) vs destination (right) so the
                    // op describes "relocated AND edited"; a plain Moved destination carries none.
                    var tokenDiff = move.OpKind == IrEditOpKind.MoveModifyBlock
                        ? TokenDiffFor(entry.Left!, entry.Right!, settings)
                        : null;
                    ops.Add(new IrEditOp(
                        move.OpKind, null, entry.Right!.Anchor.ToString(),
                        tokenDiff, move.GroupId, IsMoveSource: false,
                        RequiresWholeParagraphReplace: move.RequiresWholeParagraphReplace));
                    break;
                }
            }

            // After a paired-in-place left block's entry, STAGE the move-sources anchored to it (they
            // flush lazily, interleaved with the anchor's trailing Deleted entries — see above). A
            // Split entry's singular entry.Left is covered here (IsPairedInPlace includes Split).
            if (entry.Left is not null && IsPairedInPlace(entry.Kind))
                StageSources(sourcesAfterLeft, leftIndex[entry.Left], pendingSources);

            // A Merge entry carries its paired-in-place lefts in MultiBlocks (entry.Left is null):
            // stage move-sources anchored to each member, in ascending left order (MultiBlocks is in
            // left order, and per-anchor buckets are disjoint ascending runs, so appending keeps the
            // staged list sorted) — mirroring the aligner's deletion-flush convention.
            if (entry.Kind == IrAlignmentKind.Merge && entry.MultiBlocks is { } mergeLefts)
                foreach (var lb in mergeLefts)
                    StageSources(sourcesAfterLeft, leftIndex[lb], pendingSources);
        }

        // Trailing staged sources (an anchor at the very end of the document).
        FlushPendingSources(pendingSources, int.MaxValue, moves, ops);

        return ops;
    }

    // --------------------------------------------------- cross-paragraph token-stream run (2026-07-25)

    /// <summary>
    /// Try to fuse the maximal streak of ADJACENT word-matched Modified paragraph pairs starting at
    /// <paramref name="start"/> — plus, when the streak's following entries are a STORY-FINAL replace gap
    /// (Inserted/Deleted streamable paragraphs) or ONE story-final split/merge group, that trailing tail
    /// as one-sided stream members — into ONE <see cref="IrEditOpKind.CrossParagraphRunBlock"/> op.
    /// Returns null (and leaves <paramref name="consumed"/> 0) — so the caller emits the ordinary
    /// per-entry ops — when fewer than 2 eligible pairs are adjacent and no tail is absorbable, or when
    /// the segmenter declines (structure gate / cell-invariant guard). A pair is eligible iff both sides
    /// are <see cref="IrCrossParagraphSegmenter.IsStreamable"/> paragraphs with no structural carrier
    /// (<see cref="HasStructuralCarrier"/> — the same digest gate split/merge uses; an inline
    /// SDT/smartTag envelope or field carrier cannot be sliced across cells). The streak stops AFTER any
    /// member with move-sources staged under its left block — and such a member also blocks tail
    /// absorption — so the flat path's source-op interleave position is preserved exactly. When the TAIL
    /// stream declines, the pure pair run is retried exactly as the tail-less path would have (the
    /// ordinary replace-gap / split / merge grammar then renders the tail). On success
    /// <paramref name="consumed"/> is the total ENTRY count the op replaces (pairs + tail entries).
    /// </summary>
    private static IrEditOp? TryBuildCrossParagraphRunOp(
        IrNodeList<IrAlignedBlock> entries, int start, IrDiffSettings settings,
        Dictionary<int, List<int>> sourcesAfterLeft, Dictionary<IrBlock, int> leftIndex, out int consumed)
    {
        consumed = 0;

        var leftParas = new List<IrParagraph>();
        var rightParas = new List<IrParagraph>();
        bool lastPairHasSources = false;
        for (int i = start; i < entries.Count && entries[i].Kind == IrAlignmentKind.Modified; i++)
        {
            if (entries[i].Left is not IrParagraph lp || entries[i].Right is not IrParagraph rp ||
                !IrCrossParagraphSegmenter.IsStreamable(lp) || !IrCrossParagraphSegmenter.IsStreamable(rp) ||
                HasStructuralCarrier(lp) || HasStructuralCarrier(rp))
                break;
            leftParas.Add(lp);
            rightParas.Add(rp);
            // Move-sources staged under this member's left block flush AFTER the member's op on the flat
            // path. Fusing further members would displace them, so this member closes the streak.
            if (sourcesAfterLeft.ContainsKey(leftIndex[lp]))
            {
                lastPairHasSources = true;
                break;
            }
        }
        int pairCount = leftParas.Count;
        if (pairCount == 0)
            return null;

        // Absorb a story-final tail (replace gap / split / merge surplus) as one-sided stream members —
        // unless the last pair carries staged move-sources (the flat path interleaves those with the
        // tail's ops by left order; absorbing would displace them).
        int pairedCount = pairCount;
        int tailEntries = lastPairHasSources
            ? 0
            : TryCollectStoryFinalTail(entries, start + pairCount, leftParas, rightParas, ref pairedCount);
        bool hasTail = leftParas.Count != pairedCount || rightParas.Count != pairedCount;
        if (!hasTail && pairCount < 2)
            return null; // the flat-stream shape is decoded for RUNS of ≥2 word-matched pairs (a lone pair streams only with a tail).

        // Story-final run: nothing follows the consumed entries except section-break metadata. With a
        // tail this is TRUE by construction (the collector demanded it); the segmenter then keeps the
        // last pair's in-slot anchors (the fusion construct is the tail itself). Without one, the
        // segmenter treats the run's LAST pair as the tail-fusion construct (no in-slot anchors — its
        // retained tokens arrive only by forward cross-flow), per the decoded grammar.
        bool runEndsStory = OnlySectionBreakEntriesFrom(entries, start + pairCount + tailEntries);

        var cells = IrCrossParagraphSegmenter.Segment(leftParas, rightParas, settings, runEndsStory, pairedCount);
        if (cells is null && hasTail)
        {
            // The tail stream declined (zero crossings, or a bail): retry the PURE pair run exactly as
            // the tail-less path would have; the ordinary grammar renders the tail entries.
            if (pairCount < 2)
                return null;
            leftParas.RemoveRange(pairCount, leftParas.Count - pairCount);
            rightParas.RemoveRange(pairCount, rightParas.Count - pairCount);
            tailEntries = 0;
            runEndsStory = OnlySectionBreakEntriesFrom(entries, start + pairCount);
            cells = IrCrossParagraphSegmenter.Segment(leftParas, rightParas, settings, runEndsStory, pairCount);
        }
        if (cells is null)
            return null;

        consumed = pairCount + tailEntries;
        return new IrEditOp(IrEditOpKind.CrossParagraphRunBlock,
            leftParas[0].Anchor.ToString(), rightParas[0].Anchor.ToString(),
            TokenDiff: null, MoveGroupId: null, IsMoveSource: null,
            CrossParagraphCells: IrNodeList.From(cells));
    }

    /// <summary>True iff every entry from <paramref name="from"/> on carries only section-break
    /// metadata — the shared story-end test of the cross-paragraph constructs.</summary>
    private static bool OnlySectionBreakEntriesFrom(IrNodeList<IrAlignedBlock> entries, int from)
    {
        for (int j = from; j < entries.Count; j++)
        {
            if (entries[j].Left is IrSectionBreak || entries[j].Right is IrSectionBreak)
                continue;
            return false;
        }
        return true;
    }

    /// <summary>
    /// Collect the story-final TAIL following a word-matched run, appending its paragraphs to the
    /// member lists as one-sided stream members (decoded run+gap construct, 2026-07-27). Two shapes:
    /// (a) a contiguous run of Inserted/Deleted entries — a replace gap — whose deleted paragraphs
    /// join the LEFT members and inserted ones the RIGHT members, in entry (= document) order; or
    /// (b) exactly ONE Split/Merge entry — the aligner's story-tail surplus (or a genuine story-final
    /// split/merge) — whose singular↔first-member becomes the run's LAST pair
    /// (<paramref name="pairedCount"/> grows by one) and whose surplus members become one-sided.
    /// Either shape must reach the story end (only section-break entries after) and every touched
    /// paragraph must be streamable with no structural carrier; anything else collects NOTHING (the
    /// member lists are untouched and 0 is returned, keeping the tail-less behavior). Returns the
    /// number of ENTRIES consumed.
    /// </summary>
    private static int TryCollectStoryFinalTail(
        IrNodeList<IrAlignedBlock> entries, int tailStart,
        List<IrParagraph> leftParas, List<IrParagraph> rightParas, ref int pairedCount)
    {
        static bool Eligible(IrBlock? block, out IrParagraph para)
        {
            para = null!;
            if (block is not IrParagraph p || !IrCrossParagraphSegmenter.IsStreamable(p) || HasStructuralCarrier(p))
                return false;
            para = p;
            return true;
        }

        if (tailStart >= entries.Count)
            return 0;

        var kind = entries[tailStart].Kind;
        if (kind is IrAlignmentKind.Inserted or IrAlignmentKind.Deleted)
        {
            var dels = new List<IrParagraph>();
            var inss = new List<IrParagraph>();
            int j = tailStart;
            for (; j < entries.Count &&
                   entries[j].Kind is IrAlignmentKind.Inserted or IrAlignmentKind.Deleted; j++)
            {
                if (entries[j].Kind == IrAlignmentKind.Deleted)
                {
                    if (!Eligible(entries[j].Left, out var dp))
                        return 0;
                    dels.Add(dp);
                }
                else
                {
                    if (!Eligible(entries[j].Right, out var ip))
                        return 0;
                    inss.Add(ip);
                }
            }
            if (!OnlySectionBreakEntriesFrom(entries, j))
                return 0;
            leftParas.AddRange(dels);
            rightParas.AddRange(inss);
            return j - tailStart;
        }

        if (kind is IrAlignmentKind.Split or IrAlignmentKind.Merge)
        {
            var entry = entries[tailStart];
            if (!OnlySectionBreakEntriesFrom(entries, tailStart + 1) ||
                entry.MultiBlocks is not { Count: >= 2 } members)
                return 0;
            var memberParas = new List<IrParagraph>(members.Count);
            foreach (var m in members)
            {
                if (!Eligible(m, out var mp))
                    return 0;
                memberParas.Add(mp);
            }
            if (kind == IrAlignmentKind.Split)
            {
                if (!Eligible(entry.Left, out var singular))
                    return 0;
                leftParas.Add(singular);
                rightParas.Add(memberParas[0]);
                pairedCount++;
                for (int m = 1; m < memberParas.Count; m++)
                    rightParas.Add(memberParas[m]);
            }
            else
            {
                if (!Eligible(entry.Right, out var singular))
                    return 0;
                leftParas.Add(memberParas[0]);
                rightParas.Add(singular);
                pairedCount++;
                for (int m = 1; m < memberParas.Count; m++)
                    leftParas.Add(memberParas[m]);
            }
            return 1;
        }

        return 0;
    }

    /// <summary>
    /// Story-final MIXED-region collector (2026-07-27): the decoded gap-stream shapes the pair-run
    /// collector cannot own. Collects the maximal run of streamable Inserted/Deleted/Modified
    /// paragraph entries from <paramref name="start"/> that reaches the story end, builds the general
    /// member lists + pair links, and segments them via
    /// <see cref="IrCrossParagraphSegmenter.SegmentRegion"/>. Fires ONLY for configurations outside
    /// the pair-run path's contract: a ZERO-pair region (a story-final replace gap whose one-sided
    /// paragraphs share words — reference compare output streams function words through such gaps),
    /// or a region with a one-sided member BEFORE its last pair (a leading deleted/inserted paragraph
    /// joining the pair's stream). Pure pair-runs and pairs-then-trailing-gap shapes return null here
    /// — they belong to <see cref="TryBuildCrossParagraphRunOp"/>. A 1×1 gap never enters (the
    /// replace-gap grammar owns it); body-full-rewrite entries, structural carriers, non-streamable
    /// paragraphs, and any member with staged move-sources all decline. The segmenter's structure
    /// gates then require a genuine structural change (unpaired interior pilcrow / crossing match),
    /// so a matchless region falls back to the ordinary per-entry grammar unchanged.
    /// </summary>
    private static IrEditOp? TryBuildStoryFinalMixedRegionOp(
        IrNodeList<IrAlignedBlock> entries, int start, IrDiffSettings settings,
        Dictionary<int, List<int>> sourcesAfterLeft, Dictionary<IrBlock, int> leftIndex, out int consumed,
        out List<IrEditOp>? trailingSectionOps)
    {
        consumed = 0;
        trailingSectionOps = null;

        var leftParas = new List<IrParagraph>();
        var rightParas = new List<IrParagraph>();
        var pairs = new List<(int L, int R)>();
        var sectionEntries = new List<IrAlignedBlock>();
        int j = start;
        for (; j < entries.Count; j++)
        {
            var e = entries[j];
            // One-sided SECTION-BREAK entries are story-end metadata, not gap matter: when the two
            // sides' section properties differ, the aligner emits a Deleted + an Inserted section
            // break INTERLEAVED with the story's final replace region (the deleted one after the
            // left paragraphs, the inserted one after the right). The stream looks straight through
            // them — they are collected transparently and their ops re-emitted after the fused op,
            // exactly as the per-entry path would render them (a section break op emits no body
            // paragraph in place).
            if (e.Kind == IrAlignmentKind.Deleted && e.Left is IrSectionBreak)
            {
                sectionEntries.Add(e);
                continue;
            }
            if (e.Kind == IrAlignmentKind.Inserted && e.Right is IrSectionBreak)
            {
                sectionEntries.Add(e);
                continue;
            }
            if (e.Kind == IrAlignmentKind.Deleted)
            {
                // A BodyFullRewriteGroupId-marked del/ins residue still joins the region: with no
                // stream match the gates decline and the full-rewrite arrangement keeps it; with a
                // genuine cross-flow match the reference output streams it (a leading rewrite's
                // shared word lands retained inside the following pair's paragraph).
                if (e.Left is not IrParagraph dp || !IrCrossParagraphSegmenter.IsStreamable(dp) ||
                    HasStructuralCarrier(dp))
                    break;
                leftParas.Add(dp);
            }
            else if (e.Kind == IrAlignmentKind.Inserted)
            {
                if (e.Right is not IrParagraph ip || !IrCrossParagraphSegmenter.IsStreamable(ip) ||
                    HasStructuralCarrier(ip))
                    break;
                rightParas.Add(ip);
            }
            else if (e.Kind == IrAlignmentKind.Modified)
            {
                if (e.Left is not IrParagraph lp || e.Right is not IrParagraph rp ||
                    !IrCrossParagraphSegmenter.IsStreamable(lp) || !IrCrossParagraphSegmenter.IsStreamable(rp) ||
                    HasStructuralCarrier(lp) || HasStructuralCarrier(rp))
                    break;
                pairs.Add((leftParas.Count, rightParas.Count));
                leftParas.Add(lp);
                rightParas.Add(rp);
            }
            else
            {
                break;
            }
        }
        if (j <= start)
            return null;
        if (leftParas.Count == 0 || rightParas.Count == 0)
            return null;
        // MAXIMALITY: the stream owns whole regions, never fragments. If more one-sided gap matter
        // directly precedes the start (an upstream non-streamable/table break) or the collection
        // broke ON a one-sided entry (a mid-gap table or non-streamable paragraph), this run is a
        // slice of a larger replace region — the replace-gap grammar arranges those as a whole.
        // Section-break entries are transparent to this test on both sides (story-end metadata).
        static bool IsOneSidedGapMatter(IrAlignedBlock e) =>
            e.Kind is IrAlignmentKind.Inserted or IrAlignmentKind.Deleted &&
            e.Left is not IrSectionBreak && e.Right is not IrSectionBreak;
        int prev = start - 1;
        while (prev >= 0 && !IsOneSidedGapMatter(entries[prev]) &&
               entries[prev].Kind is IrAlignmentKind.Inserted or IrAlignmentKind.Deleted)
            prev--;
        if (prev >= 0 && IsOneSidedGapMatter(entries[prev]))
            return null;
        if (j < entries.Count && IsOneSidedGapMatter(entries[j]))
            return null;
        bool runEndsStory = OnlySectionBreakEntriesFrom(entries, j);

        // Move-source discipline: the region op would displace the flat path's source/deletion
        // interleave, so any member with sources staged under its left block declines outright.
        foreach (var lp in leftParas)
            if (sourcesAfterLeft.ContainsKey(leftIndex[lp]))
                return null;

        if (pairs.Count == 0)
        {
            // The 1×1 story-final del+ins belongs to the replace-gap grammar (its fused-tail re-diff
            // already retains shared tokens); the stream only owns multi-member regions. INTERIOR
            // zero-pair regions are admitted too — the segmenter ships them only on a count-equal
            // boundary construct. LARGE regions never stream: a whole-document rewrite is one giant
            // zero-pair gap whose scattered incidental shared words would otherwise ship a stream,
            // while the reference compare output arranges such regions with the replace-gap grammar
            // (the decoded stream constructs all live in small regions — the grammar corpus was
            // validated on exactly the large ones).
            if (leftParas.Count + rightParas.Count < 3)
                return null;
            if (leftParas.Count > 8 || rightParas.Count > 8)
                return null;
        }
        else if (runEndsStory)
        {
            // Only shapes the pair-run collector cannot own: ≥1 one-sided member BEFORE the last
            // pair (the run path already handles pairs + a trailing story-final tail).
            var (lastL, lastR) = pairs[^1];
            bool leadingOneSided = false;
            for (int i = 0; i < lastL && !leadingOneSided; i++)
                leadingOneSided = !pairs.Exists(pr => pr.L == i);
            for (int i = 0; i < lastR && !leadingOneSided; i++)
                leadingOneSided = !pairs.Exists(pr => pr.R == i);
            if (!leadingOneSided)
                return null;
        }
        else
        {
            // INTERIOR region with pairs: decoded (the mid-document construct) only when a
            // one-sided DELETED member accompanies the pair(s) — the construct's signature is the
            // pair's leading ins-residue hoisted into a preceding ¶DEL cell. Insert-only
            // surroundings and pure interior pair runs stay with the per-entry grammar.
            if (leftParas.Count == pairs.Count)
                return null;
        }

        var cells = IrCrossParagraphSegmenter.SegmentRegion(
            leftParas, rightParas, pairs, settings, runEndsStory);
        if (cells is null)
            return null;

        // Transparently-collected section-break entries re-emit their per-entry ops AFTER the fused
        // op, in entry order — a section break op renders no body paragraph in place, so its
        // position among the region's paragraph content is immaterial; only its presence is.
        if (sectionEntries.Count > 0)
        {
            trailingSectionOps = new List<IrEditOp>(sectionEntries.Count);
            foreach (var se in sectionEntries)
            {
                trailingSectionOps.Add(se.Kind == IrAlignmentKind.Deleted
                    ? new IrEditOp(IrEditOpKind.DeleteBlock,
                        se.Left!.Anchor.ToString(), null, null, null, null,
                        BodyFullRewriteGroupId: se.BodyFullRewriteGroupId)
                    : new IrEditOp(IrEditOpKind.InsertBlock,
                        null, se.Right!.Anchor.ToString(), null, null, null,
                        BodyFullRewriteGroupId: se.BodyFullRewriteGroupId));
            }
        }

        consumed = j - start;
        return new IrEditOp(IrEditOpKind.CrossParagraphRunBlock,
            leftParas[0].Anchor.ToString(), rightParas[0].Anchor.ToString(),
            TokenDiff: null, MoveGroupId: null, IsMoveSource: null,
            CrossParagraphCells: IrNodeList.From(cells));
    }

    /// <summary>Append the move-source left indexes bucketed under <paramref name="anchorLeftIndex"/> to
    /// the staged list (ascending order preserved — buckets are disjoint ascending runs staged in
    /// anchor order).</summary>
    private static void StageSources(
        Dictionary<int, List<int>> sourcesAfterLeft, int anchorLeftIndex, List<int> pendingSources)
    {
        if (sourcesAfterLeft.TryGetValue(anchorLeftIndex, out var list))
            pendingSources.AddRange(list);
    }

    /// <summary>Emit (and remove) the staged move-source ops whose left index is &lt; <paramref name="limit"/>,
    /// in ascending left order — the lazy half of the source/deletion left-order interleave.</summary>
    private static void FlushPendingSources(
        List<int> pendingSources, int limit, Dictionary<int, MoveInfo> moves, List<IrEditOp> ops)
    {
        int n = 0;
        while (n < pendingSources.Count && pendingSources[n] < limit)
        {
            var move = moves[pendingSources[n]];
            // The source op mirrors the destination's kind; the token diff lives only on the destination.
            ops.Add(new IrEditOp(
                move.OpKind, move.LeftBlock.Anchor.ToString(), null, null, move.GroupId, IsMoveSource: true,
                RequiresWholeParagraphReplace: move.RequiresWholeParagraphReplace));
            n++;
        }
        pendingSources.RemoveRange(0, n);
    }

    // ------------------------------------------------------------------ modify op (token / table diff)

    /// <summary>
    /// Build a <see cref="IrEditOpKind.ModifyBlock"/> op for a Modified pair. A paragraph pair carries a
    /// token diff; a TABLE pair carries a nested <see cref="IrTableDiff"/> (M2.2 Task 4) — so a cell-text
    /// edit surfaces as a token diff inside the cell, not a whole-table blob; any other non-paragraph
    /// pair (opaque / section break) carries neither.
    /// </summary>
    private static IrEditOp MakeModifyOp(
        IrBlock leftBlock, IrBlock rightBlock, IrDiffSettings settings, bool storyFinalPair = false)
    {
        if (leftBlock is IrTable lt && rightBlock is IrTable rt)
            return new IrEditOp(IrEditOpKind.ModifyBlock,
                leftBlock.Anchor.ToString(), rightBlock.Anchor.ToString(),
                null, null, null, IrTableDiffer.Diff(lt, rt, settings));

        if (leftBlock is IrParagraph lp && rightBlock is IrParagraph rp)
            return MakeParagraphModifyOp(lp, rp, settings, storyFinalPair);

        return new IrEditOp(IrEditOpKind.ModifyBlock,
            leftBlock.Anchor.ToString(), rightBlock.Anchor.ToString(),
            TokenDiffFor(leftBlock, rightBlock, settings), null, null);
    }

    /// <summary>
    /// Build the ModifyBlock for a Modified PARAGRAPH pair (M2.4 Task 1: textbox interiors). When both
    /// paragraphs carry textboxes whose placeholder tokens differ, recurse: pair the textboxes positionally
    /// within the paragraph, align each pair's inner blocks, and attach the nested ops as
    /// <see cref="IrEditOp.TextboxDiffs"/> (mirroring the table-diff nesting). The paragraph's OWN token
    /// diff then EXCLUDES the placeholder-token change (the differ keys on a MASKED token list whose textbox
    /// placeholders share one constant key, so they pair as Equal) — the textbox change is reported once,
    /// through the nested ops, never also as an opaque-placeholder token op.
    /// </summary>
    private static IrEditOp MakeParagraphModifyOp(
        IrParagraph lp, IrParagraph rp, IrDiffSettings settings, bool storyFinalPair = false)
    {
        var leftBoxes = CollectTextboxes(lp.Inlines);
        var rightBoxes = CollectTextboxes(rp.Inlines);

        // Build the textbox diffs (positional pairing + surplus insert/delete). Only keep them when at least
        // one carries a real change; if every box is unchanged we leave the paragraph as a plain token diff.
        var textboxDiffs = BuildTextboxDiffs(leftBoxes, rightBoxes, settings);
        bool nest = textboxDiffs is not null;

        // When nesting, mask the placeholder tokens so the paragraph token diff does not also report the
        // textbox change; otherwise tokenize normally.
        var leftTokens = IrDiffTokenizer.Tokenize(lp, settings);
        var rightTokens = IrDiffTokenizer.Tokenize(rp, settings);
        var diffLeft = nest ? MaskTextboxKeys(leftTokens) : leftTokens;
        var diffRight = nest ? MaskTextboxKeys(rightTokens) : rightTokens;
        var tokenDiff = IrTokenDiffer.Diff(diffLeft, diffRight, settings,
            suppressLonePunctuation: PlainTextPair(lp, rp));

        if (storyFinalPair && !nest && PlainTextPair(lp, rp))
            tokenDiff = ApplyStoryFinalWeakPairFlush(tokenDiff, diffLeft, diffRight);

        return new IrEditOp(IrEditOpKind.ModifyBlock,
            lp.Anchor.ToString(), rp.Anchor.ToString(),
            tokenDiff, null, null, null,
            nest ? IrNodeList.From(textboxDiffs!) : null,
            RequiresWholeParagraphReplace: StructuralCarrierDiffers(lp, rp) ||
                HasUnsliceableFieldCarrier(lp, rp));
    }

    /// <summary>
    /// Inline SDT/smartTag wrappers and non-hyperlink field code/state are deliberately transparent to ordinary
    /// token alignment, but either carrier can be corrupted by slicing its source around a revision. Keep the
    /// alignment stable and mark a differing paired paragraph for a reversible whole-paragraph replacement.
    /// </summary>
    private static bool StructuralCarrierDiffers(IrBlock left, IrBlock right) =>
        left is IrParagraph lp && right is IrParagraph rp &&
        (lp.InlineEnvelopeDigest != rp.InlineEnvelopeDigest ||
         lp.FieldEnvelopeDigest != rp.FieldEnvelopeDigest);

    private static bool HasStructuralCarrier(IrParagraph paragraph) =>
        paragraph.InlineEnvelopeDigest != default ||
        paragraph.FieldEnvelopeDigest != default ||
        ContainsFieldCarrier(paragraph.Inlines);

    /// <summary>
    /// A field with no visible text has no character span in the token stream. Fine source slicing cannot assign
    /// it deterministically when adjacent prose is replaced (there may be no Equal token to own the carrier), so
    /// a modified paired paragraph containing one is rendered as one reversible old/new paragraph pair. Direct
    /// simple fields with a tab/image/textbox-only result are included: all are zero-width to the tokenizer.
    /// Textbox-contained fields are included too because the owning drawing is itself zero-width in the outer
    /// paragraph's source coordinates. A canonicalized HYPERLINK field is included regardless of result width:
    /// its field origin is equality-neutral in the IR, but its raw simple/complex plumbing cannot be assigned to
    /// a token slice without risking a bare right-side field surviving Reject.
    /// </summary>
    private static bool HasUnsliceableFieldCarrier(IrParagraph left, IrParagraph right) =>
        HasUnsliceableFieldCarrier(left.Inlines) || HasUnsliceableFieldCarrier(right.Inlines);

    private static bool HasUnsliceableFieldCarrier(IReadOnlyList<IrInline> inlines)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case IrFieldRun field:
                    if (VisibleTextLength(field.CachedResult) == 0 || HasUnsliceableFieldCarrier(field.CachedResult))
                        return true;
                    break;
                case IrHyperlink hyperlink:
                    if (hyperlink.IsFieldHyperlink || HasUnsliceableFieldCarrier(hyperlink.Inlines))
                        return true;
                    break;
                case IrTextbox textbox when ContainsFieldCarrier(textbox.Blocks):
                    return true;
            }
        }
        return false;
    }

    private static int VisibleTextLength(IReadOnlyList<IrInline> inlines)
    {
        int length = 0;
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case IrTextRun text:
                    length += text.Text.Length;
                    break;
                case IrFieldRun field:
                    length += VisibleTextLength(field.CachedResult);
                    break;
                case IrHyperlink hyperlink:
                    length += VisibleTextLength(hyperlink.Inlines);
                    break;
            }
        }
        return length;
    }

    private static bool ContainsFieldCarrier(IReadOnlyList<IrBlock> blocks)
    {
        foreach (var block in blocks)
        {
            switch (block)
            {
                case IrParagraph paragraph when ContainsFieldCarrier(paragraph.Inlines):
                    return true;
                case IrTable table:
                    foreach (var row in table.Rows)
                        foreach (var cell in row.Cells)
                            if (ContainsFieldCarrier(cell.Blocks))
                                return true;
                    break;
                case IrSdtBlock sdt when ContainsFieldCarrier(sdt.Blocks):
                    return true;
            }
        }
        return false;
    }

    private static bool ContainsFieldCarrier(IReadOnlyList<IrInline> inlines)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case IrFieldRun:
                    return true;
                case IrHyperlink hyperlink:
                    if (hyperlink.IsFieldHyperlink || ContainsFieldCarrier(hyperlink.Inlines))
                        return true;
                    break;
                case IrTextbox textbox when ContainsFieldCarrier(textbox.Blocks):
                    return true;
            }
        }
        return false;
    }

    // ------------------------------------------------------------------ textbox interiors (M2.4 Task 1)

    /// <summary>
    /// Collect the <see cref="IrTextbox"/> inlines of a paragraph in DOCUMENT ORDER, recursing transparently
    /// through fields' cached results and hyperlinks exactly as <see cref="IrDiffTokenizer"/> does — so the
    /// i-th collected textbox corresponds to the i-th Textbox placeholder TOKEN, which is what positional
    /// pairing relies on.
    /// </summary>
    private static List<IrTextbox> CollectTextboxes(IReadOnlyList<IrInline> inlines)
    {
        var boxes = new List<IrTextbox>();
        WalkForTextboxes(inlines, boxes);
        return boxes;
    }

    private static void WalkForTextboxes(IReadOnlyList<IrInline> inlines, List<IrTextbox> sink)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case IrTextbox tbx:
                    sink.Add(tbx);
                    break;
                case IrFieldRun field:
                    WalkForTextboxes(field.CachedResult, sink);
                    break;
                case IrHyperlink link:
                    WalkForTextboxes(link.Inlines, sink);
                    break;
            }
        }
    }

    /// <summary>
    /// Pair the paragraph's textboxes positionally and diff each pair's inner blocks. Returns null when there
    /// is no real textbox change to report — either side has no textboxes, OR every positionally-paired
    /// textbox is content-equal AND there is no surplus. A surplus textbox (a paragraph gained/lost a box)
    /// yields an all-insert / all-delete inner diff.
    /// </summary>
    private static List<IrTextboxDiff>? BuildTextboxDiffs(
        List<IrTextbox> leftBoxes, List<IrTextbox> rightBoxes, IrDiffSettings settings)
    {
        if (!settings.CompareTextboxes)
            return null;
        if (leftBoxes.Count == 0 && rightBoxes.Count == 0)
            return null;

        int paired = Math.Min(leftBoxes.Count, rightBoxes.Count);
        var diffs = new List<IrTextboxDiff>();
        bool anyChange = false;

        for (int i = 0; i < paired; i++)
        {
            var alignment = IrBlockAligner.AlignBlocks(leftBoxes[i].Blocks, rightBoxes[i].Blocks, settings);
            var ops = ProjectAlignment(leftBoxes[i].Blocks, alignment, settings);
            diffs.Add(new IrTextboxDiff(IrNodeList.From(ops)));
            if (ops.Any(o => o.Kind is not IrEditOpKind.EqualBlock))
                anyChange = true;
        }
        for (int i = paired; i < leftBoxes.Count; i++)
        {
            var ops = leftBoxes[i].Blocks
                .Select(b => new IrEditOp(IrEditOpKind.DeleteBlock, b.Anchor.ToString(), null, null, null, null))
                .ToList();
            diffs.Add(new IrTextboxDiff(IrNodeList.From(ops)));
            anyChange = anyChange || ops.Count > 0;
        }
        for (int i = paired; i < rightBoxes.Count; i++)
        {
            var ops = rightBoxes[i].Blocks
                .Select(b => new IrEditOp(IrEditOpKind.InsertBlock, null, b.Anchor.ToString(), null, null, null))
                .ToList();
            diffs.Add(new IrTextboxDiff(IrNodeList.From(ops)));
            anyChange = anyChange || ops.Count > 0;
        }

        return anyChange ? diffs : null;
    }

    /// <summary>The constant match key textbox placeholders collapse to when masked (so a textbox change does
    /// not surface in the paragraph's own token diff — it is reported through the nested ops instead).</summary>
    private const string MaskedTextboxKey = "tbx";

    /// <summary>
    /// Return a token list identical to <paramref name="tokens"/> except every <see cref="IrDiffTokenKind.Textbox"/>
    /// token's <see cref="IrDiffToken.MatchKey"/> is replaced by <see cref="MaskedTextboxKey"/>. Index
    /// positions are preserved, so token-op spans still index the REAL tokens; only equality is neutralized.
    /// </summary>
    private static IReadOnlyList<IrDiffToken> MaskTextboxKeys(IReadOnlyList<IrDiffToken> tokens)
    {
        var masked = new List<IrDiffToken>(tokens.Count);
        foreach (var t in tokens)
            masked.Add(t.Kind == IrDiffTokenKind.Textbox ? t with { MatchKey = MaskedTextboxKey } : t);
        return masked;
    }

    /// <summary>
    /// Token-diff a Modified (or MovedModified) pair. Paragraph pairs are tokenized + Myers-diffed;
    /// non-paragraph pairs other than tables (opaque blocks, section breaks) get a null TokenDiff — they
    /// have no sub-block token model. Tables are handled by <see cref="MakeModifyOp"/> via the table diff.
    /// </summary>
    private static IrTokenDiff? TokenDiffFor(IrBlock leftBlock, IrBlock rightBlock, IrDiffSettings settings)
    {
        if (leftBlock is IrParagraph lp && rightBlock is IrParagraph rp)
        {
            var leftTokens = IrDiffTokenizer.Tokenize(lp, settings);
            var rightTokens = IrDiffTokenizer.Tokenize(rp, settings);
            return IrTokenDiffer.Diff(leftTokens, rightTokens, settings,
                suppressLonePunctuation: PlainTextPair(lp, rp));
        }

        return null;
    }

    /// <summary>Both paragraphs are plain, sliceable direct-text-run content — the precondition for the
    /// lone-punctuation anchor rule (dropping an anchor inside a transparent field/hyperlink region
    /// re-tiles it into del+ins and double-emits the container's zero-width plumbing).</summary>
    private static bool PlainTextPair(IrParagraph lp, IrParagraph rp) =>
        IrCrossParagraphSegmenter.IsStreamable(lp) && IrCrossParagraphSegmenter.IsStreamable(rp);

    /// <summary>Decoded corpus boundary for the lone-word flush: every flushing reference instance
    /// retains a single word of at most 5 characters ("Red" 3, "font"/"text"/"bold" 4, "Green"/"Title" 5);
    /// every retaining lone-word instance is 6+ ("Roboto" 6, "Calibri"/"tracked" 7, "document" 8,
    /// "suggestions" 11). A char-FRACTION threshold ALONE provably cannot separate the classes (a kept
    /// "Roboto" sits at 0.059 of its pair, a flushed "Title" at 0.143) — but every flushing instance is
    /// ALSO a sliver of its pair (fraction &lt; 0.15), so the fraction is a necessary second condition:
    /// it keeps a short word that dominates a SHORT pair ("delta four" ↔ "delta EDITED" retains "delta"
    /// at 0.45; "same old" ↔ "same new" retains the format-tracked "same " at 0.5).</summary>
    private const int StoryFinalLoneWordFlushMaxChars = 5;

    /// <summary>See <see cref="StoryFinalLoneWordFlushMaxChars"/> — the sliver condition.</summary>
    private const double StoryFinalLoneWordFlushMaxFraction = 0.15;

    /// <summary>
    /// The story-final weak-pair FLUSH (decoded 2026-07-27 from reference compare output): the story-final
    /// word-matched pair — the per-pair form of the tail-fusion construct — drops ALL its retentions and
    /// renders as one whole ins+del rewrite when its matched content is at most ONE word of at most
    /// <see cref="StoryFinalLoneWordFlushMaxChars"/> characters that is also a sliver of the pair
    /// (2·matched-word-chars / total token chars below <see cref="StoryFinalLoneWordFlushMaxFraction"/>).
    /// "Red headings create…" ↔ "Red strikethrough text marks…" flushes even the shared leading "Red"
    /// AND the end-anchored period; ten corpus instances, every one a lone short-word sliver. Any second
    /// matched word keeps the retention (a lone shared " is … and " pair retains), as does a single LONG
    /// word ("Roboto", "Calibri", "document" all retain in the reference) or a short word dominating a
    /// SHORT pair ("delta" in "delta four" ↔ "delta EDITED"). INTERIOR pairs never flush — the reference
    /// retains even a lone shared function word mid-story ("This ") — so the caller gates this on the
    /// cleanly-paired story-final pair. Punctuation matches are anchoring artifacts, not content
    /// evidence: they count for nothing and are flushed along with the lone word.
    /// </summary>
    private static IrTokenDiff ApplyStoryFinalWeakPairFlush(
        IrTokenDiff diff, IReadOnlyList<IrDiffToken> left, IReadOnlyList<IrDiffToken> right)
    {
        int n = left.Count, m = right.Count;
        if (n == 0 || m == 0)
            return diff;

        bool anyRetained = false;
        int matchedWords = 0;
        int matchedWordChars = 0;
        int longestMatchedWord = 0;
        foreach (var op in diff.Ops)
        {
            if (op.Kind is not (IrTokenOpKind.Equal or IrTokenOpKind.FormatChanged))
                continue;
            anyRetained = true;
            for (int i = 0; i < op.LeftLength; i++)
            {
                var lt = left[op.LeftStart + i];
                if (lt.Kind != IrDiffTokenKind.Word)
                    continue;
                matchedWords++;
                int chars = Math.Min(lt.Text.Length, right[op.RightStart + i].Text.Length);
                matchedWordChars += chars;
                longestMatchedWord = Math.Max(longestMatchedWord, chars);
            }
        }
        if (!anyRetained)
            return diff; // already a whole rewrite
        if (matchedWords > 1 || longestMatchedWord > StoryFinalLoneWordFlushMaxChars)
            return diff;

        int totalChars = 0;
        foreach (var t in left)
            totalChars += t.Text.Length;
        foreach (var t in right)
            totalChars += t.Text.Length;
        if (totalChars == 0 ||
            2.0 * matchedWordChars / totalChars >= StoryFinalLoneWordFlushMaxFraction)
            return diff;

        return new IrTokenDiff(IrNodeList.From(new List<IrTokenOp>
        {
            new(IrTokenOpKind.Delete, 0, n, 0, 0),
            new(IrTokenOpKind.Insert, n, n, 0, m),
        }));
    }

    // ------------------------------------------------------------------ move interleave

    /// <summary>Map each left block to its index by reference identity (for deterministic ordering).</summary>
    private static Dictionary<IrBlock, int> BuildLeftIndexMap(IrNodeList<IrBlock> blocks)
    {
        var map = new Dictionary<IrBlock, int>(ReferenceEqualityComparer.Instance);
        for (int i = 0; i < blocks.Count; i++)
            map[blocks[i]] = i;
        return map;
    }

    /// <summary>
    /// Bucket each move-source left block under the left index of the nearest preceding PAIRED-IN-PLACE
    /// left block (left-anchored convention; -1 = front). "Paired-in-place" = the left block participated
    /// as the left partner of an Unchanged/FormatOnly/Modified op (a move destination never carries a
    /// left block; a Deleted left block is itself removed and does not anchor). We walk the LEFT document
    /// order so the adjacency exactly mirrors the aligner's deletion interleave.
    /// </summary>
    private static Dictionary<int, List<int>> BuildSourceInterleave(
        IrNodeList<IrBlock> blocks, IrBlockAlignment alignment,
        Dictionary<IrBlock, int> leftIndex, Dictionary<int, MoveInfo> moves)
    {
        var pairedInPlace = new HashSet<int>();
        foreach (var entry in alignment.Entries)
        {
            if (entry.Left is not null && IsPairedInPlace(entry.Kind))
                pairedInPlace.Add(leftIndex[entry.Left]);

            // A Merge entry's N left members are all paired in place (consumed by the merge, not
            // deleted), so each anchors the move-source interleave exactly like a Modified left.
            if (entry.Kind == IrAlignmentKind.Merge && entry.MultiBlocks is { } lefts)
                foreach (var lb in lefts)
                    pairedInPlace.Add(leftIndex[lb]);
        }

        var sourcesAfterLeft = new Dictionary<int, List<int>>();
        int lastPairedLeft = -1;
        for (int i = 0; i < blocks.Count; i++)
        {
            if (moves.ContainsKey(i)) // this left block is a move source
            {
                if (!sourcesAfterLeft.TryGetValue(lastPairedLeft, out var list))
                    sourcesAfterLeft[lastPairedLeft] = list = new List<int>();
                list.Add(i);
            }
            else if (pairedInPlace.Contains(i))
            {
                lastPairedLeft = i;
            }
        }

        return sourcesAfterLeft;
    }

    // Split counts: its entry.Left is the singular left block, paired in place at the entry's right
    // position (the N right members replace it there). Merge is handled separately — its entry.Left is
    // null; its N left members are added to the paired-in-place set explicitly (see BuildSourceInterleave)
    // and flushed explicitly after the Merge entry in ProjectAlignment.
    private static bool IsPairedInPlace(IrAlignmentKind kind) =>
        kind is IrAlignmentKind.Unchanged or IrAlignmentKind.FormatOnly or IrAlignmentKind.Modified
            or IrAlignmentKind.Split;

}
