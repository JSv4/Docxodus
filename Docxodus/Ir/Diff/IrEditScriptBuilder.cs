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
    private readonly record struct MoveInfo(int GroupId, IrBlock LeftBlock, IrEditOpKind OpKind);

    public static IrEditScript Build(IrDocument left, IrDocument right, IrDiffSettings settings)
    {
        ArgumentNullException.ThrowIfNull(left);
        ArgumentNullException.ThrowIfNull(right);
        ArgumentNullException.ThrowIfNull(settings);

        var alignment = IrBlockAligner.Align(left, right, settings);
        var bodyOps = ProjectAlignment(left.Body.Blocks, alignment, settings);
        var noteOps = BuildNoteOps(left, right, settings);
        return new IrEditScript(IrNodeList.From(bodyOps),
            noteOps.Count == 0 ? null : IrNodeList.From(noteOps));
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

    /// <summary>Walk the body in document order (recursing into tables, textboxes, fields, and hyperlinks the
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
    /// tokens of every paragraph block concatenated in document order, keyed by <see cref="IrDiffToken.MatchKey"/>.
    /// Whitespace, separators, note-ref/opaque markers and other non-word tokens are EXCLUDED: they are shared
    /// near-uniformly across notes (every note opens with the same note-ref marker and is mostly spaces) and
    /// would otherwise dominate a Jaccard score, pairing notes by length rather than by content. Non-paragraph
    /// blocks contribute nothing. The residue similarity is a coarse content-overlap heuristic — it disambiguates
    /// which reference a leftover note corresponds to, never gating a forced lone-left/lone-right pair.</summary>
    private static Dictionary<string, int> NoteTokenBag(
        Dictionary<string, Dictionary<string, int>> cache, string cacheKey, IrScope scope, IrDiffSettings settings)
    {
        if (cache.TryGetValue(cacheKey, out var bag)) return bag;
        bag = new Dictionary<string, int>();
        foreach (var block in scope.Blocks)
            if (block is IrParagraph p)
                foreach (var t in IrDiffTokenizer.Tokenize(p, settings))
                    if (t.Kind == IrDiffTokenKind.Word)
                        bag[t.MatchKey] = bag.TryGetValue(t.MatchKey, out int c) ? c + 1 : 1;
        cache[cacheKey] = bag;
        return bag;
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
        IrNodeList<IrBlock> leftBlocks, IrBlockAlignment alignment, IrDiffSettings settings)
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
                moves[li] = new MoveInfo(nextGroup++, entry.Left!, opKind);
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

        foreach (var entry in alignment.Entries)
        {
            if (pendingSources.Count > 0)
            {
                // A Deleted entry releases only the staged sources whose left index PRECEDES it; any
                // other entry ends the anchor's deletion run and releases everything staged.
                int limit = entry.Kind == IrAlignmentKind.Deleted ? leftIndex[entry.Left!] : int.MaxValue;
                FlushPendingSources(pendingSources, limit, moves, ops);
            }

            switch (entry.Kind)
            {
                case IrAlignmentKind.Unchanged:
                    ops.Add(new IrEditOp(IrEditOpKind.EqualBlock,
                        entry.Left!.Anchor.ToString(), entry.Right!.Anchor.ToString(),
                        null, null, null));
                    break;

                case IrAlignmentKind.FormatOnly:
                    ops.Add(new IrEditOp(IrEditOpKind.FormatOnlyBlock,
                        entry.Left!.Anchor.ToString(), entry.Right!.Anchor.ToString(),
                        null, null, null));
                    break;

                case IrAlignmentKind.Modified:
                    ops.Add(MakeModifyOp(entry.Left!, entry.Right!, settings));
                    break;

                case IrAlignmentKind.Inserted:
                    ops.Add(new IrEditOp(IrEditOpKind.InsertBlock,
                        null, entry.Right!.Anchor.ToString(), null, null, null));
                    break;

                case IrAlignmentKind.Deleted:
                    ops.Add(new IrEditOp(IrEditOpKind.DeleteBlock,
                        entry.Left!.Anchor.ToString(), null, null, null, null));
                    break;

                // The detection gate (IrBlockAligner split/merge scan) only ever groups PARAGRAPH
                // blocks into a Split/Merge entry, so the IrParagraph casts below are safe.
                case IrAlignmentKind.Split:
                {
                    var lp = (IrParagraph)entry.Left!;
                    var members = entry.MultiBlocks!.Cast<IrParagraph>().ToList();
                    ops.Add(new IrEditOp(IrEditOpKind.SplitBlock,
                        lp.Anchor.ToString(), null, null, null, null, null, null,
                        IrNodeList.From(members.Select(m => m.Anchor.ToString()).ToList()),
                        IrSplitSegmenter.ComputeSegmentDiffs(lp, members, settings)));
                    break;
                }

                case IrAlignmentKind.Merge:
                {
                    var rp = (IrParagraph)entry.Right!;
                    var members = entry.MultiBlocks!.Cast<IrParagraph>().ToList();
                    // Segment diffs are computed singular-vs-members (rp sliced against each left
                    // member) then MIRRORED so each stored diff reads left-member → right-slice,
                    // keeping the universal "left = left document" orientation for every consumer.
                    var sliced = IrSplitSegmenter.ComputeSegmentDiffs(rp, members, settings);
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
                        tokenDiff, move.GroupId, IsMoveSource: false));
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
                move.OpKind, move.LeftBlock.Anchor.ToString(), null, null, move.GroupId, IsMoveSource: true));
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
    private static IrEditOp MakeModifyOp(IrBlock leftBlock, IrBlock rightBlock, IrDiffSettings settings)
    {
        if (leftBlock is IrTable lt && rightBlock is IrTable rt)
            return new IrEditOp(IrEditOpKind.ModifyBlock,
                leftBlock.Anchor.ToString(), rightBlock.Anchor.ToString(),
                null, null, null, IrTableDiffer.Diff(lt, rt, settings));

        if (leftBlock is IrParagraph lp && rightBlock is IrParagraph rp)
            return MakeParagraphModifyOp(lp, rp, settings);

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
    private static IrEditOp MakeParagraphModifyOp(IrParagraph lp, IrParagraph rp, IrDiffSettings settings)
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
        var tokenDiff = IrTokenDiffer.Diff(diffLeft, diffRight, settings);

        return new IrEditOp(IrEditOpKind.ModifyBlock,
            lp.Anchor.ToString(), rp.Anchor.ToString(),
            tokenDiff, null, null, null,
            nest ? IrNodeList.From(textboxDiffs!) : null);
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
            return IrTokenDiffer.Diff(leftTokens, rightTokens, settings);
        }

        return null;
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
