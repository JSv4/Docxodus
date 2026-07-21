#nullable enable

using System.Collections.Generic;
using System.Text;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Renders an <see cref="IrEditScript"/> into a flat, ordered list of consumer-facing
/// <see cref="IrRevision"/>s (M2.3 Task 1) — the IR engine's first <c>WmlComparer.GetRevisions</c>-shaped
/// output. Each block- and token-level edit op projects to one or more revisions in SCRIPT ORDER;
/// <see cref="IrEditOpKind.EqualBlock"/> projects to nothing.
/// </summary>
/// <remarks>
/// <para><b>Author/Date.</b> Every revision is stamped with
/// <see cref="IrDiffSettings.AuthorForRevisions"/> and <see cref="IrDiffSettings.DateTimeForRevisions"/>
/// — deterministic epoch by default (see those members). The renderer never reads the wall clock itself.</para>
///
/// <para><b>Block text.</b> A block's revision <see cref="IrRevision.Text"/> is its concatenated raw token
/// text: for a paragraph, the tokenizer's <see cref="IrDiffToken.Text"/> joined in order (raw, NOT match
/// keys — so case/NBSP/link normalization does not leak into the surface); for a table, every descendant
/// paragraph's text joined the same way. A non-paragraph/non-table block (opaque, section break) yields
/// empty text. Text is always non-null (possibly empty), which the corpus smoke asserts.</para>
///
/// <para><b>ModifyBlock token ops.</b> Projected in token-diff op order: each Insert span → one Inserted
/// (right-token raw text), each Delete span → one Deleted (left-token raw text), each FormatChanged span →
/// one OR MORE FormatChanged revisions. A FormatChanged token span is a maximal run of format-differing
/// positions, but the (oldFormat,newFormat) transition can be HETEROGENEOUS across the span (e.g. positions
/// 0–1 go non-bold→bold while position 2 goes 10pt→12pt). We therefore split the span into maximal SUB-RUNS
/// of UNIFORM (modeled-old-key, modeled-new-key) and emit one FormatChanged revision per sub-run, its text =
/// that sub-run's right-token raw text and its details = the sub-run's single transition. Equal spans emit
/// nothing.</para>
///
/// <para><b>FormatOnlyBlock.</b> Content-equal, modeled-format-differing block pair. We tokenize both sides;
/// when the token counts match we pair positionally and emit a FormatChanged revision per uniform sub-run of
/// differing positions (same sub-run rule as ModifyBlock). When counts differ — the known run-boundary
/// word-split case where two content-equal paragraphs tokenize to different token counts — we FALL BACK to a
/// single FormatChanged revision for the whole block, with details from the FIRST position at which the
/// per-token modeled-format keys diverge under positional scan of the shorter length (or, if every paired
/// position agrees, the first position only present on one side). Documented fallback; rare in practice.</para>
///
/// <para><b>Moves.</b> MoveBlock → two Moved revisions sharing a <see cref="IrRevision.MoveGroupId"/>: a
/// source (<see cref="IrRevision.IsMoveSource"/>=true, left block text) and a destination (false, right
/// block text). They are emitted at their op positions in script order (source op and destination op are
/// already separately placed by the builder). MoveModifyBlock additionally emits the destination's nested
/// token-op revisions (Inserted/Deleted/FormatChanged, exactly as ModifyBlock) IMMEDIATELY AFTER the
/// destination Moved revision — the ordering rule: relocate first, then describe the in-move edits, so a
/// consumer reads "this block moved here, and here is what changed within it".</para>
///
/// <para><b>Tables (TableDiff recursion).</b> A ModifyBlock carrying an <see cref="IrTableDiff"/> recurses:
/// InsertRow → Inserted (row text), DeleteRow → Deleted (row text), MovedRow → a Moved pair (row text, shared
/// group id local to the table), ModifyRow → recurse its cell ops, each cell op recursing its block ops
/// through the SAME block-revision machinery. Row/cell anchors flow into the revisions' anchors.</para>
///
/// <para><b>Determinism.</b> Output is a pure function of the (deterministic) edit script, the two
/// documents, and the settings. No dictionary iteration order is observed.</para>
/// </remarks>
internal static class IrRevisionRenderer
{
    public static IrNodeList<IrRevision> Render(
        IrEditScript script, IrDocument left, IrDocument right, IrDiffSettings settings)
    {
        // Pre-pass: map each MoveGroupId to its source (left) block anchor. A MoveModify destination op
        // carries only the right anchor, but its token diff's Delete spans index the SOURCE block tokens,
        // so the destination needs the source anchor to resolve deleted-token text. The source op (emitted
        // separately, IsMoveSource=true) carries that left anchor.
        var moveSourceAnchor = new Dictionary<int, string>();
        foreach (var op in script.Operations)
            if (op.IsMoveSource == true && op.MoveGroupId is { } gid && op.LeftAnchor is { } la)
                moveSourceAnchor[gid] = la;

        var ctx = new Context(left, right, settings, moveSourceAnchor);
        var revisions = new List<IrRevision>();
        RenderBlockOpList(script.Operations, ctx, revisions);

        // Note scopes (M2.4 Task 1): footnotes then endnotes, in the script's deterministic note order.
        // Each note's block ops render through the SAME block-op machinery as the body — its fn/en blocks
        // are in the shared AnchorIndex, so anchor→block/token resolution works unchanged, and the note's
        // distinct fn/en anchors carry the scope context into every revision.
        //
        // NB: the adjacent-block COALESCING (RenderBlockOpList) is deliberately NOT applied to note ops.
        // WmlComparer groups note revisions PER NOTE (GetFootnoteEndnoteRevisionList builds a separate atom
        // list per footnote/endnote), so consecutive inserted note PARAGRAPHS that the body coalescing would
        // merge are reported individually by the oracle (WC-1710/1720: each endnote's paragraphs surface
        // separately). Render note ops one-per-op to preserve that grain.
        if (script.NoteOps is { } noteOps)
            foreach (var noteDiff in noteOps)
                foreach (var op in noteDiff.Ops)
                    RenderBlockOp(op, ctx, revisions);

        // Header/footer scopes (2026-07-03 campaign): appended AFTER note revisions (strictly additive
        // ordering — existing consumers' body/note indices never shift), one revision per op like notes
        // (a story is its own grouping unit; no cross-block coalescing). EXCLUDED in WmlComparerCompatible
        // mode: that granularity is defined as "match the legacy comparer's revision set", and WmlComparer
        // does not diff header/footer scopes at all — so the compatible projection reports none, keeping
        // count parity with the oracle by construction. Fine mode (the default) reports them.
        if (settings.RevisionGranularity == RevisionGranularity.Fine &&
            script.HeaderFooterOps is { } headerFooterOps)
            foreach (var hfDiff in headerFooterOps)
                foreach (var op in hfDiff.Ops)
                    RenderBlockOp(op, ctx, revisions);

        // Section-break zero-width prune (M2.4 Task 2, prelim a). A whole-block Inserted/Deleted over a
        // SECTION-BREAK block (a `sec:` anchor) carries no surface text and is a structural-only change that
        // WmlComparer does not report as a revision. Suppress it in compatible mode so it never inflates the
        // count (WC-1960). Math/image/opaque BLOCK ins/del are NOT pruned — those carry no token text either
        // but WmlComparer DOES count them (WC-1230/WC-1320), so they must survive. This prune is scoped to
        // section breaks only, and only in compatible mode, so Fine output stays byte-stable.
        //
        // Token-level zero-width ins/del (a masked-textbox placeholder Delete inside a Modified paragraph's
        // token diff — empty text, non-text Textbox token) are suppressed at the SOURCE in RenderTokenOps for
        // BOTH modes: a placeholder token carries no surface text, so reporting it as an empty Inserted/Deleted
        // is a spurious revision regardless of granularity (the real textbox change is reported through the
        // nested TextboxDiffs). See the two-textbox test.
        // Trailing section-property change (block-format-change family, Phase 3): compare the two documents'
        // trailing section formats and, when the MODELED fields differ, append one Section-scope FormatChanged
        // (mirrors the markup renderer's w:sectPrChange on the trailing sectPr). Appended after all body/note/
        // header-footer ops (additive ordering). Excluded from WmlComparerCompatible by the scope filter below.
        if (settings.TrackSectionFormatChanges && settings.EmitTrailingSectionRevision
            && ctx.Left.Body.Blocks.Count > 0 && ctx.Left.Body.Blocks[^1] is IrSectionBreak lsec
            && ctx.Right.Body.Blocks.Count > 0 && ctx.Right.Body.Blocks[^1] is IrSectionBreak rsec)
        {
            var details = IrModeledFormat.SectionFormatChangeDetails(lsec.Format, rsec.Format);
            if (details.ChangedPropertyNames.Count > 0)
                revisions.Add(new IrRevision(IrRevisionType.FormatChanged, string.Empty, ctx.Author, ctx.Date,
                    FormatChange: details,
                    LeftAnchor: lsec.Anchor.ToString(), RightAnchor: rsec.Anchor.ToString()));
        }

        if (settings.RevisionGranularity == RevisionGranularity.WmlComparerCompatible)
        {
            revisions.RemoveAll(IsSectionBreakZeroWidth);

            // Block-scope FormatChanged exclusion (block-format-change family, 2026-07-03): the compatible
            // granularity is defined as the ORACLE's revision set, and WmlComparer cannot produce
            // paragraph/table/section-scope format revisions — so the compatible projection reports none,
            // keeping count parity by construction (the hdr/ftr precedent). Fine mode reports them.
            revisions.RemoveAll(r => r.FormatChange is { } fc && fc.Scope != IrFormatChangeScope.Run);
        }

        return IrNodeList.From(revisions);
    }

    /// <summary>An empty-text Inserted/Deleted over a section-break block (a structural-only `sec:` change
    /// WmlComparer does not surface).</summary>
    private static bool IsSectionBreakZeroWidth(IrRevision r) =>
        r.Type is IrRevisionType.Inserted or IrRevisionType.Deleted && r.Text.Length == 0 &&
        ((r.RightAnchor?.StartsWith("sec:", System.StringComparison.Ordinal) ?? false) ||
         (r.LeftAnchor?.StartsWith("sec:", System.StringComparison.Ordinal) ?? false));

    /// <summary>Per-render immutable context: the two docs (for anchor→block lookup), settings, and the
    /// MoveGroupId→source-anchor map (for MoveModify destinations to resolve left-token text).</summary>
    private readonly record struct Context(
        IrDocument Left, IrDocument Right, IrDiffSettings Settings,
        IReadOnlyDictionary<int, string> MoveSourceAnchor)
    {
        public string Author => Settings.AuthorForRevisions;
        public string Date => Settings.DateTimeForRevisions;
    }

    // ------------------------------------------------------------------ adjacent-block coalescing (compat)

    /// <summary>
    /// Render an ordered list of block ops, applying (compatible mode only) WmlComparer's adjacent-block
    /// insert/delete COALESCING (M2.4b Workstream C). WmlComparer's <c>GetRevisions</c> groups the produced
    /// document's atoms by adjacent correlation status, so a maximal contiguous run of inserted (resp. deleted)
    /// blocks collapses to ONE revision whose text is the run's blocks joined with their paragraph-mark
    /// newlines. The IR's per-block <see cref="IrEditOpKind.InsertBlock"/>/<see cref="IrEditOpKind.DeleteBlock"/>
    /// ops would otherwise each surface their own revision (WC-1440/1450/1830/1840, WC-1210).
    ///
    /// <para><b>Run segmentation.</b> A run is consecutive same-direction whole-block Insert (or Delete) ops.
    /// A <see cref="IrTable"/> or opaque block STARTS A NEW SUB-REGION (it breaks the run before itself but
    /// joins with the inserts that FOLLOW it) — this reproduces WmlComparer splitting `Abcde` from the empty
    /// structural table + `fghij` that follows it (WC-1210: `Abcde` | `\n\nfghij\n`), and folding an inserted
    /// image/opaque into the adjacent paragraph run (WC-1440). Each sub-region is coalesced to ONE revision
    /// ONLY IF it carries at least one text-bearing paragraph (≥1 Word token); a sub-region with NO word
    /// content (pure math/image/opaque/empty-mark) is left as one revision PER block, because WmlComparer
    /// counts standalone math/image paragraph inserts individually (WC-1550 two-maths, WC-1320/1340/1350).</para>
    ///
    /// <para>Fine mode renders every op straight through (the engine's grain is its truth).</para>
    /// </summary>
    private static void RenderBlockOpList(IrNodeList<IrEditOp> ops, in Context ctx, List<IrRevision> sink)
    {
        if (ctx.Settings.RevisionGranularity != RevisionGranularity.WmlComparerCompatible)
        {
            foreach (var op in ops)
                RenderBlockOp(op, ctx, sink);
            return;
        }

        int i = 0;
        int n = ops.Count;
        while (i < n)
        {
            var kind = ops[i].Kind;
            if (kind is IrEditOpKind.InsertBlock or IrEditOpKind.DeleteBlock)
            {
                // Gather the maximal run of consecutive same-direction whole-block ins/del ops.
                int runEnd = i + 1;
                while (runEnd < n && ops[runEnd].Kind == kind)
                    runEnd++;
                RenderInsDelRun(ops, i, runEnd, kind, ctx, sink);
                i = runEnd;
            }
            else
            {
                RenderBlockOp(ops[i], ctx, sink);
                i++;
            }
        }
    }

    /// <summary>
    /// Render a run of consecutive same-direction whole-block ins/del ops [start,end), splitting it into
    /// sub-regions at table/opaque boundaries and coalescing each text-bearing sub-region into one revision.
    /// </summary>
    private static void RenderInsDelRun(
        IrNodeList<IrEditOp> ops, int start, int end, IrEditOpKind kind, in Context ctx, List<IrRevision> sink)
    {
        bool insert = kind == IrEditOpKind.InsertBlock;
        int subStart = start;
        for (int k = start; k < end; k++)
        {
            // A table/opaque block starts a new sub-region: flush the run accumulated BEFORE it, then begin
            // a fresh sub-region AT this block (it joins with the following inserts of the same run).
            string? anchor = insert ? ops[k].RightAnchor : ops[k].LeftAnchor;
            var doc = insert ? ctx.Right : ctx.Left;
            bool isRegionStarter = anchor is not null && doc.AnchorIndex.TryGetValue(anchor, out var b)
                && b is not IrParagraph;
            if (isRegionStarter && k > subStart)
            {
                FlushInsDelSubRegion(ops, subStart, k, insert, ctx, sink);
                subStart = k;
            }
        }
        FlushInsDelSubRegion(ops, subStart, end, insert, ctx, sink);
    }

    /// <summary>
    /// Emit one sub-region [start,end) of same-direction whole-block ins/del ops. If the sub-region has any
    /// text-bearing paragraph, emit ONE coalesced revision (block texts joined with their paragraph-mark
    /// newlines, empty-mark paragraphs still pruned); otherwise emit one revision per block (the per-block
    /// path, which prunes empty-mark paragraphs and keeps standalone math/image inserts).
    /// </summary>
    private static void FlushInsDelSubRegion(
        IrNodeList<IrEditOp> ops, int start, int end, bool insert, in Context ctx, List<IrRevision> sink)
    {
        if (end - start <= 1 || !SubRegionHasText(ops, start, end, insert, ctx))
        {
            // Single op, or no word content: render each op individually (prunes empty marks, keeps math/image).
            for (int k = start; k < end; k++)
                RenderBlockOp(ops[k], ctx, sink);
            return;
        }

        // Coalesce: join the blocks' texts with the paragraph-mark convention. WmlComparer surfaces each
        // paragraph's content followed by its mark (a newline), so the run reads as one multi-paragraph
        // ins/del. An empty-mark paragraph contributes only its newline; a math/image paragraph contributes
        // its surface text then a newline (matching the oracle's coalesced multi-paragraph text).
        string? firstAnchor = insert ? ops[start].RightAnchor : ops[start].LeftAnchor;
        var doc = insert ? ctx.Right : ctx.Left;
        // GetRevisions groups adjacent atoms by revision metadata, then bases the whole group's text
        // on its first atom.  Math and drawings are its explicit no-text atoms, so a coalesced region
        // which begins with one remains a counted revision but has no semantic text even if a later
        // paragraph carries words (WC-1840's math paragraph followed by "Click").
        bool legacyNoTextLeader = StartsWithLegacyNoTextRevisionContent(firstAnchor, doc);
        var sb = new StringBuilder();
        for (int k = start; k < end; k++)
        {
            string? anchor = insert ? ops[k].RightAnchor : ops[k].LeftAnchor;
            sb.Append(BlockText(anchor, insert ? ctx.Right : ctx.Left, ctx.Settings));
            sb.Append('\n');
        }
        sink.Add(insert
            ? new IrRevision(IrRevisionType.Inserted, legacyNoTextLeader ? string.Empty : sb.ToString(),
                ctx.Author, ctx.Date, RightAnchor: firstAnchor)
            : new IrRevision(IrRevisionType.Deleted, legacyNoTextLeader ? string.Empty : sb.ToString(),
                ctx.Author, ctx.Date, LeftAnchor: firstAnchor));
    }

    /// <summary>True iff any block in the ins/del sub-region [start,end) is a paragraph with ≥1 Word token.</summary>
    private static bool SubRegionHasText(IrNodeList<IrEditOp> ops, int start, int end, bool insert, in Context ctx)
    {
        var doc = insert ? ctx.Right : ctx.Left;
        for (int k = start; k < end; k++)
        {
            string? anchor = insert ? ops[k].RightAnchor : ops[k].LeftAnchor;
            if (anchor is not null && doc.AnchorIndex.TryGetValue(anchor, out var b) && b is IrParagraph p)
            {
                foreach (var t in IrDiffTokenizer.Tokenize(p, ctx.Settings))
                    if (t.Kind == IrDiffTokenKind.Word)
                        return true;
            }
        }
        return false;
    }

    /// <summary>
    /// Whether a coalesced region's first paragraph begins with the same content kind that the legacy
    /// comparer deliberately reports with null text: Office Math or a drawing.  The IR keeps a non-null
    /// string surface, so the equivalent observable value here is <see cref="string.Empty"/>.
    /// </summary>
    private static bool StartsWithLegacyNoTextRevisionContent(string? anchor, IrDocument doc)
    {
        if (anchor is null || !doc.AnchorIndex.TryGetValue(anchor, out var block) || block is not IrParagraph p)
            return false;

        foreach (var inline in p.Inlines)
        {
            if (inline is IrTextRun { Text.Length: 0 })
                continue;
            if (inline is IrInlineImage)
                return true;
            if (inline is IrOpaqueInline opaque &&
                (opaque.ElementName == M.oMath || opaque.ElementName == M.oMathPara))
                return true;
            return false;
        }
        return false;
    }

    // ------------------------------------------------------------------ block-op dispatch

    private static void RenderBlockOp(IrEditOp op, in Context ctx, List<IrRevision> sink)
    {
        switch (op.Kind)
        {
            case IrEditOpKind.EqualBlock:
                break;

            case IrEditOpKind.InsertBlock:
                // Empty-paragraph-mark prune (M2.4b Workstream B). A whole-block insert of a paragraph that
                // carries NO content tokens (a bare paragraph mark — e.g. the empty cell paragraph a
                // moved-into-table block leaves behind, WC-1190) has empty surface text; WmlComparer reports
                // no revision for it (an empty w:ins paragraph has no text run to key on). Suppress it in
                // compatible mode so it never inflates the count. A paragraph with ANY content token (text,
                // image, math/opaque) still emits — those WmlComparer counts (WC-1230/1320 math/image blocks).
                if (IsZeroWidthBlock(op.RightAnchor, ctx.Right, ctx.Settings))
                    break;
                sink.Add(new IrRevision(IrRevisionType.Inserted,
                    BlockText(op.RightAnchor, ctx.Right, ctx.Settings), ctx.Author, ctx.Date,
                    RightAnchor: op.RightAnchor));
                break;

            case IrEditOpKind.DeleteBlock:
                if (IsZeroWidthBlock(op.LeftAnchor, ctx.Left, ctx.Settings))
                    break;
                sink.Add(new IrRevision(IrRevisionType.Deleted,
                    BlockText(op.LeftAnchor, ctx.Left, ctx.Settings), ctx.Author, ctx.Date,
                    LeftAnchor: op.LeftAnchor));
                break;

            case IrEditOpKind.FormatOnlyBlock:
                RenderFormatOnlyBlock(op, ctx, sink);
                break;

            case IrEditOpKind.ModifyBlock:
                RenderModifyBlock(op, ctx, sink);
                break;

            case IrEditOpKind.MoveBlock:
            case IrEditOpKind.MoveModifyBlock:
                RenderMoveOp(op, ctx, sink);
                break;

            case IrEditOpKind.SplitBlock:
                RenderSplitBlock(op, ctx, sink);
                break;

            case IrEditOpKind.MergeBlock:
                RenderMergeBlock(op, ctx, sink);
                break;
        }
    }

    // ------------------------------------------------------------------ split / merge (M2.6)

    /// <summary>
    /// Render a 1:N paragraph split. FINE mode is the engine's truth: each segment's token diff
    /// renders through <see cref="RenderTokenOps"/> over (left slice, member tokens) — a clean split
    /// yields no revisions at all (every token is Equal; the structural mark account lives in the edit
    /// script, not this surface), an edited split yields exactly its per-segment ins/del.
    /// <para><b>Compatible mode</b> reproduces WmlComparer's account of a split. The oracle expresses a
    /// split as ONE contiguous inserted region spanning the split-off paragraphs — the new mark plus
    /// each member's text plus its mark, e.g. WC-1830's <c>Inserted "\nA=πr2\nWhen you click…add.\n"</c>
    /// (the oracle's delete-side counterpart, the re-deleted tail, COALESCES into the adjacent deleted
    /// paragraph's region and so adds no count of its own; when no deleted neighbor exists the oracle's
    /// anchored account is the same single inserted region — either way the split contributes its
    /// member-0 inline edits plus exactly ONE Inserted). So: segment 0 renders through the compat token
    /// path (its inline edits, e.g. a prefix insert, surface normally), and members 1..N-1 collapse to
    /// one Inserted whose text is <c>"\n" + Σ(memberText + "\n")</c> — the leading newline is the new
    /// pilcrow that made the split. The trim/score gates guarantee at least one split-off member has
    /// content, so the coalesced text is never empty.</para>
    /// </summary>
    private static void RenderSplitBlock(IrEditOp op, in Context ctx, List<IrRevision> sink)
    {
        if (op.SplitMergeAnchors is not { } anchors || op.SegmentDiffs is not { } diffs)
            return; // malformed op — the pairing assert catches this in tests; render nothing here
        var leftTokens = ParagraphTokens(op.LeftAnchor, ctx.Left, ctx.Settings);

        int offset = 0;
        for (int s = 0; s < anchors.Count; s++)
        {
            var diff = diffs[s];
            int sliceLen = SegmentLeftLength(diff);
            var slice = SubList(leftTokens, offset, sliceLen);
            offset += sliceLen;

            if (s == 0 || ctx.Settings.RevisionGranularity != RevisionGranularity.WmlComparerCompatible)
            {
                var memberTokens = ParagraphTokens(anchors[s], ctx.Right, ctx.Settings);
                RenderSegmentTokenOps(diff, slice, memberTokens, op.LeftAnchor, anchors[s], ctx, sink);
            }

            // Fine mode reports each NEW pilcrow as its own Inserted "\n" (spec §4.4: the per-segment
            // account plus each new mark) — without it a CLEAN split would be invisible on this
            // surface despite being a real document change (the mark IS content at the text level).
            if (s > 0 && ctx.Settings.RevisionGranularity != RevisionGranularity.WmlComparerCompatible)
                sink.Add(new IrRevision(IrRevisionType.Inserted, "\n", ctx.Author, ctx.Date,
                    LeftAnchor: op.LeftAnchor, RightAnchor: anchors[s]));
        }

        if (ctx.Settings.RevisionGranularity == RevisionGranularity.WmlComparerCompatible && anchors.Count > 1)
        {
            var sb = new StringBuilder();
            sb.Append('\n'); // the inserted pilcrow that split the paragraph
            for (int s = 1; s < anchors.Count; s++)
            {
                sb.Append(BlockText(anchors[s], ctx.Right, ctx.Settings));
                sb.Append('\n');
            }
            sink.Add(new IrRevision(IrRevisionType.Inserted, sb.ToString(), ctx.Author, ctx.Date,
                LeftAnchor: op.LeftAnchor, RightAnchor: anchors[1]));
        }
    }

    /// <summary>
    /// Render one split/merge segment's token diff, suppressing (compatible mode only) revisions whose
    /// text is pure whitespace — the SEAM separator the segmentation left on one side of the split
    /// boundary (e.g. WC-1450's <c>"point. "</c> trailing space when the prefix member ends at
    /// <c>"point."</c>). WmlComparer accounts that seam space inside its coalesced tail region, so a
    /// standalone whitespace-only ins/del here would double-count it (+1 vs the oracle). Fine mode
    /// passes everything through untouched.
    /// </summary>
    private static void RenderSegmentTokenOps(
        IrTokenDiff diff,
        IReadOnlyList<IrDiffToken> leftTokens, IReadOnlyList<IrDiffToken> rightTokens,
        string? leftAnchor, string? rightAnchor, in Context ctx, List<IrRevision> sink)
    {
        if (ctx.Settings.RevisionGranularity != RevisionGranularity.WmlComparerCompatible)
        {
            RenderTokenOps(diff, leftTokens, rightTokens, leftAnchor, rightAnchor, ctx, sink);
            return;
        }

        var buffer = new List<IrRevision>();
        RenderTokenOps(diff, leftTokens, rightTokens, leftAnchor, rightAnchor, ctx, buffer);
        buffer.RemoveAll(rv => rv.Type is IrRevisionType.Inserted or IrRevisionType.Deleted
            && rv.Text.Length > 0 && rv.Text.Trim().Length == 0);
        sink.AddRange(buffer);
    }

    /// <summary>
    /// Render an N:1 paragraph merge — the byte-mirror of <see cref="RenderSplitBlock"/>. The stored
    /// segment diffs read left-member → right-slice, so member 0's diff renders directly through the
    /// token path against the right stream's first slice. Compatible mode collapses members 1..N-1 to
    /// one Deleted (<c>"\n" + Σ(memberText + "\n")</c> — the removed pilcrows and the re-deleted tail,
    /// WmlComparer's account of a merge); Fine mode renders every member's segment diff (a clean merge
    /// has no token-level changes and reports nothing — the mark removal is the edit script's account).
    /// </summary>
    private static void RenderMergeBlock(IrEditOp op, in Context ctx, List<IrRevision> sink)
    {
        if (op.SplitMergeAnchors is not { } anchors || op.SegmentDiffs is not { } diffs)
            return;
        var rightTokens = ParagraphTokens(op.RightAnchor, ctx.Right, ctx.Settings);

        int offset = 0;
        for (int m = 0; m < anchors.Count; m++)
        {
            var diff = diffs[m];
            int sliceLen = SegmentRightLength(diff);
            var slice = SubList(rightTokens, offset, sliceLen);
            offset += sliceLen;

            if (m == 0 || ctx.Settings.RevisionGranularity != RevisionGranularity.WmlComparerCompatible)
            {
                var memberTokens = ParagraphTokens(anchors[m], ctx.Left, ctx.Settings);
                RenderSegmentTokenOps(diff, memberTokens, slice, anchors[m], op.RightAnchor, ctx, sink);
            }

            // Fine mode reports each REMOVED pilcrow as its own Deleted "\n" (mirror of the split's
            // new-mark account; a clean merge is otherwise invisible on this surface).
            if (m > 0 && ctx.Settings.RevisionGranularity != RevisionGranularity.WmlComparerCompatible)
                sink.Add(new IrRevision(IrRevisionType.Deleted, "\n", ctx.Author, ctx.Date,
                    LeftAnchor: anchors[m], RightAnchor: op.RightAnchor));
        }

        if (ctx.Settings.RevisionGranularity == RevisionGranularity.WmlComparerCompatible && anchors.Count > 1)
        {
            var sb = new StringBuilder();
            sb.Append('\n'); // the removed pilcrow that joined the paragraphs
            for (int m = 1; m < anchors.Count; m++)
            {
                sb.Append(BlockText(anchors[m], ctx.Left, ctx.Settings));
                sb.Append('\n');
            }
            sink.Add(new IrRevision(IrRevisionType.Deleted, sb.ToString(), ctx.Author, ctx.Date,
                LeftAnchor: anchors[1], RightAnchor: op.RightAnchor));
        }
    }

    /// <summary>A segment diff's singular-side slice length: Σ non-Insert left-span lengths (the F3.3
    /// partition convention — boundaries are implicit in the diff ops).</summary>
    private static int SegmentLeftLength(IrTokenDiff diff)
    {
        int n = 0;
        foreach (var o in diff.Ops)
            if (o.Kind != IrTokenOpKind.Insert)
                n += o.LeftLength;
        return n;
    }

    /// <summary>A merge segment diff's right-slice length: Σ non-Delete right-span lengths (the stored
    /// merge orientation is member → right-slice).</summary>
    private static int SegmentRightLength(IrTokenDiff diff)
    {
        int n = 0;
        foreach (var o in diff.Ops)
            if (o.Kind != IrTokenOpKind.Delete)
                n += o.RightLength;
        return n;
    }

    private static IReadOnlyList<IrDiffToken> SubList(IReadOnlyList<IrDiffToken> tokens, int offset, int len)
    {
        var list = new List<IrDiffToken>(len);
        for (int i = offset; i < offset + len && i < tokens.Count; i++)
            list.Add(tokens[i]);
        return list;
    }

    // ------------------------------------------------------------------ modify / move

    /// <summary>
    /// True iff (compatible mode only) the anchor resolves to a PARAGRAPH with ZERO content tokens — a bare
    /// paragraph mark (only separators, or wholly empty). Such a paragraph's whole-block insert/delete carries
    /// no surface text and no <c>w:r</c> text run, so <c>WmlComparer.GetRevisions</c> surfaces no revision for
    /// it (WC-1190: the empty cell paragraph a moved-into-table block leaves behind). Pruning it keeps count
    /// parity. A paragraph with ANY content token — text, image, OR math/opaque — still emits: WmlComparer
    /// DOES report whole-block math/SmartArt/image paragraph inserts and deletes (WC-1320 deleted SmartArt,
    /// WC-1550 two-maths, WC-1340/1350 images), so the prune is strictly the empty-mark case.
    /// </summary>
    private static bool IsZeroWidthBlock(string? anchor, IrDocument doc, IrDiffSettings settings)
    {
        if (settings.RevisionGranularity != RevisionGranularity.WmlComparerCompatible)
            return false;
        if (anchor is null || !doc.AnchorIndex.TryGetValue(anchor, out var block) || block is not IrParagraph p)
            return false;
        // Scope (M2.4b Workstream C): the empty-mark prune applies to BODY paragraphs only. In a footnote/
        // endnote scope, WmlComparer's per-note atom grouping DOES surface an inserted/deleted empty paragraph
        // as a revision (WC-1750/1760: deleting a table's trailing rows leaves an empty-paragraph insert the
        // oracle reports as `\n`), so a note-scope empty mark must NOT be pruned. The WC-1190 prune case (a
        // moved-into-table block's leftover empty cell mark) is a BODY-scope anchor.
        if (anchor.StartsWith("p:fn:", System.StringComparison.Ordinal)
            || anchor.StartsWith("p:en:", System.StringComparison.Ordinal))
            return false;
        return CountContent(IrDiffTokenizer.Tokenize(p, settings)) == 0;
    }

    private static void RenderModifyBlock(IrEditOp op, in Context ctx, List<IrRevision> sink)
    {
        // Inline SDT/smartTag envelopes and non-hyperlink field code/state carriers are not token-sliceable.
        // Even when visible text is identical, the markup renderer emits a whole old/new paragraph pair so the
        // carrier itself is reversible; surface the same pair here instead of silently reporting no revision.
        if (op.RequiresWholeParagraphReplace)
        {
            if (op.LeftAnchor is { } wholeLeft)
                sink.Add(new IrRevision(IrRevisionType.Deleted, BlockText(wholeLeft, ctx.Left, ctx.Settings),
                    ctx.Author, ctx.Date, LeftAnchor: wholeLeft, RightAnchor: op.RightAnchor));
            if (op.RightAnchor is { } wholeRight)
                sink.Add(new IrRevision(IrRevisionType.Inserted, BlockText(wholeRight, ctx.Right, ctx.Settings),
                    ctx.Author, ctx.Date, LeftAnchor: op.LeftAnchor, RightAnchor: wholeRight));
            return;
        }

        // Block-level SDTs are deliberately atomic: metadata, wrapper topology and nested controls belong to
        // the envelope, not to an independently token-diffable paragraph. Surface an explicit old/new pair
        // (even for a metadata-only change whose visible text happens to be identical) so GetRevisions agrees
        // with the markup renderer's whole-control replacement.
        if ((op.LeftAnchor is { } sdtLeftAnchor && ctx.Left.AnchorIndex.TryGetValue(sdtLeftAnchor, out var leftBlock) && leftBlock is IrSdtBlock) ||
            (op.RightAnchor is { } sdtRightAnchor && ctx.Right.AnchorIndex.TryGetValue(sdtRightAnchor, out var rightBlock) && rightBlock is IrSdtBlock))
        {
            if (op.LeftAnchor is { } leftAnchor)
                sink.Add(new IrRevision(IrRevisionType.Deleted, BlockText(leftAnchor, ctx.Left, ctx.Settings),
                    ctx.Author, ctx.Date, LeftAnchor: leftAnchor, RightAnchor: op.RightAnchor));
            if (op.RightAnchor is { } rightAnchor)
                sink.Add(new IrRevision(IrRevisionType.Inserted, BlockText(rightAnchor, ctx.Right, ctx.Settings),
                    ctx.Author, ctx.Date, LeftAnchor: op.LeftAnchor, RightAnchor: rightAnchor));
            return;
        }

        if (op.TableDiff is { } tableDiff)
        {
            // A left-only cell (remove/merge topology) still bails the markup renderer to a whole-table
            // del(left)+ins(right) fallback. A right-only ordinary-grid insertion stays granular only while
            // table-shell tracking is enabled: tblGridChange is what makes its widened geometry reversible.
            if (TableDiffNeedsWholeTableFallback(tableDiff, ctx.Settings.TrackTableFormatChanges))
            {
                if (op.LeftAnchor is { } la)
                    sink.Add(new IrRevision(IrRevisionType.Deleted, BlockText(la, ctx.Left, ctx.Settings),
                        ctx.Author, ctx.Date, LeftAnchor: la));
                if (op.RightAnchor is { } ra)
                    sink.Add(new IrRevision(IrRevisionType.Inserted, BlockText(ra, ctx.Right, ctx.Settings),
                        ctx.Author, ctx.Date, RightAnchor: ra));
                return;
            }
            RenderTableDiff(tableDiff, ctx, sink);
            // Block-format-change family (2026-07-03): report table/row/cell SHELL changes (tblPr/tblGrid/
            // trPr/tcPr) as digest-grade FormatChanged revisions, after the cell text revisions. Gated on
            // TrackBlockFormatChanges so the Consolidate ceiling holds on the REVISIONS surface too — the
            // composite renderers force the flag off, and the markup emits no *PrChange there.
            if (ctx.Settings.TrackTableFormatChanges
                && ResolveTable(op.LeftAnchor, ctx.Left) is { } lt && ResolveTable(op.RightAnchor, ctx.Right) is { } rt)
                EmitTableModifiedShellRevisions(lt, rt, tableDiff, ctx, sink);
            return;
        }

        if (op.TokenDiff is { } tokenDiff)
        {
            // Block-format-change family (2026-07-03): a Modified paragraph whose pPr ALSO changed reports
            // the Paragraph-scope FormatChanged first, then its token-level revisions (mirrors the markup
            // renderer stamping w:pPrChange on the modified paragraph).
            EmitParagraphScopeFormatChanged(op, ctx, sink);
            EmitInlineSectionFormatChanged(op, ctx, sink);

            var leftTokens = ParagraphTokens(op.LeftAnchor, ctx.Left, ctx.Settings);
            var rightTokens = ParagraphTokens(op.RightAnchor, ctx.Right, ctx.Settings);
            RenderTokenOps(tokenDiff, leftTokens, rightTokens, op.LeftAnchor, op.RightAnchor, ctx, sink);
        }

        // Textbox interiors (M2.4 Task 1): a Modified paragraph carrying textbox diffs recurses each
        // textbox's inner block ops through the SAME block-op machinery, AFTER the paragraph's own token
        // ops. The placeholder-token change was masked out of the token diff above, so the textbox change
        // is reported exactly once — here, from the inner blocks' text.
        if (op.TextboxDiffs is { } textboxDiffs)
            RenderTextboxDiffs(textboxDiffs, ctx, sink);

        // A non-paragraph, non-table Modified pair (opaque / section break) has no sub-block model and
        // no token diff — it produces no token-level revisions (its content change is not describable at
        // this granularity by this surface; M2.4 OOXML markup is the place for it).
    }

    /// <summary>
    /// Render a Modified paragraph's textbox-interior diffs. In Fine mode every textbox diff renders straight
    /// through. In WmlComparer-compatible mode we DEDUP the Choice/Fallback duplicate: Word emits one logical
    /// textbox twice inside an <c>mc:AlternateContent</c> — a DrawingML <c>mc:Choice</c> and a VML
    /// <c>mc:Fallback</c> with byte-identical inner content — and the IR reader (by design, to mirror the
    /// oracle's both-branch text walk) emits an <see cref="IrTextbox"/> per occurrence, so the two adjacent
    /// textbox diffs render to IDENTICAL revision sequences. WmlComparer MC-preprocesses its input and sees
    /// only the Choice branch, reporting the change once. We reproduce that at render time by collapsing a
    /// textbox diff whose rendered revisions are value-equal to those of the IMMEDIATELY PRECEDING textbox diff
    /// (the Choice/Fallback adjacency). This is a render-time projection choice — the edit script's textbox
    /// diffs are untouched, and Fine mode still reports both, preserving the oracle parity the reader defends.
    /// </summary>
    private static void RenderTextboxDiffs(
        IrNodeList<IrTextboxDiff> textboxDiffs, in Context ctx, List<IrRevision> sink)
    {
        if (ctx.Settings.RevisionGranularity != RevisionGranularity.WmlComparerCompatible)
        {
            foreach (var tbxDiff in textboxDiffs)
                foreach (var blockOp in tbxDiff.Ops)
                    RenderBlockOp(blockOp, ctx, sink);
            return;
        }

        // Render each textbox diff to its own revision batch first, so we can compare adjacent batches.
        var batches = new List<List<IrRevision>>(textboxDiffs.Count);
        foreach (var tbxDiff in textboxDiffs)
        {
            var batch = new List<IrRevision>();
            foreach (var blockOp in tbxDiff.Ops)
                RenderTextboxInnerOp(blockOp, ctx, batch);
            batches.Add(batch);
        }

        // Collapse the Choice/Fallback DUPLICATE. Word emits each logical textbox TWICE — a DrawingML
        // mc:Choice and a VML mc:Fallback with byte-identical inner content — and WmlComparer MC-resolves its
        // input (PreProcessMarkup opens with MarkupCompatibilityProcessMode.ProcessAllParts for Office2007,
        // discarding the branch it cannot satisfy), so the oracle sees ONE branch and reports the change once.
        // The IR reader walks BOTH branches (M1.4 markdown-projection parity), so a changed textbox renders to
        // two value-equal revision batches that we must collapse to mirror the oracle's count.
        //
        // We dedup by CONTENT-SIGNATURE OCCURRENCE PARITY rather than the old adjacent pair-walk: within one
        // carrier paragraph each distinct batch signature from the AlternateContent duplication appears an even
        // number of times (once per branch), so we emit ODD occurrences (1st, 3rd, …) and drop EVEN ones (the
        // Fallback copy). This is robust to the branches being NON-ADJACENT, which a NESTED textbox produces:
        // WC-1900 (WC048-Text-Box-in-Cell) interleaves an empty wrapper body between the two `Textbox3` copies
        // — the order is [Textbox3, ⌀, Textbox3, ⌀], NOT [Textbox3, Textbox3] — so the old i/i+1 pair-walk
        // saw `Textbox3` and `⌀` as non-equal neighbours and never collapsed the pair (+2). Parity matching
        // collapses each signature's pair wherever it lands. For two DISTINCT textboxes (the WC037 case,
        // [A, A, B, B] or [A, B, A, B]) each of A and B still appears twice and is halved independently, and a
        // lone textbox with no fallback (one occurrence, odd) is always kept.
        var seen = new Dictionary<string, int>(System.StringComparer.Ordinal);
        foreach (var batch in batches)
        {
            string sig = BatchSignature(batch);
            int n = seen.TryGetValue(sig, out var c) ? c + 1 : 1;
            seen[sig] = n;
            if (n % 2 == 1)            // odd occurrence: the Choice copy (or a lone/unpaired textbox) — keep.
                sink.AddRange(batch);  // even occurrence: the Fallback duplicate — drop.
        }
    }

    /// <summary>A stable content signature for a rendered textbox-diff revision batch: the ordered
    /// (Type, Text) pairs joined. Empty batches share the empty signature (so an empty wrapper body pairs with
    /// its Fallback empty body). Anchors are deliberately excluded — they differ between the DrawingML Choice
    /// and VML Fallback branches, which is exactly the pair we are collapsing.</summary>
    private static string BatchSignature(List<IrRevision> batch)
    {
        if (batch.Count == 0)
            return "";
        var sb = new StringBuilder();
        foreach (var r in batch)
        {
            sb.Append((int)r.Type).Append(':').Append(r.Text).Append('\u001F');
        }
        return sb.ToString();
    }

    /// <summary>
    /// Render one textbox-INTERIOR block op in compatible mode (M2.4b Workstream C — WC-1770). WmlComparer
    /// never descends <c>w:txbxContent</c> — it treats the whole <c>w:drawing</c> as an OPAQUE atom (see
    /// WmlComparer.cs ~L8225/8673, where <c>w:drawing</c> is one comparison unit), so a CHANGED textbox
    /// surfaces as a single del+ins over the textbox paragraph's WHOLE text, never a finer interior token
    /// diff. The IR reader DOES model the textbox interior (to mirror the markdown projection), and its token
    /// diff can split an interior edit finer than the oracle (WC-1770: `In1`→`In` token-diffs to a lone
    /// Deleted `1`, where the oracle reports del `In1` + ins `In`). We reproduce the oracle's coarser
    /// whole-paragraph grain here: a textbox-interior Modified paragraph renders as one whole-block Deleted
    /// (left text) + Inserted (right text). This already-coincides for WC-1890/2080 (their interior token diff
    /// happens to be whole-paragraph) and WC-2090/2092 (interior insert/delete, no Modify) — all keep passing.
    /// Fine mode keeps the interior token diff (the engine's more precise account).
    /// </summary>
    private static void RenderTextboxInnerOp(IrEditOp op, in Context ctx, List<IrRevision> sink)
    {
        if (op.Kind == IrEditOpKind.ModifyBlock && op.TableDiff is null
            && op.LeftAnchor is not null && op.RightAnchor is not null
            && ctx.Left.AnchorIndex.TryGetValue(op.LeftAnchor, out var lb) && lb is IrParagraph
            && ctx.Right.AnchorIndex.TryGetValue(op.RightAnchor, out var rb) && rb is IrParagraph)
        {
            string delText = BlockText(op.LeftAnchor, ctx.Left, ctx.Settings);
            string insText = BlockText(op.RightAnchor, ctx.Right, ctx.Settings);
            // Only coarsen when there is actual text on both sides changing; a wholly-equal pair (no text
            // delta — should not reach here as a Modify) would otherwise emit empty revisions.
            if (delText.Length > 0)
                sink.Add(new IrRevision(IrRevisionType.Deleted, delText, ctx.Author, ctx.Date,
                    LeftAnchor: op.LeftAnchor, RightAnchor: op.RightAnchor));
            if (insText.Length > 0)
                sink.Add(new IrRevision(IrRevisionType.Inserted, insText, ctx.Author, ctx.Date,
                    LeftAnchor: op.LeftAnchor, RightAnchor: op.RightAnchor));
            return;
        }

        RenderBlockOp(op, ctx, sink);
    }


    /// <summary>
    /// True iff (compatible mode only) the moved block's text has FEWER than
    /// <see cref="IrDiffSettings.MoveMinimumTokenCount"/> whitespace-delimited words — the threshold below
    /// which WmlComparer excludes a relocation from move detection to avoid short-text false positives. In
    /// Fine mode this always returns false (the engine's move classification is reported verbatim).
    /// </summary>
    private static bool BelowMoveMinimum(string text, IrDiffSettings settings)
    {
        if (settings.RevisionGranularity != RevisionGranularity.WmlComparerCompatible)
            return false;
        int words = 0;
        bool inWord = false;
        foreach (char c in text)
        {
            if (char.IsWhiteSpace(c)) { inWord = false; }
            else if (!inWord) { inWord = true; words++; }
        }
        return words < settings.MoveMinimumTokenCount;
    }

    private static void RenderMoveOp(IrEditOp op, in Context ctx, List<IrRevision> sink)
    {
        bool isSource = op.IsMoveSource == true;
        // Source op carries the left anchor + left text; destination carries the right anchor + right text.
        string text = isSource
            ? BlockText(op.LeftAnchor, ctx.Left, ctx.Settings)
            : BlockText(op.RightAnchor, ctx.Right, ctx.Settings);

        // SDT moves are structurally lowered to delete+insert by the aligner. Keep this defensive projection
        // in lockstep with the markup renderer if a pre-existing script still reaches this path.
        var moveDoc = isSource ? ctx.Left : ctx.Right;
        var moveAnchor = isSource ? op.LeftAnchor : op.RightAnchor;
        if (moveAnchor is not null && moveDoc.AnchorIndex.TryGetValue(moveAnchor, out var moveBlock) &&
            moveBlock is IrSdtBlock)
        {
            sink.Add(isSource
                ? new IrRevision(IrRevisionType.Deleted, text, ctx.Author, ctx.Date, LeftAnchor: op.LeftAnchor)
                : new IrRevision(IrRevisionType.Inserted, text, ctx.Author, ctx.Date, RightAnchor: op.RightAnchor));
            return;
        }

        // A move is RELABELLED as Inserted+Deleted (not Moved) when either move rendering is off
        // (DetectMoves=false) OR — in compatible mode — the moved block is BELOW the minimum word count
        // WmlComparer requires for a move (very short text is excluded to avoid false positives). The IR
        // aligner's exact off-spine anchoring catches a short exact relocation as a move regardless of the
        // minimum (that gates only the fuzzy similarity pass), so the minimum is enforced here at render time.
        bool demoteToInsDel = op.RequiresWholeParagraphReplace ||
            !ctx.Settings.RenderMoves || BelowMoveMinimum(text, ctx.Settings);
        if (demoteToInsDel)
        {
            // The engine still ALIGNED this as a move; we only change how it is reported. A MoveModify
            // destination's in-move token edits are dropped: the whole destination block is reported as one
            // Inserted, exactly as WmlComparer reports a non-move insertion of relocated content.
            sink.Add(isSource
                ? new IrRevision(IrRevisionType.Deleted, text, ctx.Author, ctx.Date, LeftAnchor: op.LeftAnchor)
                : new IrRevision(IrRevisionType.Inserted, text, ctx.Author, ctx.Date, RightAnchor: op.RightAnchor));
            return;
        }

        sink.Add(new IrRevision(IrRevisionType.Moved, text, ctx.Author, ctx.Date,
            MoveGroupId: op.MoveGroupId, IsMoveSource: isSource,
            LeftAnchor: isSource ? op.LeftAnchor : null,
            RightAnchor: isSource ? null : op.RightAnchor));

        // MoveModify destination: emit the in-move token-op revisions IMMEDIATELY AFTER the destination
        // Moved revision (ordering rule: relocate, then describe the edits). The source op carries no diff.
        if (!isSource && op.Kind == IrEditOpKind.MoveModifyBlock && op.TokenDiff is { } tokenDiff)
        {
            // The destination op carries only the right anchor; its token diff's LEFT side indexes the
            // move's SOURCE block (the builder token-diffed source-vs-destination). Resolve the source
            // anchor via the pre-pass MoveGroupId map so Delete spans can recover left-token text.
            string? sourceAnchor = op.MoveGroupId is { } gid && ctx.MoveSourceAnchor.TryGetValue(gid, out var sa)
                ? sa : null;
            var leftTokens = ParagraphTokens(sourceAnchor, ctx.Left, ctx.Settings);
            var rightTokens = ParagraphTokens(op.RightAnchor, ctx.Right, ctx.Settings);
            RenderTokenOps(tokenDiff, leftTokens, rightTokens, sourceAnchor, op.RightAnchor, ctx, sink);
        }
    }

    /// <summary>
    /// Project a token diff to per-span revisions in op order. Insert→Inserted (right raw text),
    /// Delete→Deleted (left raw text), FormatChanged→one-per-uniform-sub-run, Equal→nothing.
    /// </summary>
    private static void RenderTokenOps(
        IrTokenDiff tokenDiff,
        IReadOnlyList<IrDiffToken> leftTokens, IReadOnlyList<IrDiffToken> rightTokens,
        string? leftAnchor, string? rightAnchor, in Context ctx, List<IrRevision> sink)
    {
        if (ctx.Settings.RevisionGranularity == RevisionGranularity.WmlComparerCompatible)
        {
            RenderTokenOpsCompatible(tokenDiff, leftTokens, rightTokens, leftAnchor, rightAnchor, ctx, sink);
            return;
        }

        foreach (var tokenOp in tokenDiff.Ops)
        {
            switch (tokenOp.Kind)
            {
                case IrTokenOpKind.Equal:
                    break;

                case IrTokenOpKind.Insert:
                {
                    // Suppress a masked-textbox placeholder insert: a Textbox-kind token carries no surface
                    // text and its real change is reported through the nested TextboxDiffs, so reporting it as
                    // an (empty) Inserted is a spurious double-report (prelim a, both modes). Other non-text
                    // tokens (Image/Opaque/math) are NOT suppressed — WmlComparer reports those as revisions.
                    if (IsMaskedTextboxSpan(rightTokens, tokenOp.RightStart, tokenOp.RightEnd))
                        break;
                    sink.Add(new IrRevision(IrRevisionType.Inserted,
                        RawText(rightTokens, tokenOp.RightStart, tokenOp.RightEnd), ctx.Author, ctx.Date,
                        LeftAnchor: leftAnchor, RightAnchor: rightAnchor));
                    break;
                }

                case IrTokenOpKind.Delete:
                {
                    if (IsMaskedTextboxSpan(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd))
                        break;
                    sink.Add(new IrRevision(IrRevisionType.Deleted,
                        RawText(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd), ctx.Author, ctx.Date,
                        LeftAnchor: leftAnchor, RightAnchor: rightAnchor));
                    break;
                }

                case IrTokenOpKind.FormatChanged:
                    RenderFormatChangedSpan(tokenOp, leftTokens, rightTokens, leftAnchor, rightAnchor, ctx, sink);
                    break;
            }
        }
    }

    // ------------------------------------------------------------------ compatible-mode token coalescing

    /// <summary>
    /// WmlComparer-compatible projection of a token diff (M2.4 Task 2). WmlComparer's revisions come from the
    /// produced document's contiguous <c>w:ins</c>/<c>w:del</c> regions — ONE revision per maximal contiguous
    /// changed region, separators included. We reproduce that here by walking the token-op stream and grouping
    /// consecutive Insert/Delete ops into a single changed REGION, BRIDGING across any Equal op that is purely
    /// separators (whitespace/punctuation between two changed words — part of WmlComparer's contiguous region).
    /// An Equal op carrying any Word token, or a FormatChanged op, is a true region boundary: it flushes the
    /// current region. Per region we emit at most one Deleted then one Inserted, after trimming the common
    /// char prefix/suffix the two share (WmlComparer keeps the unchanged edges). FormatChanged ops render
    /// exactly as in Fine mode (their own sub-run revisions), preserving format-change parity.
    /// </summary>
    private static void RenderTokenOpsCompatible(
        IrTokenDiff tokenDiff,
        IReadOnlyList<IrDiffToken> leftTokens, IReadOnlyList<IrDiffToken> rightTokens,
        string? leftAnchor, string? rightAnchor, in Context ctx, List<IrRevision> sink)
    {
        // Low-coverage coarsening (M2.4b Workstream B — the "coincidental Equal island" family). The aligner's
        // 1×1 gap residue (IrBlockAligner.FillGaps) pairs a near-rewritten paragraph/cell as Modified REGARDLESS
        // of similarity score (it is the only sensible reading of a lone in-gap pair). Myers then credits the
        // few COINCIDENTALLY shared words ("Video", a stray space) as Equal islands, splitting one logical
        // rewrite into several Inserted/Deleted regions — MORE revisions than WmlComparer's whole-document LCS,
        // which reports the rewrite as one contiguous del + one ins. When the Equal(+FormatChanged) token
        // coverage of the pair is BELOW the threshold, the shared islands are diff noise: we coalesce the
        // ENTIRE token stream into ONE region (bridging the word-bearing Equal ops too, not just separators) so
        // the common-affix trim recovers WmlComparer's clean whole-region del+ins. Compatible mode only — Fine
        // is the engine's truth and keeps every island. The coarse region still runs through the SAME Region
        // accumulator + word-boundary affix trim, so a wholly-common edge is still kept unchanged.
        bool coarsen = IsLowEqualCoverage(tokenDiff, leftTokens, rightTokens);

        var region = new Region();

        foreach (var tokenOp in tokenDiff.Ops)
        {
            switch (tokenOp.Kind)
            {
                case IrTokenOpKind.Equal:
                    if (coarsen && region.Open)
                    {
                        // Below the coverage floor: a word-bearing Equal island is coincidental noise, not a
                        // real boundary — bridge it into the open region exactly like a pure separator, so the
                        // whole rewrite collapses to one del+ins. (A leading Equal with no open region is still
                        // genuine unchanged head content and is dropped, becoming the trimmed common prefix.)
                        region.HoldSeparator(
                            RawText(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd),
                            RawText(rightTokens, tokenOp.RightStart, tokenOp.RightEnd));
                        break;
                    }
                    if (IsPureSeparatorSpan(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd))
                    {
                        // A pure-separator Equal MIGHT bridge two changed regions. Hold it; commit only if a
                        // changed op follows while a region is open. (If no region is open, a leading
                        // pure-separator Equal is just unchanged content.)
                        if (region.Open)
                            region.HoldSeparator(
                                RawText(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd),
                                RawText(rightTokens, tokenOp.RightStart, tokenOp.RightEnd));
                    }
                    else
                    {
                        // An Equal op bearing a Word token is a true boundary — flush the open region.
                        region.Flush(leftAnchor, rightAnchor, ctx, sink);
                    }
                    break;

                case IrTokenOpKind.Insert:
                    // A masked-textbox placeholder insert is reported through the nested TextboxDiffs, never as
                    // a token revision — skip it entirely (it neither opens a region nor contributes text).
                    if (IsMaskedTextboxSpan(rightTokens, tokenOp.RightStart, tokenOp.RightEnd))
                        break;
                    region.AddInsert(RawText(rightTokens, tokenOp.RightStart, tokenOp.RightEnd));
                    break;

                case IrTokenOpKind.Delete:
                    if (IsMaskedTextboxSpan(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd))
                        break;
                    region.AddDelete(RawText(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd));
                    break;

                case IrTokenOpKind.FormatChanged:
                    // FormatChanged is a region boundary: flush, then emit the format revisions as in Fine mode.
                    region.Flush(leftAnchor, rightAnchor, ctx, sink);
                    RenderFormatChangedSpan(tokenOp, leftTokens, rightTokens, leftAnchor, rightAnchor, ctx, sink);
                    break;
            }
        }

        region.Flush(leftAnchor, rightAnchor, ctx, sink);
    }

    /// <summary>
    /// The Equal+FormatChanged content-token-coverage ceiling below which a Modified pair is treated as a
    /// near-rewrite and coalesced to one whole-region del+ins (the "coincidental Equal island" coarsening). A
    /// pair is coarsened only when the LARGER-covered side shares less than this fraction of its content (so a
    /// paragraph that is mostly unchanged on EITHER side keeps its fine islands — that is a real in-place edit,
    /// not a rewrite). Derived empirically from the M2.4b Workstream B scoreboard sweep: the true rewrites have
    /// max-side coverage at or below ~0.50 (WC-1170 at 0.50: 2-token left vs 36-token right; WC-1950 at 0.41:
    /// only coincidental function words "the"/"of"/"each" shared) while every legitimately-finer pair sits at
    /// or above ~0.73 (the WC-1420/1430 math runs, WC-1930's 0.91/0.94 short edits). 0.67 separates them.
    /// Swept over the corpus: the result is a stable plateau across floor 0.55–0.72 × min 6–10.
    /// </summary>
    private const double LowCoverageFloor = 0.67;

    /// <summary>
    /// The minimum content-token size (on the LARGER side) a Modified pair must have to be eligible for
    /// low-coverage coarsening. A SHORT pair (a 3-word cell, a one-word run) where one word coincidentally
    /// survives reads as low coverage (WC-1930's "designs that compleme" → "Designs that complement.", 3 tokens
    /// with 1 shared = 0.33) but WmlComparer's LCS still reports it at fine word grain — coalescing it to one
    /// del+ins UNDER-reports. The rewrites this coarsening targets are substantial (WC-1170 at 36 tokens,
    /// WC-1950 at 19); requiring at least this many tokens excludes the short edits without losing a rewrite.
    /// </summary>
    private const int MinCoarsenContent = 8;

    /// <summary>
    /// True iff the Modified pair is a near-rewrite eligible for whole-region coarsening: its
    /// Equal+FormatChanged CONTENT-token coverage (measured on the larger-covered side) is below
    /// <see cref="LowCoverageFloor"/> AND the larger side has at least <see cref="MinCoarsenContent"/> content
    /// tokens. Separator/placeholder tokens are excluded from the counts (they are diff noise, not content);
    /// a side with no content tokens is never coarsened.
    /// </summary>
    private static bool IsLowEqualCoverage(
        IrTokenDiff tokenDiff,
        IReadOnlyList<IrDiffToken> leftTokens, IReadOnlyList<IrDiffToken> rightTokens)
    {
        int leftContent = CountContent(leftTokens), rightContent = CountContent(rightTokens);
        if (leftContent == 0 || rightContent == 0)
            return false;
        if (System.Math.Max(leftContent, rightContent) < MinCoarsenContent)
            return false;

        int coveredLeft = 0, coveredRight = 0;
        foreach (var op in tokenDiff.Ops)
        {
            if (op.Kind is not (IrTokenOpKind.Equal or IrTokenOpKind.FormatChanged))
                continue;
            coveredLeft += CountContent(leftTokens, op.LeftStart, op.LeftEnd);
            coveredRight += CountContent(rightTokens, op.RightStart, op.RightEnd);
        }

        double covL = (double)coveredLeft / leftContent;
        double covR = (double)coveredRight / rightContent;
        return System.Math.Max(covL, covR) < LowCoverageFloor;
    }

    /// <summary>Count <see cref="IrDiffTokenKind.Word"/>/Image/Opaque/Math content tokens in a list (separators
    /// and masked textbox placeholders excluded — they are not content the coverage ratio should weigh).</summary>
    private static int CountContent(IReadOnlyList<IrDiffToken> tokens) =>
        CountContent(tokens, 0, tokens.Count);

    private static int CountContent(IReadOnlyList<IrDiffToken> tokens, int start, int end)
    {
        int n = 0;
        for (int i = start; i < end && i < tokens.Count; i++)
            if (tokens[i].Kind is not (IrDiffTokenKind.Separator or IrDiffTokenKind.Textbox))
                n++;
        return n;
    }

    /// <summary>
    /// A mutable accumulator for ONE contiguous WmlComparer-style changed region: the deleted/inserted text,
    /// which sides were touched (so a real-but-textless edit like a math/image token still emits a revision),
    /// and a held pure-separator Equal awaiting a bridge decision.
    /// </summary>
    private sealed class Region
    {
        private readonly StringBuilder _del = new();
        private readonly StringBuilder _ins = new();
        private bool _hadDelete;
        private bool _hadInsert;
        private string _pendingSepLeft = string.Empty;
        private string _pendingSepRight = string.Empty;
        private bool _hasPendingSep;

        /// <summary>True once any Delete/Insert op has joined this region.</summary>
        public bool Open => _hadDelete || _hadInsert;

        /// <summary>Hold a pure-separator Equal: it bridges into BOTH sides only if a same-region changed op
        /// follows (committed by the next <see cref="AddDelete"/>/<see cref="AddInsert"/>).</summary>
        public void HoldSeparator(string left, string right)
        {
            _pendingSepLeft = left;
            _pendingSepRight = right;
            _hasPendingSep = true;
        }

        public void AddDelete(string text)
        {
            CommitPendingSeparator();
            _del.Append(text);
            _hadDelete = true;
        }

        public void AddInsert(string text)
        {
            CommitPendingSeparator();
            _ins.Append(text);
            _hadInsert = true;
        }

        private void CommitPendingSeparator()
        {
            if (!_hasPendingSep)
                return;
            _del.Append(_pendingSepLeft);
            _ins.Append(_pendingSepRight);
            _hasPendingSep = false;
            _pendingSepLeft = _pendingSepRight = string.Empty;
        }

        /// <summary>
        /// Emit the open region as a Deleted (if a delete touched it) then an Inserted (if an insert touched
        /// it), then reset. A side that was touched but is textless (a math/image token) still emits — matching
        /// WmlComparer's null-text revision.
        /// </summary>
        public void Flush(string? leftAnchor, string? rightAnchor, in Context ctx, List<IrRevision> sink)
        {
            if (Open)
            {
                string delText = _del.ToString();
                string insText = _ins.ToString();

                // Word-boundary common-affix trim. When both sides carry text, WmlComparer attributes the
                // change only to the differing words and keeps the shared head/tail unchanged. We trim the
                // longest common char prefix and suffix, BACKED OFF to a word boundary so we never split a
                // word (cutting `Test`/`st` at the common `t` would mis-report `Te` — the back-off keeps the
                // whole differing word). A side that trims to empty is wholly common and emits no revision.
                // This recovers WmlComparer's grain for both whole-block degenerate diffs (`This is a test.`/
                // `This.` → ` is a test`) and trailing-word edits (`before too.`/`before.` → ` too`), without
                // touching a real isolated token region (`34`/`4`, which shares no word-boundary affix).
                bool emitByText = _hadDelete && _hadInsert && delText.Length > 0 && insText.Length > 0;
                if (emitByText)
                {
                    // Structural-only word change: when the deleted and inserted TEXT are byte-identical yet the
                    // token differ still produced a del+ins for this region, the words differ only in STRUCTURE
                    // (their MatchKeys diverge though their text matches) — the established case is an intra-word
                    // note-reference relocation (`Vi`⟨ref⟩`deo` vs contiguous `Video`, M2.5 Task 1), where the
                    // relocated reference carries no text. WmlComparer reports this as del `Video` + ins `Video`
                    // (the old word atom is gone, the reconstituted one is new — WC-1620/1630/1710/1720 body).
                    // The common-affix trim would cancel both to empty (they share their whole text), erasing a
                    // change the oracle counts; so when the two sides are textually identical, keep both whole.
                    if (!string.Equals(delText, insText, System.StringComparison.Ordinal))
                        TrimCommonWordAffixes(ref delText, ref insText, ctx.Settings);
                    if (delText.Length > 0)
                        sink.Add(new IrRevision(IrRevisionType.Deleted, delText, ctx.Author, ctx.Date,
                            LeftAnchor: leftAnchor, RightAnchor: rightAnchor));
                    if (insText.Length > 0)
                        sink.Add(new IrRevision(IrRevisionType.Inserted, insText, ctx.Author, ctx.Date,
                            LeftAnchor: leftAnchor, RightAnchor: rightAnchor));
                }
                else
                {
                    // A side that was TOUCHED emits a revision even when textless (a math/image token carries
                    // no raw text but WmlComparer still reports it as a null-text revision). Order Deleted-
                    // then-Inserted matches WmlComparer's per-region w:del-then-w:ins ordering.
                    if (_hadDelete)
                        sink.Add(new IrRevision(IrRevisionType.Deleted, delText, ctx.Author, ctx.Date,
                            LeftAnchor: leftAnchor, RightAnchor: rightAnchor));
                    if (_hadInsert)
                        sink.Add(new IrRevision(IrRevisionType.Inserted, insText, ctx.Author, ctx.Date,
                            LeftAnchor: leftAnchor, RightAnchor: rightAnchor));
                }
            }

            _del.Clear();
            _ins.Clear();
            _hadDelete = _hadInsert = false;
            _hasPendingSep = false;
            _pendingSepLeft = _pendingSepRight = string.Empty;
        }
    }

    /// <summary>
    /// Trim the longest common char prefix and then suffix shared by <paramref name="del"/> and
    /// <paramref name="ins"/>, BACKED OFF to a word boundary so neither trim cuts inside a word. A boundary
    /// is legal where the trim edge falls between a separator/whitespace char and a word char (or at string
    /// end). The prefix is trimmed first, then the suffix over the remainder, so they never overlap.
    /// </summary>
    private static void TrimCommonWordAffixes(ref string del, ref string ins, IrDiffSettings settings)
    {
        int n = System.Math.Min(del.Length, ins.Length);

        // Longest common char prefix, then back off so the cut lands on a word boundary in BOTH strings —
        // checking BOTH is essential: `Title`/`Title1` share the char prefix `Title`, but the cut after it is
        // a word boundary in `Title` (string end) yet splits the word `Title1`, so WmlComparer keeps both
        // whole (no trim). `This`/`This.` share `This` and the cut is a boundary in both, so it trims.
        int prefix = 0;
        while (prefix < n && del[prefix] == ins[prefix])
            prefix++;
        while (prefix > 0 && !(IsWordBoundaryBefore(del, prefix) && IsWordBoundaryBefore(ins, prefix)))
            prefix--;

        // Longest common char suffix over what remains after the prefix, backed off the same way.
        int remaining = n - prefix;
        int suffix = 0;
        while (suffix < remaining && del[del.Length - 1 - suffix] == ins[ins.Length - 1 - suffix])
            suffix++;
        while (suffix > 0 &&
               (!IsWordBoundaryBefore(del, del.Length - suffix) ||
                !IsWordBoundaryBefore(ins, ins.Length - suffix)))
            suffix--;

        del = del.Substring(prefix, del.Length - prefix - suffix);
        ins = ins.Substring(prefix, ins.Length - prefix - suffix);
    }

    /// <summary>
    /// True iff position <paramref name="i"/> in <paramref name="s"/> is a word boundary under
    /// <see cref="WmlComparer"/>'s EXACT atom-grouping rule (<c>GetComparisonUnitList</c>): the comparer groups
    /// per-character atoms into words, where a char is a SPLITTING char (its own isolated atom — boundaries on
    /// BOTH sides) iff it is a <see cref="IrDiffSettings.WordSeparators"/> member, a CJK ideograph, OR a
    /// non-digit-adjacent <c>.</c>/<c>,</c>; every OTHER char (letters, digits, and other punctuation such as
    /// <c>!</c>/<c>?</c>/<c>:</c>) JOINS the surrounding word. A boundary exists between <c>s[i-1]</c> and
    /// <c>s[i]</c> iff either straddling char is a splitting char.
    ///
    /// <para>The earlier rule treated EVERY non-letter-digit (including <c>!</c>) as a boundary, so it trimmed
    /// <c>test</c>/<c>test!</c> to just <c>ins !</c> — but the comparer keeps <c>test!</c> as one word and
    /// reports del <c>test</c> + ins <c>test!</c> (WC-1920). The digit-adjacency carve-out keeps <c>3.14</c> one
    /// word; the <c>.</c>/<c>,</c> isolation keeps the <c>This</c>/<c>This.</c> and <c>endnote</c>/<c>endnote.</c>
    /// trims working (WC-1710 bridges the trailing <c>.</c>).</para>
    /// </summary>
    private static bool IsWordBoundaryBefore(string s, int i)
    {
        if (i <= 0 || i >= s.Length)
            return true;
        return IsOracleSplitChar(s, i - 1) || IsOracleSplitChar(s, i);
    }

    /// <summary>True iff the char at <paramref name="pos"/> is an ISOLATED atom under WmlComparer's grouping
    /// (boundaries on both sides): a <see cref="IrDiffSettings.WordSeparators"/> member, a CJK ideograph, or a
    /// <c>.</c>/<c>,</c> that is NOT adjacent to a digit. Other chars (letters, digits, <c>!</c>/<c>?</c>/<c>:</c>)
    /// join the surrounding word.</summary>
    private static bool IsOracleSplitChar(string s, int pos)
    {
        char c = s[pos];
        if (DefaultWordSeparatorSet.Contains(c))
            return true;
        if (c >= 0x4e00 && c <= 0x9fff) // CJK ideographs (matches GetComparisonUnitList's range check)
            return true;
        if (c == '.' || c == ',')
        {
            bool prevDigit = pos > 0 && char.IsDigit(s[pos - 1]);
            bool nextDigit = pos < s.Length - 1 && char.IsDigit(s[pos + 1]);
            return !(prevDigit || nextDigit); // digit-adjacent .,/ joins (e.g. 3.14); otherwise isolated
        }
        return false;
    }

    /// <summary>The comparer's word-separator chars (the <see cref="IrDiffSettings.DefaultWordSeparators"/> set)
    /// used by the affix-trim's oracle-faithful boundary test. CJK and digit-adjacency-sensitive <c>.</c>/<c>,</c>
    /// are handled separately in <see cref="IsOracleSplitChar"/>.</summary>
    private static readonly System.Collections.Generic.HashSet<char> DefaultWordSeparatorSet =
        new(IrDiffSettings.DefaultWordSeparators);

    /// <summary>
    /// True iff every token in the half-open span is a <see cref="IrDiffTokenKind.Textbox"/> placeholder — a
    /// masked textbox whose real change is reported through the nested <c>TextboxDiffs</c>, so its token-op
    /// must NOT also surface as an (empty) Inserted/Deleted. A span mixing a textbox with surface text is NOT
    /// masked (it carries real content); a non-textbox non-text token (Image/Opaque/math) is NOT masked
    /// either — WmlComparer reports those.
    /// </summary>
    private static bool IsMaskedTextboxSpan(IReadOnlyList<IrDiffToken> tokens, int start, int end)
    {
        if (start >= end)
            return false;
        for (int i = start; i < end; i++)
            if (tokens[i].Kind != IrDiffTokenKind.Textbox)
                return false;
        return true;
    }

    /// <summary>True iff every token in the half-open span is a <see cref="IrDiffTokenKind.Separator"/> — a
    /// pure inter-word separator run that WmlComparer's contiguous region spans (no Word token to break it).</summary>
    private static bool IsPureSeparatorSpan(IReadOnlyList<IrDiffToken> tokens, int start, int end)
    {
        if (start >= end)
            return false;
        for (int i = start; i < end; i++)
            if (tokens[i].Kind != IrDiffTokenKind.Separator)
                return false;
        return true;
    }

    /// <summary>
    /// Split a FormatChanged token span into maximal sub-runs of UNIFORM (modeled-old-key, modeled-new-key)
    /// and emit one FormatChanged revision per sub-run (text = sub-run right raw text; details = that
    /// sub-run's single transition). The span is equal-length on both sides (invariant on
    /// <see cref="IrTokenOpKind.FormatChanged"/>).
    /// </summary>
    private static void RenderFormatChangedSpan(
        IrTokenOp span,
        IReadOnlyList<IrDiffToken> leftTokens, IReadOnlyList<IrDiffToken> rightTokens,
        string? leftAnchor, string? rightAnchor, in Context ctx, List<IrRevision> sink)
    {
        int len = span.RightLength;
        int runStart = 0;
        while (runStart < len)
        {
            int li0 = span.LeftStart + runStart;
            int ri0 = span.RightStart + runStart;
            string oldKey = IrModeledFormat.RunKey(leftTokens[li0].Format);
            string newKey = IrModeledFormat.RunKey(rightTokens[ri0].Format);

            int runEnd = runStart + 1;
            while (runEnd < len)
            {
                var lf = leftTokens[span.LeftStart + runEnd].Format;
                var rf = rightTokens[span.RightStart + runEnd].Format;
                if (IrModeledFormat.RunKey(lf) != oldKey || IrModeledFormat.RunKey(rf) != newKey)
                    break;
                runEnd++;
            }

            var details = IrModeledFormat.FormatChangeDetails(leftTokens[li0].Format, rightTokens[ri0].Format);
            string text = RawText(rightTokens, span.RightStart + runStart, span.RightStart + runEnd);
            sink.Add(new IrRevision(IrRevisionType.FormatChanged, text, ctx.Author, ctx.Date,
                FormatChange: details, LeftAnchor: leftAnchor, RightAnchor: rightAnchor));

            runStart = runEnd;
        }
    }

    // ------------------------------------------------------------------ format-only block

    private static void RenderFormatOnlyBlock(IrEditOp op, in Context ctx, List<IrRevision> sink)
    {
        var leftTokens = ParagraphTokens(op.LeftAnchor, ctx.Left, ctx.Settings);
        var rightTokens = ParagraphTokens(op.RightAnchor, ctx.Right, ctx.Settings);

        // Block-format-change family (2026-07-03): a pPr delta reports as ONE Paragraph-scope
        // FormatChanged revision (w:pPrChange-grade), emitted BEFORE the run-level pairs and even for an
        // empty paragraph (a blank line whose alignment changed still reports). Excluded in
        // WmlComparerCompatible mode by the scope filter in Render.
        bool paraEmitted = EmitParagraphScopeFormatChanged(op, ctx, sink);
        EmitInlineSectionFormatChanged(op, ctx, sink);   // A3: mid-doc inline sectPr change

        // Non-paragraph FormatOnly (no tokens on either side): a TABLE reports its shell changes
        // (tblPr/tblGrid/trPr — content-equal, so rows/cells pair positionally); anything else is
        // nothing describable at token grain.
        if (leftTokens.Count == 0 && rightTokens.Count == 0)
        {
            if (!paraEmitted && ctx.Settings.TrackTableFormatChanges
                && ResolveTable(op.LeftAnchor, ctx.Left) is { } lt
                && ResolveTable(op.RightAnchor, ctx.Right) is { } rt)
                EmitTableFormatOnlyShellRevisions(lt, rt, ctx, sink);
            return;
        }

        if (leftTokens.Count == rightTokens.Count)
        {
            // Positional pairing: emit a FormatChanged revision per maximal uniform sub-run of positions
            // whose modeled formats differ (same sub-run rule as a FormatChanged token span).
            int n = leftTokens.Count;
            int i = 0;
            bool emittedAny = false;
            while (i < n)
            {
                if (IrModeledFormat.RunFormatEqual(leftTokens[i].Format, rightTokens[i].Format, ctx.Settings.FormatComparison))
                {
                    i++;
                    continue;
                }
                string oldKey = IrModeledFormat.RunKey(leftTokens[i].Format);
                string newKey = IrModeledFormat.RunKey(rightTokens[i].Format);
                int j = i + 1;
                while (j < n &&
                       !IrModeledFormat.RunFormatEqual(leftTokens[j].Format, rightTokens[j].Format, ctx.Settings.FormatComparison) &&
                       IrModeledFormat.RunKey(leftTokens[j].Format) == oldKey &&
                       IrModeledFormat.RunKey(rightTokens[j].Format) == newKey)
                    j++;

                var details = IrModeledFormat.FormatChangeDetails(leftTokens[i].Format, rightTokens[i].Format);
                sink.Add(new IrRevision(IrRevisionType.FormatChanged,
                    RawText(rightTokens, i, j), ctx.Author, ctx.Date,
                    FormatChange: details, LeftAnchor: op.LeftAnchor, RightAnchor: op.RightAnchor));
                emittedAny = true;
                i = j;
            }

            // Equal token counts but every paired position is modeled-format-equal: the block-level
            // FormatOnly delta lives in UNMODELED rPr the token surface cannot describe (e.g. w:shd under
            // ModeledOnly). Still report the change as one whole-block FormatChanged with empty details, so
            // a FormatOnly op never silently vanishes from the revisions surface — unless the Paragraph-scope
            // revision above already represents this op.
            if (!emittedAny && !paraEmitted)
                EmitWholeBlockFormatChanged(op, leftTokens, rightTokens, ctx, sink);

            return;
        }

        // Fallback: counts differ (run-boundary word-split). One whole-block FormatChanged with details
        // from the first divergent position under positional scan of the shorter length.
        EmitWholeBlockFormatChanged(op, leftTokens, rightTokens, ctx, sink);
    }

    /// <summary>
    /// True when a paired paragraph op's PARAGRAPH formats differ under the settings' policy: ModeledOnly
    /// compares the modeled <see cref="IrModeledFormat.ParaKey"/>s (the delta a consumer-grade report can
    /// describe); Full compares the full <see cref="IrParaFormat"/> record including the unmodeled digest
    /// (mirroring the stored-fingerprint grade). False when block-format tracking is off (the Consolidate
    /// pipeline pin).
    /// </summary>
    private static bool ParaFormatDiffers(IrParagraph left, IrParagraph right, IrDiffSettings settings)
    {
        if (!settings.TrackParagraphFormatChanges)   // paragraph slice (B1); composite keeps this on
            return false;
        return settings.FormatComparison == IrFormatComparison.ModeledOnly
            ? IrModeledFormat.ParaKey(left.Format) != IrModeledFormat.ParaKey(right.Format)
            : !EqualityComparer<IrParaFormat?>.Default.Equals(left.Format, right.Format);
    }

    /// <summary>
    /// Emit ONE Paragraph-scope FormatChanged revision for a paired-paragraph op whose pPr differs under
    /// the policy (block-format-change family). Text = the right paragraph's full raw text (the
    /// whole-block convention); details = the modeled paragraph property delta; both anchors carried.
    /// Returns false (emitting nothing) for non-paragraph pairs, one-sided ops, or a format-equal pair.
    /// </summary>
    private static bool EmitParagraphScopeFormatChanged(IrEditOp op, in Context ctx, List<IrRevision> sink)
    {
        if (op.LeftAnchor is null || op.RightAnchor is null)
            return false;
        if (!ctx.Left.AnchorIndex.TryGetValue(op.LeftAnchor, out var lb) || lb is not IrParagraph lp)
            return false;
        if (!ctx.Right.AnchorIndex.TryGetValue(op.RightAnchor, out var rb) || rb is not IrParagraph rp)
            return false;
        if (!ParaFormatDiffers(lp, rp, ctx.Settings))
            return false;

        var rightTokens = ParagraphTokens(op.RightAnchor, ctx.Right, ctx.Settings);
        sink.Add(new IrRevision(IrRevisionType.FormatChanged,
            RawText(rightTokens, 0, rightTokens.Count), ctx.Author, ctx.Date,
            FormatChange: IrModeledFormat.ParaFormatChangeDetails(lp.Format, rp.Format),
            LeftAnchor: op.LeftAnchor, RightAnchor: op.RightAnchor));
        return true;
    }

    /// <summary>
    /// Emit a Section-scope FormatChanged revision for a paired paragraph whose inline (in-<c>pPr</c>)
    /// <c>w:sectPr</c> properties differ (block-format follow-up A3). Anchored on the paragraphs' inline
    /// section-break anchors. Gated on <c>TrackBlockFormatChanges</c>; excluded from compatible mode by the
    /// non-Run scope filter, like the trailing Section revision.
    /// </summary>
    private static void EmitInlineSectionFormatChanged(IrEditOp op, in Context ctx, List<IrRevision> sink)
    {
        if (!ctx.Settings.TrackSectionFormatChanges || op.LeftAnchor is null || op.RightAnchor is null)
            return;
        if (!ctx.Left.AnchorIndex.TryGetValue(op.LeftAnchor, out var lb) || lb is not IrParagraph lp)
            return;
        if (!ctx.Right.AnchorIndex.TryGetValue(op.RightAnchor, out var rb) || rb is not IrParagraph rp)
            return;
        if (lp.InlineSectionFormat is not { } lsf || rp.InlineSectionFormat is not { } rsf)
            return;

        var details = IrModeledFormat.SectionFormatChangeDetails(lsf, rsf);
        if (details.ChangedPropertyNames.Count == 0)
            return;
        sink.Add(new IrRevision(IrRevisionType.FormatChanged, string.Empty, ctx.Author, ctx.Date,
            FormatChange: details,
            LeftAnchor: lp.InlineSectionBreakAnchor?.ToString() ?? op.LeftAnchor,
            RightAnchor: rp.InlineSectionBreakAnchor?.ToString() ?? op.RightAnchor));
    }

    /// <summary>
    /// Emit ONE whole-block FormatChanged revision (the FormatOnly fallback): text = the right block's full
    /// raw text; details from the first position at which the per-token modeled keys diverge under a
    /// positional scan of the shorter length (or the first position present only on one side when every
    /// paired position agrees). When no token carries a modeled difference at all (the unmodeled-only
    /// block-format case), details are empty.
    /// </summary>
    private static void EmitWholeBlockFormatChanged(
        IrEditOp op, IReadOnlyList<IrDiffToken> leftTokens, IReadOnlyList<IrDiffToken> rightTokens,
        in Context ctx, List<IrRevision> sink)
    {
        int min = leftTokens.Count < rightTokens.Count ? leftTokens.Count : rightTokens.Count;
        IrRunFormat? oldFmt = null;
        IrRunFormat? newFmt = null;
        for (int i = 0; i < min; i++)
        {
            if (IrModeledFormat.RunKey(leftTokens[i].Format) != IrModeledFormat.RunKey(rightTokens[i].Format))
            {
                oldFmt = leftTokens[i].Format;
                newFmt = rightTokens[i].Format;
                break;
            }
        }
        if (oldFmt is null && newFmt is null && leftTokens.Count != rightTokens.Count)
        {
            // Every paired position agrees; the divergence is the surplus tail on one side.
            if (leftTokens.Count > rightTokens.Count)
                oldFmt = leftTokens[min].Format;
            else
                newFmt = rightTokens[min].Format;
        }

        var details = IrModeledFormat.FormatChangeDetails(oldFmt, newFmt);
        string text = RawText(rightTokens, 0, rightTokens.Count);
        sink.Add(new IrRevision(IrRevisionType.FormatChanged, text, ctx.Author, ctx.Date,
            FormatChange: details, LeftAnchor: op.LeftAnchor, RightAnchor: op.RightAnchor));
    }

    // ------------------------------------------------------------------ table recursion

    /// <summary>
    /// A table diff requires the whole-table del+ins fallback when a ModifyRow carries a LEFT-only cell.
    /// Right-only cells are native <c>w:cellIns</c> insertions only while <c>w:tblGridChange</c> is enabled;
    /// without table-shell tracking their accepted grid would not be reversible on reject.
    /// </summary>
    private static bool TableDiffNeedsWholeTableFallback(IrTableDiff td, bool trackTableFormatChanges) =>
        td.RowOps.Any(r => r.Kind == IrRowOpKind.ModifyRow && r.CellOps is { } cells
            && cells.Any(c => c.RightCellAnchor == null ||
                (!trackTableFormatChanges && c.LeftCellAnchor == null)));

    private static void RenderTableDiff(IrTableDiff tableDiff, in Context ctx, List<IrRevision> sink)
    {
        foreach (var rowOp in tableDiff.RowOps)
        {
            switch (rowOp.Kind)
            {
                case IrRowOpKind.EqualRow:
                    break;

                case IrRowOpKind.InsertRow:
                    sink.Add(new IrRevision(IrRevisionType.Inserted,
                        RowText(rowOp.RightRowAnchor, ctx.Right, ctx.Settings), ctx.Author, ctx.Date,
                        RightAnchor: rowOp.RightRowAnchor));
                    break;

                case IrRowOpKind.DeleteRow:
                    sink.Add(new IrRevision(IrRevisionType.Deleted,
                        RowText(rowOp.LeftRowAnchor, ctx.Left, ctx.Settings), ctx.Author, ctx.Date,
                        LeftAnchor: rowOp.LeftRowAnchor));
                    break;

                case IrRowOpKind.MovedRow:
                {
                    bool isSource = rowOp.IsMoveSource == true;
                    string text = isSource
                        ? RowText(rowOp.LeftRowAnchor, ctx.Left, ctx.Settings)
                        : RowText(rowOp.RightRowAnchor, ctx.Right, ctx.Settings);
                    if (!ctx.Settings.RenderMoves)
                    {
                        // DetectMoves=false: a moved row renders as a Deleted (source) + Inserted (dest) pair.
                        sink.Add(isSource
                            ? new IrRevision(IrRevisionType.Deleted, text, ctx.Author, ctx.Date, LeftAnchor: rowOp.LeftRowAnchor)
                            : new IrRevision(IrRevisionType.Inserted, text, ctx.Author, ctx.Date, RightAnchor: rowOp.RightRowAnchor));
                        break;
                    }
                    sink.Add(new IrRevision(IrRevisionType.Moved, text, ctx.Author, ctx.Date,
                        MoveGroupId: rowOp.MoveGroupId, IsMoveSource: isSource,
                        LeftAnchor: isSource ? rowOp.LeftRowAnchor : null,
                        RightAnchor: isSource ? null : rowOp.RightRowAnchor));
                    break;
                }

                case IrRowOpKind.ModifyRow:
                    if (rowOp.CellOps is { } cellOps)
                        foreach (var cellOp in cellOps)
                            if (cellOp.LeftCellAnchor == null && cellOp.RightCellAnchor is { } insertedAnchor)
                                sink.Add(new IrRevision(IrRevisionType.Inserted,
                                    CellText(insertedAnchor, ctx.Right, ctx.Settings), ctx.Author, ctx.Date,
                                    RightAnchor: insertedAnchor));
                            else if (cellOp.BlockOps is { } blockOps)
                                RenderBlockOpList(blockOps, ctx, sink);
                    break;
            }
        }
    }

    // ------------------------------------------------ table-shell format revisions (block-format family)

    /// <summary>Resolve a table by anchor; null for a missing/non-table anchor.</summary>
    private static IrTable? ResolveTable(string? anchor, IrDocument doc)
        => anchor is not null && doc.AnchorIndex.TryGetValue(anchor, out var b) ? b as IrTable : null;

    /// <summary>One digest-grade table-family FormatChanged revision (empty property dictionaries; a single
    /// changed-name token — "shell" or "grid"). Text is empty (a shell change carries no surface text).</summary>
    private static IrRevision TableShellRevision(
        IrFormatChangeScope scope, string changed, string leftAnchor, string rightAnchor, in Context ctx)
        => new IrRevision(IrRevisionType.FormatChanged, string.Empty, ctx.Author, ctx.Date,
            FormatChange: new IrFormatChangeDetails(
                EmptyProps, EmptyProps, new[] { changed }, scope),
            LeftAnchor: leftAnchor, RightAnchor: rightAnchor);

    private static readonly IReadOnlyDictionary<string, string> EmptyProps =
        new Dictionary<string, string>();

    /// <summary>Report shell changes for a content-equal (FormatOnly) table pair: tblPr/tblGrid at the table,
    /// then trPr/tcPr walking rows and cells positionally (content-equality guarantees the alignment).</summary>
    private static void EmitTableFormatOnlyShellRevisions(IrTable left, IrTable right, in Context ctx, List<IrRevision> sink)
    {
        EmitTableLevelShellRevisions(left, right, ctx, sink);

        int rn = System.Math.Min(left.Rows.Count, right.Rows.Count);
        for (int i = 0; i < rn; i++)
            EmitRowAndCellShellRevisions(left.Rows[i], right.Rows[i], null, ctx, sink);
    }

    /// <summary>Report shell changes for a Modified table pair: tblPr/tblGrid at the table, then trPr/tcPr for
    /// each Equal/Modify row op (paired by base/right row anchor — inserted/deleted rows have no old/new pair).</summary>
    private static void EmitTableModifiedShellRevisions(
        IrTable left, IrTable right, IrTableDiff diff, in Context ctx, List<IrRevision> sink)
    {
        EmitTableLevelShellRevisions(left, right, ctx, sink);

        var leftRows = new Dictionary<string, IrRow>(System.StringComparer.Ordinal);
        foreach (var r in left.Rows) leftRows[r.Anchor.ToString()] = r;
        var rightRows = new Dictionary<string, IrRow>(System.StringComparer.Ordinal);
        foreach (var r in right.Rows) rightRows[r.Anchor.ToString()] = r;

        foreach (var rowOp in diff.RowOps)
        {
            if (rowOp.Kind is not (IrRowOpKind.EqualRow or IrRowOpKind.ModifyRow))
                continue;
            if (rowOp.LeftRowAnchor is { } la && rowOp.RightRowAnchor is { } ra
                && leftRows.TryGetValue(la, out var lr) && rightRows.TryGetValue(ra, out var rr))
                EmitRowAndCellShellRevisions(lr, rr, rowOp.CellOps, ctx, sink);
        }
    }

    private static void EmitTableLevelShellRevisions(IrTable left, IrTable right, in Context ctx, List<IrRevision> sink)
    {
        if (!left.TblPrDigest.Equals(right.TblPrDigest))
            sink.Add(TableShellRevision(IrFormatChangeScope.Table, "shell",
                left.Anchor.ToString(), right.Anchor.ToString(), ctx));
        if (!left.TblGridDigest.Equals(right.TblGridDigest))
            sink.Add(TableShellRevision(IrFormatChangeScope.Table, "grid",
                left.Anchor.ToString(), right.Anchor.ToString(), ctx));
    }

    /// <summary>
    /// Emit row-shell changes plus paired-cell shell changes.  A null <paramref name="cellOps"/> is the
    /// content-equal positional case; a ModifyRow supplies its explicit monotone cell pairing so a right-only
    /// insertion cannot shift later tcPr revisions onto the wrong cell.
    /// </summary>
    private static void EmitRowAndCellShellRevisions(
        IrRow left, IrRow right, IrNodeList<IrCellOp>? cellOps, in Context ctx, List<IrRevision> sink)
    {
        // Compare the flattened trackable projections (w:trPr children only, empty ≡ absent) — the exact
        // subset the markup's w:trPrChange/w:tcPrChange attribution uses — so GetRevisions and Compare agree
        // (an empty-vs-absent-shell change is untracked in BOTH, never reported by one alone).
        if (!left.TrPrShellDigest.Equals(right.TrPrShellDigest))
            sink.Add(TableShellRevision(IrFormatChangeScope.TableRow, "shell",
                left.Anchor.ToString(), right.Anchor.ToString(), ctx));
        // w:tblPrEx (row-level table property exceptions) — its own TableRow revision with a distinct
        // changed-name so it never double-fires with the trPr "shell" revision.
        if (!left.TrPrExDigest.Equals(right.TrPrExDigest))
            sink.Add(TableShellRevision(IrFormatChangeScope.TableRow, "tblPrEx",
                left.Anchor.ToString(), right.Anchor.ToString(), ctx));

        if (cellOps is null)
        {
            int cn = System.Math.Min(left.Cells.Count, right.Cells.Count);
            for (int c = 0; c < cn; c++)
                EmitCellShellRevisionIfChanged(left.Cells[c], right.Cells[c], ctx, sink);
            return;
        }

        var leftCells = new Dictionary<string, IrCell>(System.StringComparer.Ordinal);
        foreach (var cell in left.Cells) leftCells[cell.Anchor.ToString()] = cell;
        var rightCells = new Dictionary<string, IrCell>(System.StringComparer.Ordinal);
        foreach (var cell in right.Cells) rightCells[cell.Anchor.ToString()] = cell;

        foreach (var cellOp in cellOps)
            if (cellOp.LeftCellAnchor is { } la && cellOp.RightCellAnchor is { } ra &&
                leftCells.TryGetValue(la, out var lc) && rightCells.TryGetValue(ra, out var rc))
                EmitCellShellRevisionIfChanged(lc, rc, ctx, sink);
    }

    private static void EmitCellShellRevisionIfChanged(IrCell left, IrCell right, in Context ctx, List<IrRevision> sink)
    {
        if (!left.TcPrShellDigest.Equals(right.TcPrShellDigest))
            sink.Add(TableShellRevision(IrFormatChangeScope.TableCell, "shell",
                left.Anchor.ToString(), right.Anchor.ToString(), ctx));
    }

    // ------------------------------------------------------------------ text + token helpers

    /// <summary>Tokens of a paragraph resolved by anchor; empty list for a missing/non-paragraph anchor.</summary>
    private static IReadOnlyList<IrDiffToken> ParagraphTokens(string? anchor, IrDocument doc, IrDiffSettings settings)
    {
        if (anchor is not null && doc.AnchorIndex.TryGetValue(anchor, out var block) && block is IrParagraph p)
            return IrDiffTokenizer.Tokenize(p, settings);
        return System.Array.Empty<IrDiffToken>();
    }

    /// <summary>
    /// Concatenated raw text of a block resolved by anchor: a paragraph's tokens joined, or every
    /// descendant paragraph's text for a table; empty for an unknown/opaque/section block.
    /// </summary>
    private static string BlockText(string? anchor, IrDocument doc, IrDiffSettings settings)
    {
        if (anchor is null || !doc.AnchorIndex.TryGetValue(anchor, out var block))
            return string.Empty;
        return BlockTextOf(block, settings);
    }

    private static string BlockTextOf(IrBlock block, IrDiffSettings settings)
    {
        switch (block)
        {
            case IrParagraph p:
                return ParagraphText(p, settings);
            case IrTable t:
            {
                var sb = new StringBuilder();
                foreach (var row in t.Rows)
                    AppendRowText(sb, row, settings);
                return sb.ToString();
            }
            case IrSdtBlock sdt:
            {
                var sb = new StringBuilder();
                foreach (var child in sdt.Blocks)
                    sb.Append(BlockTextOf(child, settings));
                return sb.ToString();
            }
            default:
                return string.Empty;
        }
    }

    /// <summary>Concatenated raw text of a row resolved by anchor (its cells' paragraphs).</summary>
    private static string RowText(string? anchor, IrDocument doc, IrDiffSettings settings)
    {
        if (anchor is null || !doc.AnchorIndex.TryGetValue(anchor, out var block))
        {
            // Rows are not indexed as IrBlock; resolve them by scanning the document's tables.
            return anchor is null ? string.Empty : RowTextByScan(anchor, doc, settings);
        }
        return BlockTextOf(block, settings);
    }

    /// <summary>
    /// Resolve a row anchor by scanning the body's tables (rows are not in <see cref="IrDocument.AnchorIndex"/>,
    /// which holds blocks). Deterministic document-order scan; returns empty if not found.
    /// </summary>
    private static string RowTextByScan(string anchor, IrDocument doc, IrDiffSettings settings)
    {
        if (RowTextInBlocks(anchor, doc.Body.Blocks, settings) is { } bodyText)
            return bodyText;
        // Note scopes (M2.4 Task 1): a footnote/endnote may contain a table whose rows are not block-indexed.
        foreach (var scope in doc.Footnotes.Notes.Values)
            if (RowTextInBlocks(anchor, scope.Blocks, settings) is { } t)
                return t;
        foreach (var scope in doc.Endnotes.Notes.Values)
            if (RowTextInBlocks(anchor, scope.Blocks, settings) is { } t)
                return t;
        // Header/footer scopes (2026-07-03 campaign): a story may contain a table too.
        foreach (var hf in doc.Headers.Concat(doc.Footers))
            if (RowTextInBlocks(anchor, hf.Scope.Blocks, settings) is { } t)
                return t;
        return string.Empty;
    }

    /// <summary>
    /// Resolve a cell anchor by scanning table structure.  Cells, like rows, deliberately are not entries in
    /// <see cref="IrDocument.AnchorIndex"/>; a right-only cell insertion therefore needs this scoped walk to
    /// surface its text as one granular Inserted revision.
    /// </summary>
    private static string CellText(string? anchor, IrDocument doc, IrDiffSettings settings)
    {
        if (anchor is null)
            return string.Empty;
        if (CellTextInBlocks(anchor, doc.Body.Blocks, settings) is { } bodyText)
            return bodyText;
        foreach (var scope in doc.Footnotes.Notes.Values)
            if (CellTextInBlocks(anchor, scope.Blocks, settings) is { } t)
                return t;
        foreach (var scope in doc.Endnotes.Notes.Values)
            if (CellTextInBlocks(anchor, scope.Blocks, settings) is { } t)
                return t;
        foreach (var hf in doc.Headers.Concat(doc.Footers))
            if (CellTextInBlocks(anchor, hf.Scope.Blocks, settings) is { } t)
                return t;
        return string.Empty;
    }

    private static string? CellTextInBlocks(string anchor, IrNodeList<IrBlock> blocks, IrDiffSettings settings)
    {
        foreach (var block in blocks)
            if (block is IrTable table && CellTextInTable(anchor, table, settings) is { } text)
                return text;
        return null;
    }

    private static string? CellTextInTable(string anchor, IrTable table, IrDiffSettings settings)
    {
        foreach (var row in table.Rows)
            foreach (var cell in row.Cells)
            {
                if (cell.Anchor.ToString() == anchor)
                    return CellTextOf(cell, settings);
                foreach (var block in cell.Blocks)
                    if (block is IrTable nested && CellTextInTable(anchor, nested, settings) is { } nestedText)
                        return nestedText;
            }
        return null;
    }

    private static string CellTextOf(IrCell cell, IrDiffSettings settings)
    {
        var sb = new StringBuilder();
        foreach (var block in cell.Blocks)
            sb.Append(BlockTextOf(block, settings));
        return sb.ToString();
    }

    private static string? RowTextInBlocks(string anchor, IrNodeList<IrBlock> blocks, IrDiffSettings settings)
    {
        foreach (var block in blocks)
        {
            if (block is IrTable table)
            {
                foreach (var row in table.Rows)
                {
                    if (row.Anchor.ToString() == anchor)
                    {
                        var sb = new StringBuilder();
                        AppendRowText(sb, row, settings);
                        return sb.ToString();
                    }
                }
            }
        }
        return null;
    }

    private static void AppendRowText(StringBuilder sb, IrRow row, IrDiffSettings settings)
    {
        foreach (var cell in row.Cells)
            foreach (var b in cell.Blocks)
                if (b is IrParagraph p)
                    sb.Append(ParagraphText(p, settings));
    }

    private static string ParagraphText(IrParagraph p, IrDiffSettings settings)
    {
        var tokens = IrDiffTokenizer.Tokenize(p, settings);
        return RawText(tokens, 0, tokens.Count);
    }

    /// <summary>Concatenate the raw <see cref="IrDiffToken.Text"/> over a half-open token span.</summary>
    private static string RawText(IReadOnlyList<IrDiffToken> tokens, int start, int end)
    {
        if (start >= end)
            return string.Empty;
        var sb = new StringBuilder();
        for (int i = start; i < end; i++)
            sb.Append(tokens[i].Text);
        return sb.ToString();
    }
}
