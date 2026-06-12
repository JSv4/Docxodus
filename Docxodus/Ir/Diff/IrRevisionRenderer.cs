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
        foreach (var op in script.Operations)
            RenderBlockOp(op, ctx, revisions);

        // Note scopes (M2.4 Task 1): footnotes then endnotes, in the script's deterministic note order.
        // Each note's block ops render through the SAME block-op machinery as the body — its fn/en blocks
        // are in the shared AnchorIndex, so anchor→block/token resolution works unchanged, and the note's
        // distinct fn/en anchors carry the scope context into every revision.
        if (script.NoteOps is { } noteOps)
            foreach (var noteDiff in noteOps)
                foreach (var op in noteDiff.Ops)
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
        if (settings.RevisionGranularity == RevisionGranularity.WmlComparerCompatible)
            revisions.RemoveAll(IsSectionBreakZeroWidth);

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

    // ------------------------------------------------------------------ block-op dispatch

    private static void RenderBlockOp(IrEditOp op, in Context ctx, List<IrRevision> sink)
    {
        switch (op.Kind)
        {
            case IrEditOpKind.EqualBlock:
                break;

            case IrEditOpKind.InsertBlock:
                sink.Add(new IrRevision(IrRevisionType.Inserted,
                    BlockText(op.RightAnchor, ctx.Right, ctx.Settings), ctx.Author, ctx.Date,
                    RightAnchor: op.RightAnchor));
                break;

            case IrEditOpKind.DeleteBlock:
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
        }
    }

    // ------------------------------------------------------------------ modify / move

    private static void RenderModifyBlock(IrEditOp op, in Context ctx, List<IrRevision> sink)
    {
        if (op.TableDiff is { } tableDiff)
        {
            RenderTableDiff(tableDiff, ctx, sink);
            return;
        }

        if (op.TokenDiff is { } tokenDiff)
        {
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
                RenderBlockOp(blockOp, ctx, batch);
            batches.Add(batch);
        }

        // Collapse Choice/Fallback PAIRS. Word emits each logical textbox as TWO adjacent IrTextbox inlines
        // (DrawingML Choice then VML Fallback) with byte-identical content, so they form consecutive PAIRS in
        // document order. We pair-walk: when batch[i+1] duplicates batch[i], emit batch[i] and SKIP its
        // duplicate (advance by 2); otherwise emit batch[i] alone (advance by 1, covering a lone textbox with
        // no fallback). This is robust to two DISTINCT textboxes that happen to share identical edits — they
        // are positions (Choice_A, Fallback_A, Choice_B, Fallback_B), and pair-walking dedups A and B
        // separately rather than greedily merging A's fallback with B (the WC037 two-textbox case).
        int k = 0;
        while (k < batches.Count)
        {
            sink.AddRange(batches[k]);
            bool dup = k + 1 < batches.Count && SameRevisionContent(batches[k], batches[k + 1]);
            k += dup ? 2 : 1;
        }
    }

    /// <summary>True iff two revision batches carry the same ordered (Type, Text) content — the Choice/Fallback
    /// duplicate test. Anchors are deliberately ignored (they differ between the DrawingML and VML branches).</summary>
    private static bool SameRevisionContent(List<IrRevision> a, List<IrRevision> b)
    {
        if (a.Count != b.Count || a.Count == 0)
            return false;
        for (int i = 0; i < a.Count; i++)
            if (a[i].Type != b[i].Type || !string.Equals(a[i].Text, b[i].Text, System.StringComparison.Ordinal))
                return false;
        return true;
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

        // A move is RELABELLED as Inserted+Deleted (not Moved) when either move rendering is off
        // (DetectMoves=false) OR — in compatible mode — the moved block is BELOW the minimum word count
        // WmlComparer requires for a move (very short text is excluded to avoid false positives). The IR
        // aligner's exact off-spine anchoring catches a short exact relocation as a move regardless of the
        // minimum (that gates only the fuzzy similarity pass), so the minimum is enforced here at render time.
        bool demoteToInsDel = !ctx.Settings.RenderMoves || BelowMoveMinimum(text, ctx.Settings);
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
        var region = new Region();

        foreach (var tokenOp in tokenDiff.Ops)
        {
            switch (tokenOp.Kind)
            {
                case IrTokenOpKind.Equal:
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
    /// True iff position <paramref name="i"/> in <paramref name="s"/> is a word boundary — i.e. cutting here
    /// does not split a run of word characters. A cut is legal at the string edge, or when the chars straddling
    /// the cut are not BOTH word characters (so the boundary falls next to whitespace or punctuation). This
    /// matches WmlComparer's atom granularity, where punctuation like <c>.</c> is its own atom and so a legal
    /// trim edge — note the IR <c>WordSeparators</c> set is deliberately NOT used here (it excludes ASCII
    /// <c>.</c>, which WmlComparer nonetheless atomizes).
    /// </summary>
    private static bool IsWordBoundaryBefore(string s, int i)
    {
        if (i <= 0 || i >= s.Length)
            return true;
        return !(IsWordChar(s[i - 1]) && IsWordChar(s[i]));
    }

    /// <summary>A "word" char for the affix-trim back-off: a letter or digit. Everything else (whitespace,
    /// punctuation) is a potential atom boundary.</summary>
    private static bool IsWordChar(char c) => char.IsLetterOrDigit(c);

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

        // Non-paragraph FormatOnly (no tokens on either side): nothing describable at token grain.
        if (leftTokens.Count == 0 && rightTokens.Count == 0)
            return;

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
            // a FormatOnly op never silently vanishes from the revisions surface.
            if (!emittedAny)
                EmitWholeBlockFormatChanged(op, leftTokens, rightTokens, ctx, sink);

            return;
        }

        // Fallback: counts differ (run-boundary word-split). One whole-block FormatChanged with details
        // from the first divergent position under positional scan of the shorter length.
        EmitWholeBlockFormatChanged(op, leftTokens, rightTokens, ctx, sink);
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
                            if (cellOp.BlockOps is { } blockOps)
                                foreach (var blockOp in blockOps)
                                    RenderBlockOp(blockOp, ctx, sink);
                    break;
            }
        }
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
        return string.Empty;
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
