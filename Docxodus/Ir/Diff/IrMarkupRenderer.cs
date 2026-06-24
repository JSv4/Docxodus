#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Renders an <see cref="IrEditScript"/> into a NATIVE OOXML tracked-revisions document (M2.4 Task 3 —
/// the core <c>w:ins</c>/<c>w:del</c> renderer). The output obeys the <see cref="WmlComparer"/> contract:
/// <b>accept-all-revisions yields the RIGHT document's content; reject-all yields the LEFT's</b>, proven
/// against <see cref="RevisionProcessor"/> as the round-trip oracle.
/// </summary>
/// <remarks>
/// <para><b>Signature rationale.</b> <see cref="Render"/> takes the original <see cref="WmlDocument"/>s
/// (not the already-built <see cref="IrDocument"/>s) for two reasons. (1) <b>Package base.</b> The output
/// is assembled on the LEFT document's package so styles, numbering, fonts, settings, theme, and left-side
/// media parts carry over by reuse — exactly what WmlComparer does
/// (<c>WmlComparer.ProduceDocumentWithTrackedRevisions</c> opens a clone of source1/LEFT and swaps only the
/// body, then copies the right document's missing styles/numbering). We need the live LEFT package, not
/// just its IR. (2) <b>Provenance.</b> Building runs from provenance-cloned source XML preserves ALL run
/// properties — including the UNMODELED rPr the <see cref="IrRunFormat"/> model does not capture. The
/// adapter/scoreboard reads its IRs with <c>RetainSources=false</c> (no per-node <c>Source.Element</c>), so
/// the renderer re-reads both documents internally with <c>RetainSources=true</c> + <c>RevisionView=Accept</c>
/// to obtain the accept-clean source <c>w:p</c>/<c>w:tbl</c> elements it clones from. (Reading with Accept
/// matches the IR the script was built over — the adapter reads Accept too — so anchors resolve identically.)</para>
///
/// <para><b>Why clone from provenance, split at token boundaries.</b> For a Modified paragraph we must wrap
/// only the changed runs in <c>w:ins</c>/<c>w:del</c> while leaving Equal runs untouched. The token diff
/// carries half-open CHAR spans (the tokenizer's coordinate space — counting only emitted <c>w:t</c> text,
/// with tab/break/note-ref/image/opaque/textbox each 0 wide). We walk the source paragraph's run-level
/// children mirroring the tokenizer's char advance EXACTLY, and split a run whose <c>w:t</c> text straddles a
/// span boundary — cloning the run and trimming its text — so the run's <c>w:rPr</c> (modeled AND unmodeled)
/// rides along on each fragment. This is strictly more faithful than rebuilding runs from <see cref="IrRunFormat"/>.</para>
///
/// <para><b>Revision ids.</b> A single ascending counter per <see cref="Render"/> call, starting at 1 — NO
/// static state. (The s_MaxId lesson: WmlComparer's process-global <c>s_MaxId</c> static is reset at the top
/// of every run precisely because a shared mutable static collides across concurrent/re-entrant comparisons;
/// a per-call counter sidesteps that hazard entirely.) Every <c>w:ins</c>/<c>w:del</c> — run-level and the
/// paragraph-mark markers in <c>w:pPr/w:rPr</c> — gets a unique id; author/date come from
/// <see cref="IrDiffSettings.AuthorForRevisions"/>/<see cref="IrDiffSettings.DateTimeForRevisions"/>
/// (deterministic epoch by default). The counter is an instance field on a per-call <see cref="RenderState"/>,
/// so two concurrent renders never share it.</para>
///
/// <para><b>Scope (Task 3).</b> Body paragraphs only — table/move/format/note markup is Task 4. To keep THE
/// INVARIANT holding now, every construct this task does not yet render finely falls back to a CONSERVATIVE
/// whole-block insert/delete that still round-trips:
/// <list type="bullet">
/// <item><see cref="IrEditOpKind.EqualBlock"/> → the RIGHT block's content verbatim (we pick right, not left:
/// the two are content-equal, and the right side carries the trailing-format/rsid state of the ACCEPTED
/// document, so an accept-all output matches right byte-for-runs without a re-coalesce).</item>
/// <item><see cref="IrEditOpKind.InsertBlock"/> → the right block, every run wrapped in <c>w:ins</c>, the
/// paragraph mark marked inserted (<c>w:ins</c> in <c>w:pPr/w:rPr</c>).</item>
/// <item><see cref="IrEditOpKind.DeleteBlock"/> → the left block, runs wrapped in <c>w:del</c> (<c>w:t</c>→
/// <c>w:delText</c>), the paragraph mark marked deleted (<c>w:del</c> in <c>w:pPr/w:rPr</c>).</item>
/// <item><see cref="IrEditOpKind.ModifyBlock"/> with a paragraph token diff → per-span run wrapping (the fine
/// path). Equal/FormatChanged spans → right-side runs as-is; Insert spans → <c>w:ins</c>; Delete spans →
/// <c>w:del</c>/<c>delText</c>.</item>
/// <item><see cref="IrEditOpKind.FormatOnlyBlock"/> → the right block with each run stamped a
/// <c>w:rPrChange</c> carrying the LEFT run's old <c>w:rPr</c> (Task 4): accept keeps the right formatting,
/// reject restores the left.</item>
/// <item>A TABLE ModifyBlock, a non-paragraph Modified pair, moves, and notes → a conservative whole-block
/// <c>w:del</c> of the LEFT block immediately followed by a <c>w:ins</c> of the RIGHT block. Accept keeps the
/// right (correct), reject keeps the left (correct); the text-level invariant holds. Task 4 replaces these
/// with native table/move/note markup.</item>
/// </list></para>
///
/// <para><b>FormatChanged-span markup (Task 4).</b> A <see cref="IrTokenOpKind.FormatChanged"/> span is
/// TEXT-equal on both sides but FORMAT-differing. It renders as the RIGHT-side runs (accepted-state
/// formatting), each stamped a <c>w:rPrChange</c> whose inner <c>w:rPr</c> is the LEFT run's old formatting
/// (recovered positionally from the left source run at the aligned char). Accept drops the rPrChange (keeps
/// the right format); reject swaps the run's rPr to the rPrChange's inner rPr (restores the left format). The
/// strengthened invariant compares the boundary-normalized modeled-only block format signature on BOTH sides,
/// so format round-trips, not just text.</para>
///
/// <para><b>Note scopes (Task 4).</b> <see cref="IrEditScript.NoteOps"/> are NOT rendered into footnote/
/// endnote part markup yet. The body still round-trips; note-scope markup + id uniqueness across scopes is
/// Task 4.</para>
/// </remarks>
internal static class IrMarkupRenderer
{
    /// <summary>TRANSIENT marker attribute carrying a source <c>w:hyperlink</c>'s document-order ordinal onto
    /// each emitted wrapper clone, so <see cref="CoalesceAdjacentHyperlinks"/> can rejoin ONLY the fragments of
    /// the SAME source link. In the <c>pt:</c> namespace so the body's blanket <c>pt:</c> strip would catch any
    /// stray, but the coalescer removes it explicitly before output regardless.</summary>
    private static readonly XName SourceLinkId = PtOpenXml.pt + "SourceLinkId";

    /// <summary>
    /// Render <paramref name="script"/> into a tracked-revisions <see cref="WmlDocument"/> on the LEFT
    /// document's package. <paramref name="left"/>/<paramref name="right"/> are the original documents the
    /// script was built over; <paramref name="settings"/> supplies author/date/granularity. The returned
    /// document satisfies: <c>AcceptRevisions(result)</c> content-equals <paramref name="right"/> and
    /// <c>RejectRevisions(result)</c> content-equals <paramref name="left"/> at the per-block text level.
    /// </summary>
    public static WmlDocument Render(
        IrEditScript script, WmlDocument left, WmlDocument right, IrDiffSettings settings)
    {
        ArgumentNullException.ThrowIfNull(script);
        ArgumentNullException.ThrowIfNull(left);
        ArgumentNullException.ThrowIfNull(right);
        ArgumentNullException.ThrowIfNull(settings);

        // Re-read both documents WITH provenance so we can clone source w:p/w:tbl elements. RevisionView is
        // Accept to match the IR the script was built over (the adapter reads Accept), so every block anchor
        // in the script resolves to a block in these snapshots' AnchorIndex.
        var readOpts = new IrReaderOptions { RetainSources = true, RevisionView = RevisionView.Accept };
        var irLeft = IrReader.Read(left, readOpts);
        var irRight = IrReader.Read(right, readOpts);

        var state = new RenderState(irLeft, irRight, settings);

        // Assemble the new body's block-level children (w:p / w:tbl), in script order.
        var bodyBlocks = new List<XElement>();
        foreach (var op in script.Operations)
            RenderBlockOp(op, state, bodyBlocks);

        // SimplifyMoveMarkup (Task 4): rewrite native move markup as del/ins + strip range markers, a
        // post-pass mirroring WmlComparer.SimplifyMoveMarkupToDelIns (a Word-compat workaround). Operates on
        // the assembled blocks in place before they enter the package.
        if (settings is { RenderMoves: true, SimplifyMoveMarkup: true })
            foreach (var block in bodyBlocks)
                SimplifyMoveMarkup(block);

        // Drop the assembled blocks into a clone of the LEFT package, preserving its trailing top-level
        // w:sectPr (last-section metadata). Copy the RIGHT document's missing styles/numbering for continuity
        // (mirrors WmlComparer: right-only styles/legal numbering must survive in the merged output).
        var result = new WmlDocument(left);
        using (var streamDoc = new OpenXmlMemoryStreamDocument(result))
        {
            using (var wDoc = streamDoc.GetWordprocessingDocument())
            using (var rightStream = new OpenXmlMemoryStreamDocument(right))
            using (var wDocRight = rightStream.GetWordprocessingDocument())
            {
                var main = wDoc.MainDocumentPart
                    ?? throw new DocxodusException("LEFT document has no MainDocumentPart.");
                var mainXDoc = main.GetXDocument();
                var bodyEl = mainXDoc.Root?.Element(W.body)
                    ?? throw new DocxodusException("LEFT document has no w:body.");

                // Preserve the trailing top-level sectPr (a direct child of w:body that is NOT inside a pPr).
                var trailingSectPr = bodyEl.Elements(W.sectPr).LastOrDefault();

                bodyEl.Elements().Where(e => e.Name != W.sectPr).Remove();
                // Re-add the rendered blocks BEFORE the trailing sectPr (schema: sectPr is last in body).
                if (trailingSectPr != null)
                {
                    trailingSectPr.Remove();
                    bodyEl.Add(bodyBlocks);
                    bodyEl.Add(trailingSectPr);
                }
                else
                {
                    bodyEl.Add(bodyBlocks);
                }

                // Import media referenced by RIGHT-side cloned content (image embeds on inserted/equal runs)
                // into the LEFT-based package, rewriting the cloned elements' relationship ids IN PLACE — done
                // BEFORE PutXDocument so the in-tree XElements are the live ones MoveRelatedPartsToDestination
                // mutates. Uses the same proven part-copy/fresh-rId path WmlComparer uses for inserted drawings.
                var rightMain = wDocRight.MainDocumentPart;
                ImportRightSourcedMedia(state.RightSourcedClones, main, rightMain, streamDoc, rightStream);

                // Strip ALL engine-internal pt:Unid bookkeeping attributes from the assembled body (cloned runs
                // inside ins/del wrappers carry them too; a single sweep here catches every nested occurrence).
                foreach (var attr in bodyEl.DescendantsAndSelf().Attributes()
                             .Where(a => a.Name.Namespace == PtOpenXml.pt).ToList())
                    attr.Remove();

                main.PutXDocument();

                // Note-scope markup (Task 4): apply each note's edit ops INSIDE the footnotes/endnotes parts.
                // The output package carries the LEFT notes; we rebuild each diffed note's block content from its
                // ops (same dispatch as the body — anchors resolve in the shared AnchorIndex) so accept/reject
                // round-trips note content too. Done BEFORE PutXDocument of the note parts.
                if (script.NoteOps is { Count: > 0 })
                    RenderNoteScopes(script.NoteOps, state, wDoc, main, wDocRight, settings);

                // Note-id renumber pass (M2.6 Task 1): mirror the oracle's ChangeFootnoteEndnoteReferencesToUniqueRange.
                // Walk the produced body in document order; every footnote/endnote reference gets a sequential id
                // (body-reference order, base 1), each note DEFINITION is renumbered + reordered to match, and the
                // reserved separator/continuation boilerplate notes keep their ids. Runs for EVERY render (cheap and
                // idempotent when ids already coincide) so accept-by-right-order / reject-by-left-order both hold.
                RenumberNoteIds(main, W.footnoteReference, W.footnote, W.footnotes,
                    main.FootnotesPart, wDocRight.MainDocumentPart?.FootnotesPart);
                RenumberNoteIds(main, W.endnoteReference, W.endnote, W.endnotes,
                    main.EndnotesPart, wDocRight.MainDocumentPart?.EndnotesPart);

                // Carry right-only styles + numbering into the left-based package.
                if (main.StyleDefinitionsPart != null &&
                    wDocRight.MainDocumentPart?.StyleDefinitionsPart != null)
                    WmlComparer.CopyMissingStylesFromOneDocToAnother(wDocRight, wDoc);
                WmlComparer.CopyMissingNumberingFromOneDocToAnother(wDocRight, wDoc);
            }
            return streamDoc.GetModifiedWmlDocument();
        }
    }

    /// <summary>
    /// Import media (and hyperlink/external relationships) referenced by RIGHT-sourced clones into the output's
    /// LEFT-based main part, remapping the cloned elements' relationship ids IN PLACE. Extracted from
    /// <see cref="Render"/> so the composite renderer can run the same proven import per-reviewer (each reviewer
    /// package supplies its own clones). A no-op when there are no media-bearing clones or no right main part.
    /// </summary>
    internal static void ImportRightSourcedMedia(
        IReadOnlyList<XElement> rightClones, MainDocumentPart main, MainDocumentPart? rightMain,
        OpenXmlMemoryStreamDocument leftStreamDoc, OpenXmlMemoryStreamDocument rightStreamDoc)
    {
        if (rightMain == null || rightClones.Count == 0)
            return;

        // (1) Import hyperlink/external relationships (e.g. w:hyperlink/@r:id targets) the right clones reference
        // but the left package lacks — these are NOT parts, so the part-copy path below skips them; recreate them
        // with the SAME id where free so the cloned r:id resolves.
        ImportHyperlinkAndExternalRelationships(rightClones.ToList(), main, rightMain);

        // (2) Import media PARTS (image embeds, diagram data) and remap their r:ids in place, using the stream
        // documents' own packages directly (the wrapper's package is the authoritative writable one — not the
        // reflection-based OpenXmlPackage.GetPackage()).
        var leftPkgPart = leftStreamDoc.GetPackage().GetPart(main.Uri);
        var rightPkgPart = rightStreamDoc.GetPackage().GetPart(rightMain.Uri);
        // skipHeaderFooterReferences: a right-cloned Equal block can carry an inner w:sectPr whose
        // w:headerReference/w:footerReference r:ids would otherwise drag the RIGHT's header/footer parts in
        // as P<guid> duplicates. Those scopes are not diffed; the LEFT package's parts (same r:ids — shared
        // base) are authoritative, so the cloned references already resolve there. Media (drawings) still import.
        foreach (var clone in rightClones)
            WmlComparer.MoveRelatedPartsToDestination(
                rightPkgPart, leftPkgPart, clone, skipDanglingRelationships: true,
                skipHeaderFooterReferences: true);
    }

    // ----------------------------------------------------------------- block-op dispatch

    internal static void RenderBlockOp(IrEditOp op, RenderState state, List<XElement> sink)
    {
        // A standalone trailing section-break block (a `sec:` anchor, an IrSectionBreak) is last-section page
        // METADATA, not body content. Its `w:sectPr` is a direct w:body child that must be the LAST element —
        // we preserve the LEFT package's own trailing sectPr separately, so emitting this block here would put
        // a SECOND (mis-ordered) sectPr in the body (schema-invalid). Skip it in every op kind. (Equal/Insert/
        // Delete/Modify of a section break carries no revisable text, so the body-text invariant is unaffected;
        // native section-property revision markup, w:sectPrChange, is Task 4.)
        if (IsSectionBreakOp(op, state))
            return;

        switch (op.Kind)
        {
            case IrEditOpKind.EqualBlock:
                // Content-equal: emit the RIGHT block verbatim (accepted-state continuity). In a composite render
                // an EqualBlock is base-sourced — the composite renderer points RightSource at the base for it.
                EmitVerbatim(op.RightAnchor, state.RightSource, state, sink, fromRight: true);
                break;

            case IrEditOpKind.FormatOnlyBlock:
                // Text-equal, block-format-differing. Emit the right paragraph but stamp each run with a
                // w:rPrChange carrying the LEFT run's old rPr (so reject restores the left formatting). A
                // non-paragraph FormatOnly pair (no run model) falls through to a verbatim right emit.
                EmitFormatOnlyParagraph(op, state, sink);
                break;

            case IrEditOpKind.InsertBlock:
                EmitWholeBlock(op.RightAnchor, state.RightSource, state, sink, RevKind.Ins, fromRight: true);
                break;

            case IrEditOpKind.DeleteBlock:
                EmitWholeBlock(op.LeftAnchor, state.Left, state, sink, RevKind.Del, fromRight: false);
                break;

            case IrEditOpKind.ModifyBlock:
                RenderModifyBlock(op, state, sink);
                break;

            case IrEditOpKind.MoveBlock:
            case IrEditOpKind.MoveModifyBlock:
                // When move rendering is OFF (the DetectMoves=false analogue), a move is projected as a plain
                // delete-here + insert-there pair: the SOURCE op (left anchor) emits a whole-block del, the
                // DESTINATION op (right anchor) a whole-block ins. With move rendering ON, emit NATIVE move
                // markup: source → moveFromRange + w:moveFrom; destination → moveToRange + w:moveTo (a
                // MoveModify destination nests ins/del inside the moveTo for the in-move edits). Both halves
                // share a deterministic w:name keyed by MoveGroupId.
                if (!state.Settings.RenderMoves)
                {
                    if (op.IsMoveSource == true)
                        EmitWholeBlock(op.LeftAnchor, state.Left, state, sink, RevKind.Del, fromRight: false);
                    else
                        EmitWholeBlock(op.RightAnchor, state.RightSource, state, sink, RevKind.Ins, fromRight: true);
                }
                else if (op.IsMoveSource == true)
                {
                    EmitMoveSource(op, state, sink);
                }
                else
                {
                    EmitMoveDestination(op, state, sink);
                }
                break;

            case IrEditOpKind.SplitBlock:
                RenderSplitBlock(op, state, sink);
                break;

            case IrEditOpKind.MergeBlock:
                RenderMergeBlock(op, state, sink);
                break;
        }
    }

    // ----------------------------------------------------------------- split / merge markup (M2.6)

    /// <summary>
    /// Render a 1:N paragraph split as the ANCHORED-SPLIT shape (spec §3.3): emit N paragraphs, each
    /// carrying the corresponding RIGHT member's pPr and the segment diff's run content (built by the
    /// shared <see cref="BuildTokenOpContent"/> span walk over the LEFT slice vs the member). Paragraphs
    /// 0..N-2 get an INSERTED paragraph mark (<see cref="MarkParagraphMark"/>, <see cref="RevKind.Ins"/> —
    /// the new pilcrows the split introduced); the LAST paragraph's mark is the original left pilcrow's
    /// role and stays unmarked. ACCEPT keeps the marks → the N right paragraphs. REJECT removes each
    /// inserted mark — RevisionProcessor merges a reject-removed mark's paragraph into the NEXT one — so
    /// paragraphs 0..N-1 re-fuse, and the rejected per-segment ins/del restore the LEFT slices: the
    /// single LEFT paragraph reconstructs. Slice token lists retain the source paragraph's absolute char
    /// positions, so the FULL left paragraph's <see cref="SourceRunModel"/> serves every segment.
    /// Falls back to conservative whole-block del(left)+ins(members) when a source is missing.
    /// </summary>
    private static void RenderSplitBlock(IrEditOp op, RenderState state, List<XElement> sink)
    {
        var leftPara = SourceElement(op.LeftAnchor, state.Left);
        if (leftPara == null || op.SplitMergeAnchors is not { } anchors || op.SegmentDiffs is not { } diffs
            || anchors.Count != diffs.Count)
        {
            EmitWholeBlock(op.LeftAnchor, state.Left, state, sink, RevKind.Del, fromRight: false);
            if (op.SplitMergeAnchors is { } fallbackAnchors)
                foreach (var a in fallbackAnchors)
                    EmitWholeBlock(a, state.RightSource, state, sink, RevKind.Ins, fromRight: true);
            return;
        }

        var leftRuns = new SourceRunModel(leftPara);
        var leftTokens = ParagraphTokens(op.LeftAnchor, state.Left, state.Settings);

        int offset = 0;
        for (int s = 0; s < anchors.Count; s++)
        {
            var diff = diffs[s];
            int sliceLen = SegmentSliceLength(diff, leftSide: true);
            var slice = SubTokens(leftTokens, offset, sliceLen);
            offset += sliceLen;

            var memberPara = SourceElement(anchors[s], state.RightSource);
            if (memberPara == null)
            {
                EmitWholeBlock(anchors[s], state.RightSource, state, sink, RevKind.Ins, fromRight: true);
                continue;
            }
            var memberTokens = ParagraphTokens(anchors[s], state.RightSource, state.Settings);
            var rightRuns = new SourceRunModel(memberPara);

            var newPara = new XElement(W.p);
            var rightPPr = memberPara.Element(W.pPr);
            if (rightPPr != null)
                newPara.Add(StripUnids(new XElement(rightPPr)));
            newPara.Add(BuildTokenOpContent(diff, slice, memberTokens, leftRuns, rightRuns, state));
            if (s < anchors.Count - 1)
                MarkParagraphMark(newPara, RevKind.Ins, state); // the new pilcrow (RevKind.Ins — spec §3.3 nit)
            sink.Add(newPara);
        }
    }

    /// <summary>
    /// Render an N:1 paragraph merge — the inverse mark shape: emit N paragraphs; paragraphs 0..N-2
    /// carry their LEFT member's pPr (they vanish on accept and must restore left properties on reject)
    /// plus a DELETED paragraph mark (<see cref="RevKind.Del"/>); the LAST paragraph carries the RIGHT
    /// paragraph's pPr (the accepted state). Content per paragraph comes from the stored segment diff
    /// (left-member → right-slice orientation, applied directly by the shared span walk). ACCEPT removes
    /// each deleted mark — merging every paragraph into the next — yielding the single RIGHT paragraph;
    /// REJECT restores the marks and the member content: the N LEFT paragraphs reconstruct.
    /// </summary>
    private static void RenderMergeBlock(IrEditOp op, RenderState state, List<XElement> sink)
    {
        var rightPara = SourceElement(op.RightAnchor, state.RightSource);
        if (rightPara == null || op.SplitMergeAnchors is not { } anchors || op.SegmentDiffs is not { } diffs
            || anchors.Count != diffs.Count)
        {
            if (op.SplitMergeAnchors is { } fallbackAnchors)
                foreach (var a in fallbackAnchors)
                    EmitWholeBlock(a, state.Left, state, sink, RevKind.Del, fromRight: false);
            EmitWholeBlock(op.RightAnchor, state.RightSource, state, sink, RevKind.Ins, fromRight: true);
            return;
        }

        var rightRuns = new SourceRunModel(rightPara);
        var rightTokens = ParagraphTokens(op.RightAnchor, state.RightSource, state.Settings);
        var rightPPr = rightPara.Element(W.pPr);

        int offset = 0;
        for (int m = 0; m < anchors.Count; m++)
        {
            var diff = diffs[m];
            int sliceLen = SegmentSliceLength(diff, leftSide: false);
            var slice = SubTokens(rightTokens, offset, sliceLen);
            offset += sliceLen;

            var memberPara = SourceElement(anchors[m], state.Left);
            if (memberPara == null)
            {
                EmitWholeBlock(anchors[m], state.Left, state, sink, RevKind.Del, fromRight: false);
                continue;
            }
            var memberTokens = ParagraphTokens(anchors[m], state.Left, state.Settings);
            var leftRuns = new SourceRunModel(memberPara);

            var newPara = new XElement(W.p);
            bool last = m == anchors.Count - 1;
            var pPrSource = last ? rightPPr : memberPara.Element(W.pPr);
            if (pPrSource != null)
                newPara.Add(StripUnids(new XElement(pPrSource)));
            newPara.Add(BuildTokenOpContent(diff, memberTokens, slice, leftRuns, rightRuns, state));
            if (!last)
                MarkParagraphMark(newPara, RevKind.Del, state); // the joining mark accept removes
            sink.Add(newPara);
        }
    }

    /// <summary>A split/merge segment's singular-side slice length, implicit in the diff ops (F3.3):
    /// the LEFT slice of a split diff is Σ non-Insert left lengths; the RIGHT slice of a merge diff
    /// (stored member→slice orientation) is Σ non-Delete right lengths.</summary>
    private static int SegmentSliceLength(IrTokenDiff diff, bool leftSide)
    {
        int n = 0;
        foreach (var o in diff.Ops)
        {
            if (leftSide && o.Kind != IrTokenOpKind.Insert)
                n += o.LeftEnd - o.LeftStart;
            else if (!leftSide && o.Kind != IrTokenOpKind.Delete)
                n += o.RightEnd - o.RightStart;
        }
        return n;
    }

    /// <summary>A contiguous sub-list of a token list. The tokens keep their ABSOLUTE char positions in
    /// the source paragraph, which is what lets a slice compose with the full paragraph's
    /// <see cref="SourceRunModel"/> inside <see cref="BuildTokenOpContent"/>.</summary>
    private static IReadOnlyList<IrDiffToken> SubTokens(IReadOnlyList<IrDiffToken> tokens, int offset, int len)
    {
        var list = new List<IrDiffToken>(len);
        for (int i = offset; i < offset + len && i < tokens.Count; i++)
            list.Add(tokens[i]);
        return list;
    }

    /// <summary>
    /// A Modified pair. A PARAGRAPH pair with a token diff renders finely (per-span run wrapping). Any other
    /// Modified pair (table, opaque, section break, or a paragraph that somehow lacks a token diff) falls back
    /// to a conservative whole-block del(left)+ins(right) that keeps the invariant — Task 4 refines tables.
    /// </summary>
    private static void RenderModifyBlock(IrEditOp op, RenderState state, List<XElement> sink)
    {
        bool leftIsPara = ResolveBlock(op.LeftAnchor, state.Left) is IrParagraph;
        bool rightIsPara = ResolveBlock(op.RightAnchor, state.RightSource) is IrParagraph;

        if (op.TokenDiff is { } tokenDiff && leftIsPara && rightIsPara &&
            op.TextboxDiffs is null)   // textbox-interior diffs are not finely rendered in Task 3
        {
            RenderModifiedParagraph(op, tokenDiff, state, sink);
            return;
        }

        // A Modified TABLE pair with a nested table diff renders row/cell-precise markup (Task 4).
        if (op.TableDiff is { } tableDiff &&
            ResolveBlock(op.LeftAnchor, state.Left) is IrTable &&
            ResolveBlock(op.RightAnchor, state.RightSource) is IrTable)
        {
            if (RenderModifiedTable(op, tableDiff, state, sink))
                return;
            // fall through to the conservative fallback if the fine table path bailed
        }

        // Conservative fallback: delete the left block, insert the right block. Order matters only for human
        // reading; accept→right, reject→left both hold. A missing side (shouldn't happen for Modify) is skipped.
        if (op.LeftAnchor != null)
            EmitWholeBlock(op.LeftAnchor, state.Left, state, sink, RevKind.Del, fromRight: false);
        if (op.RightAnchor != null)
            EmitWholeBlock(op.RightAnchor, state.RightSource, state, sink, RevKind.Ins, fromRight: true);
    }


    /// <summary>
    /// Render a Modified table pair from its <see cref="IrTableDiff"/> (Task 4): build the new table from the
    /// RIGHT table's shell (tblPr/tblGrid) with rows assembled per <see cref="IrRowOp"/> — EqualRow passthrough,
    /// InsertRow → <c>w:trPr/w:ins</c> + run-wrapped, DeleteRow → <c>w:trPr/w:del</c> + run-wrapped, ModifyRow →
    /// cell-precise via the nested cell/block ops. Returns false (caller falls back) if a needed source row is
    /// unresolvable — the fallback still round-trips.
    /// </summary>
    private static bool RenderModifiedTable(IrEditOp op, IrTableDiff tableDiff, RenderState state, List<XElement> sink)
    {
        var rightTbl = SourceElement(op.RightAnchor, state.RightSource);
        var leftTbl = SourceElement(op.LeftAnchor, state.Left);
        if (rightTbl == null || leftTbl == null || rightTbl.Name != W.tbl || leftTbl.Name != W.tbl)
            return false;

        // Index the source rows by anchor so a row op resolves to its source w:tr.
        var leftRowsByAnchor = IndexRows(ResolveBlock(op.LeftAnchor, state.Left) as IrTable);
        var rightRowsByAnchor = IndexRows(ResolveBlock(op.RightAnchor, state.RightSource) as IrTable);

        var newTbl = new XElement(W.tbl);
        // Carry the table's non-row prelude (tblPr, tblGrid, …) from the right shell.
        foreach (var pre in rightTbl.Elements().Where(e => e.Name != W.tr))
            newTbl.Add(StripUnids(new XElement(pre)));

        foreach (var rowOp in tableDiff.RowOps)
        {
            switch (rowOp.Kind)
            {
                case IrRowOpKind.EqualRow:
                {
                    if (!rightRowsByAnchor.TryGetValue(rowOp.RightRowAnchor ?? "", out var src)) return false;
                    var row = StripUnids(new XElement(src));
                    state.RegisterMediaReferences(row);
                    newTbl.Add(row);
                    break;
                }
                case IrRowOpKind.InsertRow:
                {
                    if (!rightRowsByAnchor.TryGetValue(rowOp.RightRowAnchor ?? "", out var src)) return false;
                    var row = StripUnids(new XElement(src));
                    state.RegisterMediaReferences(row);
                    MarkWholeRow(row, RevKind.Ins, state);
                    newTbl.Add(row);
                    break;
                }
                case IrRowOpKind.DeleteRow:
                {
                    if (!leftRowsByAnchor.TryGetValue(rowOp.LeftRowAnchor ?? "", out var src)) return false;
                    var row = StripUnids(new XElement(src));
                    MarkWholeRow(row, RevKind.Del, state);
                    newTbl.Add(row);
                    break;
                }
                case IrRowOpKind.ModifyRow:
                {
                    if (!rightRowsByAnchor.TryGetValue(rowOp.RightRowAnchor ?? "", out var rightSrc)) return false;
                    if (!RenderModifyRow(rowOp, rightSrc, state, newTbl))
                        return false;
                    break;
                }
                case IrRowOpKind.MovedRow:
                    // A relocated exact-content row: render as DeleteRow at source + InsertRow at destination
                    // (the two MovedRow ops carry the left/right anchors respectively). This keeps the content
                    // round-trip without native row-move markup (out of Task-4 scope).
                    if (rowOp.IsMoveSource == true && rowOp.LeftRowAnchor is { } lr && leftRowsByAnchor.TryGetValue(lr, out var ms))
                    {
                        var row = StripUnids(new XElement(ms));
                        MarkWholeRow(row, RevKind.Del, state);
                        newTbl.Add(row);
                    }
                    else if (rowOp.RightRowAnchor is { } rr && rightRowsByAnchor.TryGetValue(rr, out var md))
                    {
                        var row = StripUnids(new XElement(md));
                        state.RegisterMediaReferences(row);
                        MarkWholeRow(row, RevKind.Ins, state);
                        newTbl.Add(row);
                    }
                    else return false;
                    break;
            }
        }

        sink.Add(newTbl);
        return true;
    }

    /// <summary>Render a ModifyRow: build the new row from the RIGHT row shell, replacing each paired cell's
    /// content per its block ops, and whole-marking an unpaired (column-surplus) cell. Returns false to bail to
    /// the caller's fallback if the structure can't be resolved.</summary>
    private static bool RenderModifyRow(IrRowOp rowOp, XElement rightRowSrc, RenderState state, XElement newTbl)
    {
        // Without a per-cell op list, emit the right row as-is (content-equal row that fell into a ModifyRow by
        // row-property change only) — still round-trips.
        if (rowOp.CellOps == null)
        {
            var row0 = StripUnids(new XElement(rightRowSrc));
            state.RegisterMediaReferences(row0);
            newTbl.Add(row0);
            return true;
        }

        // A column add/remove (a cell op missing its left or right anchor — IrTableDiffer.DiffCells emits these
        // for surplus cells) cannot be rendered as in-place per-cell markup: a deleted column's cell op would be
        // dropped (the `ci >= rightCells.Count` cutoff below), so RejectRevisions would NOT restore the column,
        // and an added column's cell would render unmarked. Bail to the caller's whole-table del(left)+ins(right)
        // fallback, which round-trips exactly (reject ≡ left, accept ≡ right) at the cost of coarser markup —
        // the honest representation, since the per-cell renderer is column-count-stable in v1.
        foreach (var cellOp in rowOp.CellOps)
            if (cellOp.LeftCellAnchor == null || cellOp.RightCellAnchor == null)
                return false;

        var newRow = new XElement(W.tr);
        foreach (var pre in rightRowSrc.Elements().Where(e => e.Name != W.tc))
            newRow.Add(StripUnids(new XElement(pre)));

        var rightCells = rightRowSrc.Elements(W.tc).ToList();
        int ci = 0;
        foreach (var cellOp in rowOp.CellOps)
        {
            if (ci >= rightCells.Count)
                break;
            var cellSrc = rightCells[ci++];
            var newCell = new XElement(W.tc);
            foreach (var pre in cellSrc.Elements().Where(e => e.Name != W.p && e.Name != W.tbl))
                newCell.Add(StripUnids(new XElement(pre)));

            if (cellOp.BlockOps != null)
            {
                // Render the cell's block ops with the same dispatch the body uses (paragraph token diffs, etc.).
                var cellSink = new List<XElement>();
                foreach (var bop in cellOp.BlockOps)
                    RenderBlockOp(bop, state, cellSink);
                // A cell must contain at least one block-level child; if the ops produced none, keep the right
                // cell's content verbatim so the table stays schema-valid.
                if (cellSink.Count == 0)
                    foreach (var b in cellSrc.Elements().Where(e => e.Name == W.p || e.Name == W.tbl))
                        cellSink.Add(StripUnids(new XElement(b)));
                newCell.Add(cellSink);
            }
            else
            {
                foreach (var b in cellSrc.Elements().Where(e => e.Name == W.p || e.Name == W.tbl))
                    newCell.Add(StripUnids(new XElement(b)));
            }
            newRow.Add(newCell);
        }
        // Append any right cells the cell-op list did not cover (column surplus) verbatim.
        for (; ci < rightCells.Count; ci++)
            newRow.Add(StripUnids(new XElement(rightCells[ci])));

        state.RegisterMediaReferences(newRow);
        newTbl.Add(newRow);
        return true;
    }

    /// <summary>Mark a whole table row inserted/deleted: a <c>w:trPr/w:ins</c>|<c>w:del</c> marker (APPENDED in
    /// trPr — the row-revision markers are near the end of the property order) plus every paragraph in the row
    /// run-and-mark wrapped. Accept/reject then add/remove the entire row (and the empty-table cleanup drops the
    /// table if it was the last row).</summary>
    private static void MarkWholeRow(XElement tr, RevKind kind, RenderState state)
    {
        var trPr = tr.Element(W.trPr);
        if (trPr == null)
        {
            trPr = new XElement(W.trPr);
            tr.AddFirst(trPr);
        }
        trPr.Elements().Where(e => e.Name == W.ins || e.Name == W.del).Remove();
        trPr.Add(new XElement(kind == RevKind.Ins ? W.ins : W.del, state.RevisionAttributes()));
        foreach (var p in tr.Descendants(W.p).ToList())
            MarkWholeParagraph(p, kind, state);
    }

    // ----------------------------------------------------------------- composed multi-reviewer table (FOLLOW-ON B)

    /// <summary>
    /// Render a COMPOSED multi-reviewer table from <see cref="IrCompositeOp.AuthoredRows"/> (FOLLOW-ON B): a
    /// SINGLE <c>w:tbl</c> built on the BASE table's tblPr/tblGrid, with each row emitted per its
    /// <see cref="IrAuthoredRowOp"/>:
    /// <list type="bullet">
    /// <item><b>EqualRow</b> → the base row verbatim (no revision markup).</item>
    /// <item><b>InsertRow / DeleteRow</b> → swap state to the relocating reviewer and reuse the whole-row
    /// insert/delete markup (the same <see cref="MarkWholeRow"/> the two-way path uses).</item>
    /// <item><b>ModifyRow</b> → a new <c>w:tr</c> from the BASE row's trPr + base cell skeletons
    /// (count-stable; v1 clones the base cell tcPr — guaranteed by the column-structure gate). Per
    /// <see cref="IrAuthoredCellOp"/>: a base passthrough (ComposedBlockOps null) keeps the base cell content
    /// verbatim; otherwise each cell-block composite op renders into the cell sink via
    /// <paramref name="renderOneCompositeBlock"/> (the shared composite-block dispatch — this recursion handles
    /// disjoint multi-author cell paragraphs AND same-cell-paragraph token composition).</item>
    /// </list>
    /// The callback breaks the layering cycle: <see cref="IrCompositeMarkupRenderer"/> owns the composite-op
    /// dispatch and supplies it here so the two-way renderer needs no reference to the composite renderer.
    /// </summary>
    internal static void RenderComposedTable(
        IrCompositeOp op,
        IrDocument baseIr,
        IReadOnlyList<IrDocument> reviewerIrs,
        RenderState state,
        List<XElement> sink,
        Action<IrCompositeOp, IrDocument, IReadOnlyList<IrDocument>, RenderState, List<XElement>> renderOneCompositeBlock)
    {
        var authoredRows = op.AuthoredRows
            ?? throw new DocxodusException("RenderComposedTable requires op.AuthoredRows.");

        var baseTblBlock = ResolveBlock(op.Op.LeftAnchor, baseIr) as IrTable;
        var baseTbl = SourceElement(op.Op.LeftAnchor, baseIr);
        if (baseTblBlock == null || baseTbl == null || baseTbl.Name != W.tbl)
        {
            // Defensive: a composed table op should always resolve its base table. Fall back to the merged
            // diff via the single-reviewer modify path so the op is not silently dropped.
            if (op.Op.TableDiff is { } td)
            {
                var savedAuthor0 = state.AuthorOverride;
                var savedSource0 = state.RightSource;
                var savedId0 = state.RightSourceId;
                state.AuthorOverride = null;
                state.RightSource = baseIr;
                state.RightSourceId = -1;
                RenderModifyBlock(op.Op, state, sink);
                state.AuthorOverride = savedAuthor0;
                state.RightSource = savedSource0;
                state.RightSourceId = savedId0;
            }
            return;
        }

        // Base row + cell source lookups (cell anchors are NOT in AnchorIndex, so map from the base table IR).
        var baseRowsByAnchor = IndexRows(baseTblBlock);
        var baseCellsByAnchor = IndexBaseCells(baseTblBlock);

        var newTbl = new XElement(W.tbl);
        foreach (var pre in baseTbl.Elements().Where(e => e.Name != W.tr))
            newTbl.Add(StripUnids(new XElement(pre)));

        foreach (var rowOp in authoredRows)
        {
            switch (rowOp.Kind)
            {
                case IrRowOpKind.EqualRow:
                {
                    if (rowOp.BaseRowAnchor is { } ra && baseRowsByAnchor.TryGetValue(ra, out var src))
                        newTbl.Add(StripUnids(new XElement(src)));
                    break;
                }
                case IrRowOpKind.InsertRow:
                {
                    // A reviewer-inserted whole row: source it from that reviewer's table at the merged
                    // op's matching InsertRow right anchor.
                    EmitComposedInsertOrDeleteRow(op, rowOp, reviewerIrs, baseIr, state, newTbl, RevKind.Ins);
                    break;
                }
                case IrRowOpKind.DeleteRow:
                {
                    EmitComposedInsertOrDeleteRow(op, rowOp, reviewerIrs, baseIr, state, newTbl, RevKind.Del);
                    break;
                }
                case IrRowOpKind.ModifyRow:
                {
                    EmitComposedModifyRow(rowOp, baseRowsByAnchor, baseCellsByAnchor, baseIr, reviewerIrs,
                        state, newTbl, renderOneCompositeBlock);
                    break;
                }
            }
        }

        sink.Add(newTbl);
    }

    /// <summary>Emit a composed whole-row insert/delete: resolve the row's source <c>w:tr</c> (a reviewer's
    /// inserted row from the merged TableDiff's matching InsertRow right anchor, or the base row for a delete)
    /// under the relocating reviewer's state, whole-mark it, and append it.</summary>
    private static void EmitComposedInsertOrDeleteRow(
        IrCompositeOp op, IrAuthoredRowOp rowOp, IReadOnlyList<IrDocument> reviewerIrs, IrDocument baseIr,
        RenderState state, XElement newTbl, RevKind kind)
    {
        var savedAuthor = state.AuthorOverride;
        var savedSource = state.RightSource;
        var savedId = state.RightSourceId;
        try
        {
            if (kind == RevKind.Del)
            {
                // Delete: source the base row by its base anchor.
                if (rowOp.BaseRowAnchor is { } ra)
                {
                    var baseTbl = ResolveBlock(op.Op.LeftAnchor, baseIr) as IrTable;
                    var src = baseTbl?.Rows.FirstOrDefault(r => r.Anchor.ToString() == ra)?.Source.Element;
                    if (src != null)
                    {
                        var row = StripUnids(new XElement(src));
                        state.AuthorOverride = rowOp.Author;
                        state.RightSourceId = rowOp.SourceReviewer;
                        MarkWholeRow(row, RevKind.Del, state);
                        newTbl.Add(row);
                    }
                }
                return;
            }

            // Insert: source the reviewer's inserted row directly by its right anchor (carried on the authored
            // row op).
            int reviewer = rowOp.SourceReviewer;
            if (reviewer < 0 || reviewer >= reviewerIrs.Count || rowOp.RightRowAnchor is not { } rra)
                return;
            var reviewerIr = reviewerIrs[reviewer];
            var rowSrc = FindRowSource(reviewerIr, rra);
            if (rowSrc == null)
                return;
            var newRow = StripUnids(new XElement(rowSrc));
            state.AuthorOverride = rowOp.Author;
            state.RightSource = reviewerIr;
            state.RightSourceId = reviewer;
            state.RegisterMediaReferences(newRow);
            MarkWholeRow(newRow, RevKind.Ins, state);
            newTbl.Add(newRow);
        }
        finally
        {
            state.AuthorOverride = savedAuthor;
            state.RightSource = savedSource;
            state.RightSourceId = savedId;
        }
    }

    /// <summary>The source <c>w:tr</c> a row anchor resolves to in <paramref name="ir"/>, or null.</summary>
    private static XElement? FindRowSource(IrDocument ir, string rowAnchor)
    {
        foreach (var block in ir.AnchorIndex.Values)
            if (block is IrTable tbl)
                foreach (var row in tbl.Rows)
                    if (row.Anchor.ToString() == rowAnchor)
                        return row.Source.Element;
        return null;
    }

    /// <summary>Emit a composed ModifyRow: a new <c>w:tr</c> from the BASE row's trPr + per-cell content (base
    /// passthrough or per-cell-block composite render).</summary>
    private static void EmitComposedModifyRow(
        IrAuthoredRowOp rowOp,
        Dictionary<string, XElement> baseRowsByAnchor,
        Dictionary<string, XElement> baseCellsByAnchor,
        IrDocument baseIr, IReadOnlyList<IrDocument> reviewerIrs, RenderState state, XElement newTbl,
        Action<IrCompositeOp, IrDocument, IReadOnlyList<IrDocument>, RenderState, List<XElement>> renderOneCompositeBlock)
    {
        if (rowOp.BaseRowAnchor is not { } rowAnchor || !baseRowsByAnchor.TryGetValue(rowAnchor, out var baseRowSrc))
            return;

        var newRow = new XElement(W.tr);
        foreach (var pre in baseRowSrc.Elements().Where(e => e.Name != W.tc))
            newRow.Add(StripUnids(new XElement(pre)));

        if (rowOp.ComposedCells is not { } cells)
        {
            // No per-cell view: keep the base row verbatim (defensive).
            newTbl.Add(StripUnids(new XElement(baseRowSrc)));
            return;
        }

        foreach (var cellOp in cells)
        {
            XElement? baseCellSrc = cellOp.BaseCellAnchor != null
                && baseCellsByAnchor.TryGetValue(cellOp.BaseCellAnchor, out var bc) ? bc : null;
            if (baseCellSrc == null)
                continue;

            var newCell = new XElement(W.tc);
            foreach (var pre in baseCellSrc.Elements().Where(e => e.Name != W.p && e.Name != W.tbl))
                newCell.Add(StripUnids(new XElement(pre)));

            if (cellOp.ComposedBlockOps is { } blockOps)
            {
                var cellSink = new List<XElement>();
                foreach (var cellBlock in blockOps)
                    renderOneCompositeBlock(cellBlock, baseIr, reviewerIrs, state, cellSink);
                if (cellSink.Count == 0)
                    foreach (var b in baseCellSrc.Elements().Where(e => e.Name == W.p || e.Name == W.tbl))
                        cellSink.Add(StripUnids(new XElement(b)));
                newCell.Add(cellSink);
            }
            else
            {
                // Base passthrough: the base cell's content verbatim.
                foreach (var b in baseCellSrc.Elements().Where(e => e.Name == W.p || e.Name == W.tbl))
                    newCell.Add(StripUnids(new XElement(b)));
            }
            newRow.Add(newCell);
        }

        // No whole-row media registration here (unlike the two-way RenderModifyRow): each reviewer-sourced
        // cell-block already registered its own media clones under the CORRECT per-cell RightSourceId inside
        // RenderOneCompositeBlock, and base-passthrough cells reference base parts already present in the
        // output package (the assembly clones the base package). A whole-row catch-all would (a) double-register
        // those reviewer clones and (b) bucket them under whatever RightSourceId is left over after the per-cell
        // restore (typically base/-1 or an unrelated reviewer), so a cell image could be skipped or imported from
        // the WRONG reviewer's package on an r:id collision.
        newTbl.Add(newRow);
    }

    /// <summary>Map each base cell's anchor to its source <c>w:tc</c> (cells are not in AnchorIndex).</summary>
    private static Dictionary<string, XElement> IndexBaseCells(IrTable table)
    {
        var map = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var row in table.Rows)
            foreach (var cell in row.Cells)
            {
                var src = cell.Source.Element;
                if (src != null)
                    map[cell.Anchor.ToString()] = src;
            }
        return map;
    }

    /// <summary>Index a table's rows by their anchor string for source-row lookup during table rendering.</summary>
    private static Dictionary<string, XElement> IndexRows(IrTable? table)
    {
        var map = new Dictionary<string, XElement>(StringComparer.Ordinal);
        if (table == null)
            return map;
        foreach (var row in table.Rows)
        {
            var src = row.Source.Element;
            if (src != null)
                map[row.Anchor.ToString()] = src;
        }
        return map;
    }

    /// <summary>In-place rewrite of native move markup under one block to plain del/ins (mirrors
    /// <see cref="WmlComparer"/>'s <c>SimplifyMoveMarkupToDelIns</c>): <c>w:moveFrom</c> → <c>w:del</c>,
    /// <c>w:moveTo</c> → <c>w:ins</c> (attributes + children preserved), and all four range markers removed.</summary>
    internal static void SimplifyMoveMarkup(XElement block)
    {
        foreach (var moveFrom in block.DescendantsAndSelf(W.moveFrom).ToList())
            moveFrom.ReplaceWith(new XElement(W.del, moveFrom.Attributes(), moveFrom.Nodes()));
        foreach (var moveTo in block.DescendantsAndSelf(W.moveTo).ToList())
            moveTo.ReplaceWith(new XElement(W.ins, moveTo.Attributes(), moveTo.Nodes()));
        block.DescendantsAndSelf()
            .Where(e => e.Name == W.moveFromRangeStart || e.Name == W.moveFromRangeEnd ||
                        e.Name == W.moveToRangeStart || e.Name == W.moveToRangeEnd)
            .Remove();
    }

    // ----------------------------------------------------------------- note-scope markup

    /// <summary>
    /// Apply note-scope edit ops inside the footnotes/endnotes parts of the output package. For each
    /// <see cref="IrNoteDiff"/>, locate the matching <c>w:footnote</c>/<c>w:endnote</c> (by <c>@w:id</c>) in the
    /// LEFT-based part, render its block ops (reusing <see cref="RenderBlockOp"/> — note anchors resolve in the
    /// shared AnchorIndex), and replace the note's block-level children with the rendered blocks. Notes the diff
    /// did not touch are left untouched. A note id present in the diff but absent in the part is skipped (the
    /// body still round-trips).
    /// </summary>
    private static void RenderNoteScopes(
        IReadOnlyList<IrNoteDiff> noteOps, RenderState state, WordprocessingDocument wDoc, MainDocumentPart main,
        WordprocessingDocument? wDocRight, IrDiffSettings settings)
    {
        // Group note diffs by their target part so each part is loaded/saved once.
        var footnoteDiffs = noteOps.Where(n => n.Kind == IrNoteKind.Footnote).ToList();
        var endnoteDiffs = noteOps.Where(n => n.Kind == IrNoteKind.Endnote).ToList();

        var rightMain = wDocRight?.MainDocumentPart;
        ApplyNoteDiffsToPart(footnoteDiffs, EnsureNotePart(main, isFootnote: true, rightMain),
            rightMain?.FootnotesPart, W.footnote, W.footnotes, state, settings);
        ApplyNoteDiffsToPart(endnoteDiffs, EnsureNotePart(main, isFootnote: false, rightMain),
            rightMain?.EndnotesPart, W.endnote, W.endnotes, state, settings);
    }

    /// <summary>Return the output's footnotes/endnotes part, creating an EMPTY one (with the right part's
    /// boilerplate separator/continuation notes copied so references resolve) when the LEFT package lacks it
    /// but the diff inserts notes into that scope. Returns null only if there is genuinely no such scope.</summary>
    private static OpenXmlPart? EnsureNotePart(MainDocumentPart main, bool isFootnote, MainDocumentPart? rightMain)
    {
        var existing = isFootnote ? (OpenXmlPart?)main.FootnotesPart : main.EndnotesPart;
        if (existing != null)
            return existing;
        // No part on the left. If the right side has none either, nothing to render.
        var rightPart = isFootnote ? (OpenXmlPart?)rightMain?.FootnotesPart : rightMain?.EndnotesPart;
        if (rightPart == null)
            return null;

        // Create the part and seed it with the right part's BOILERPLATE notes only (the reserved separator /
        // continuation notes, ids ≤ 0), under a fresh root — so the real inserted notes start from a clean
        // LEFT-side (empty) baseline and reject-all yields no real note content.
        var newPart = isFootnote ? (OpenXmlPart)main.AddNewPart<FootnotesPart>() : main.AddNewPart<EndnotesPart>();
        var rootName = isFootnote ? W.footnotes : W.endnotes;
        var noteName = isFootnote ? W.footnote : W.endnote;
        var rightRoot = rightPart.GetXDocument().Root;
        var newRoot = new XElement(rootName,
            rightRoot?.Attributes() ?? Enumerable.Empty<XAttribute>());
        if (rightRoot != null)
            foreach (var note in rightRoot.Elements(noteName)
                         .Where(n => int.TryParse((string?)n.Attribute(W.id), out var id) && id <= 0))
                newRoot.Add(new XElement(note));
        var xDoc = newPart.GetXDocument();
        if (xDoc.Root == null)
            xDoc.Add(newRoot);
        else
            xDoc.Root.ReplaceWith(newRoot);
        newPart.PutXDocument();
        return newPart;
    }

    private static void ApplyNoteDiffsToPart(
        List<IrNoteDiff> diffs, OpenXmlPart? part, OpenXmlPart? rightPart, XName noteName, XName rootName,
        RenderState state, IrDiffSettings settings)
    {
        if (diffs.Count == 0 || part == null)
            return;
        var xDoc = part.GetXDocument();
        var root = xDoc.Root;
        if (root == null)
            return;
        var rightRoot = rightPart?.GetXDocument().Root;

        bool changed = false;
        foreach (var diff in diffs)
        {
            // M2.5 Task 3: the output part is seeded from the LEFT document, so a MATCHED note is located by its
            // LEFT id (which may differ from the right/scope id under reference-order correspondence). A
            // wholly-inserted note has no LeftNoteId and is built from the right note's shell.
            var noteEl = diff.LeftNoteId is { } lid
                ? root.Elements(noteName).FirstOrDefault(e => (string?)e.Attribute(W.id) == lid)
                : null;
            if (noteEl == null)
            {
                // The note is absent in the LEFT part (a wholly-inserted note). Create its wrapper by cloning
                // the RIGHT note element's shell (attributes + non-block prelude) so the inserted blocks land in
                // a schema-valid w:footnote/w:endnote; the ops (all InsertBlock) supply the content.
                var rightNote = rightRoot?.Elements(noteName)
                    .FirstOrDefault(e => (string?)e.Attribute(W.id) == diff.NoteId);
                if (rightNote == null)
                    continue;
                noteEl = new XElement(noteName, rightNote.Attributes());
                foreach (var pre in rightNote.Elements().Where(e => e.Name != W.p && e.Name != W.tbl))
                    noteEl.Add(StripUnids(new XElement(pre)));
                root.Add(noteEl);
            }

            // Render the note's block ops to a fresh block list (same dispatch as the body).
            var noteBlocks = new List<XElement>();
            foreach (var op in diff.Ops)
                RenderBlockOp(op, state, noteBlocks);
            if (settings is { RenderMoves: true, SimplifyMoveMarkup: true })
                foreach (var b in noteBlocks)
                    SimplifyMoveMarkup(b);

            // Strip engine-internal pt bookkeeping from the rendered blocks.
            foreach (var b in noteBlocks)
                StripUnids(b);

            // Replace the note's block-level children (w:p / w:tbl), keeping any non-block prelude.
            noteEl.Elements().Where(e => e.Name == W.p || e.Name == W.tbl).Remove();
            noteEl.Add(noteBlocks);

            // Re-id a MATCHED note's definition to its RIGHT/scope id so the definition shares an id space with
            // the body's ins/equal references (which clone from the RIGHT and carry the right id). The output part
            // was seeded from the LEFT, so a matched note still carries its LEFT id here — left and right id spaces
            // diverge whenever an inserted note shifts the numbering (WC034-After3: matched left-en#1 → right-en#2).
            // Without this, the equal body reference (right id) and its definition (left id) disagree and the
            // RenumberNoteIds pass below cannot link them. Del-only notes (no diff, left content only) keep their
            // LEFT id and are reconciled by the renumber pass via their del reference. Right ids never collide with
            // a kept left id here because matched notes move OUT of the left space and inserted notes were created
            // in the right space.
            if (diff.LeftNoteId != null && diff.NoteId != diff.LeftNoteId)
                noteEl.SetAttributeValue(W.id, diff.NoteId);
            changed = true;
        }

        if (changed)
        {
            // A note part should never carry pt bookkeeping in the output.
            foreach (var attr in root.DescendantsAndSelf().Attributes()
                         .Where(a => a.Name.Namespace == PtOpenXml.pt).ToList())
                attr.Remove();
            part.PutXDocument();
        }
    }

    // ----------------------------------------------------------------- note-id renumber (M2.6 Task 1)

    /// <summary>
    /// Renumber footnote/endnote ids in the produced package to <b>body-reference document order</b>, mirroring
    /// <see cref="WmlComparer"/>'s <c>ChangeFootnoteEndnoteReferencesToUniqueRange</c>. Walk every body reference
    /// (<paramref name="refName"/>) in document order; the n-th reference (1-based) names note ordinal <c>n</c>.
    /// Each reference's <c>@w:id</c> is rewritten to <c>n</c> and its definition (matched by side: a reference
    /// inside <c>w:del</c> resolves a LEFT-sourced definition, an <c>w:ins</c>/equal reference a RIGHT-sourced one)
    /// is renumbered to <c>n</c> and emitted in that order. Reserved separator/continuation boilerplate notes
    /// (<c>w:type</c> present, or id ≤ 0) keep their ids and lead the part. Definitions that no surviving reference
    /// names are still carried (after the renumbered ones, original order) so accept/reject — which drop the
    /// opposite side's references — never dangle: each surviving reference still resolves, and the kept ids stay an
    /// ASCENDING subsequence of the renumbered space, so the read-order note sequence matches the right document on
    /// accept and the left on reject. Idempotent when ids already coincide; runs for every render.
    /// <para><b>Known limitations (unexercised in the M2.6 corpus; documented per the T1 review).</b>
    /// (1) <i>Note-ref nested in a note body is not renumbered.</i> The reference walk scans <c>w:body</c>
    /// descendants only, so a footnote/endnote reference that lives INSIDE another note's definition body
    /// (a note that itself cites a note) keeps its original <c>@w:id</c> — only body-anchored references drive
    /// the renumber. No corpus fixture exercises note-in-note references; if one arises, the walk must also
    /// visit references inside the note part(s).
    /// (2) <i>Deleted EMPTY-bodied note dequeue keys on <c>w:delText</c>.</i> <c>IsDeletedOnly</c> classifies a
    /// definition as deleted-only via "has <c>w:delText</c> and no live <c>w:t</c>"; a deleted note whose body
    /// carries NO text at all (no <c>w:delText</c>, no <c>w:t</c>) is therefore not enqueued in <c>delDefs</c>,
    /// so a <c>w:del</c> body reference could dequeue the wrong deleted def (or none). No corpus fixture has a
    /// textless deleted note; a robust fix would key deletedness on the reference/definition correspondence the
    /// builder already records rather than on body text presence.</para>
    /// </summary>
    private static void RenumberNoteIds(MainDocumentPart main, XName refName, XName noteName, XName rootName,
        OpenXmlPart? notePart, OpenXmlPart? rightNotePart)
    {
        if (notePart == null)
            return;
        var noteXDoc = notePart.GetXDocument();
        var noteRoot = noteXDoc.Root;
        if (noteRoot == null)
            return;

        // Partition: reserved boilerplate (kept verbatim, leads the part) vs real notes (renumber candidates).
        bool IsReserved(XElement note) =>
            note.Attribute(W.type) != null ||
            (int.TryParse((string?)note.Attribute(W.id), out var nid) && nid <= 0);
        var reserved = noteRoot.Elements(noteName).Where(IsReserved).ToList();
        var realNotes = noteRoot.Elements(noteName).Where(e => !IsReserved(e)).ToList();

        // Reference walk in document order. Each reference's revision side selects which definition it names.
        var mainXDoc = main.GetXDocument();
        var bodyRefs = mainXDoc.Root?.Element(W.body)?.Descendants(refName).ToList() ?? new List<XElement>();
        if (bodyRefs.Count == 0)
        {
            // No references — nothing to renumber against; leave the part as-is.
            return;
        }

        // A definition is DELETED-ONLY (left-sourced, vanishes on accept) iff every run carrying text is inside a
        // w:del — i.e. it has w:delText and no live w:t outside a w:del. Its body reference lives in a w:del. A
        // NON-deleted definition (matched or inserted) is named by an ins/equal reference. This deletedness — NOT
        // the raw id — is the reliable side discriminator: left and right ids can collide numerically (a deleted
        // note and a matched note can BOTH land on id 1), but a del reference always names a deleted-only def and
        // an ins/equal reference always names a non-deleted def. Partitioning the defs this way mirrors the
        // oracle's disjoint left/right id ranges without needing the preprocess range trick.
        bool IsDeletedOnly(XElement note)
        {
            bool hasLiveText = note.Descendants(W.t)
                .Any(t => !t.Ancestors().Any(a => a.Name == W.del));
            bool hasDelText = note.Descendants(W.delText).Any();
            return hasDelText && !hasLiveText;
        }
        var delDefs = new Queue<XElement>(realNotes.Where(IsDeletedOnly));
        var liveById = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var note in realNotes.Where(n => !IsDeletedOnly(n)))
        {
            var id = (string?)note.Attribute(W.id);
            if (id != null) liveById[id] = note;   // last wins; ids are unique among live defs post matched-id fix
        }

        var orderedDefs = new List<XElement>();
        var assignedIdByDef = new Dictionary<XElement, string>();
        // Real notes renumber to 1..N — but a RESERVED boilerplate note can occupy a POSITIVE id (Word's
        // continuationNotice rides at id 1 in the NVCA contract), and reserved notes keep their ids and lead the
        // part. Starting the real-note counter at 1 would then re-mint id 1 for the first real note, colliding
        // with continuationNotice (a duplicate w:id on EVERY edit, even body/format-only). Start above the
        // highest positive reserved id so the renumbered range is disjoint from the kept boilerplate ids. The
        // {-1,0}-only reserved set (the synthetic corpus, and most docs) yields 1 — unchanged.
        int next = reserved
            .Select(n => int.TryParse((string?)n.Attribute(W.id), out var v) ? v : 0)
            .Where(v => v > 0)
            .DefaultIfEmpty(0)
            .Max() + 1;
        foreach (var r in bodyRefs)
        {
            var oldId = (string?)r.Attribute(W.id);
            if (oldId == null) continue;
            bool isDel = r.Ancestors().Any(a => a.Name == W.del);
            // ins/equal → the live definition with the reference's (right) id. del → the next deleted-only
            // definition (left-sourced, vanishes on accept); but a del reference whose note was NOT deleted —
            // its DEFINITION is preserved (a matched note whose only reference was deleted, so the def lingers
            // unreferenced) has no deleted-only def to consume, so fall back to the LIVE def carrying the
            // reference's id. Without the fallback the del reference gets a fresh sequential id while its
            // preserved def keeps its original id, and reject dangles (the renumbered reference resolves to no
            // definition) whenever the original id ≠ the reference's ordinal — masked by the {1,2}-in-order
            // corpus, exposed by gapped ids (e.g. the NVCA contract's 111 footnotes).
            XElement? def = isDel
                ? (delDefs.Count > 0 ? delDefs.Dequeue() : liveById.GetValueOrDefault(oldId))
                : liveById.GetValueOrDefault(oldId);

            // A note referenced more than once corresponds once: the FIRST reference fixes its id; later references
            // to the same definition reuse it (mirroring the builder's first-reference-wins correspondence).
            if (def != null && assignedIdByDef.TryGetValue(def, out var existing))
            {
                r.SetAttributeValue(W.id, existing);
                continue;
            }

            var newId = next.ToString();
            r.SetAttributeValue(W.id, newId);
            next++;
            if (def != null)
            {
                def.SetAttributeValue(W.id, newId);
                assignedIdByDef[def] = newId;
                orderedDefs.Add(def);
            }
        }

        // Carry any real definitions no surviving reference named (defensive: orphaned/unreferenced notes), after
        // the renumbered ones, preserving their relative order and existing ids.
        foreach (var note in realNotes)
            if (!assignedIdByDef.ContainsKey(note))
                orderedDefs.Add(note);

        // Rewrite the part: reserved boilerplate first, then notes in body-reference order.
        noteRoot.Elements(noteName).Remove();
        foreach (var note in reserved)
            noteRoot.Add(note);
        foreach (var note in orderedDefs)
            noteRoot.Add(note);

        main.PutXDocument();
        notePart.PutXDocument();
    }

    // ----------------------------------------------------------------- native move markup

    /// <summary>
    /// Emit the SOURCE half of a move: the LEFT paragraph bracketed by <c>w:moveFromRangeStart</c>/
    /// <c>w:moveFromRangeEnd</c> (sharing one range id + the group's <c>w:name</c>) with every run wrapped in
    /// <c>w:moveFrom</c> and the paragraph mark marked deleted. Accept removes the moved-from content (it
    /// relocated); reject restores it. Mirrors <see cref="WmlComparer"/>'s emission.
    /// </summary>
    private static void EmitMoveSource(IrEditOp op, RenderState state, List<XElement> sink)
    {
        var src = SourceElement(op.LeftAnchor, state.Left);
        if (src == null || src.Name != W.p || op.MoveGroupId is not { } gid)
        {
            // Defensive fallback: a non-paragraph or group-less move source degrades to a whole-block delete.
            EmitWholeBlock(op.LeftAnchor, state.Left, state, sink, RevKind.Del, fromRight: false);
            return;
        }
        var para = StripUnids(new XElement(src));
        MarkWholeParagraphAs(para, RevKind.MoveFrom, state);
        BracketParagraphWithMoveRange(para, isFrom: true, state.MoveName(gid), state);
        sink.Add(para);
    }

    /// <summary>
    /// Emit the DESTINATION half of a move: the RIGHT paragraph bracketed by <c>w:moveToRangeStart</c>/
    /// <c>w:moveToRangeEnd</c> with content wrapped in <c>w:moveTo</c> and the paragraph mark marked inserted.
    /// A plain <see cref="IrEditOpKind.MoveBlock"/> wraps every run in <c>w:moveTo</c>; a
    /// <see cref="IrEditOpKind.MoveModifyBlock"/> (the destination carries a token diff) renders the in-move
    /// edits as NESTED <c>w:ins</c>/<c>w:del</c> inside the moveTo range — moved-and-unchanged text in
    /// <c>w:moveTo</c>, newly-inserted text in <c>w:ins</c>, removed text in <c>w:del</c>.
    /// </summary>
    private static void EmitMoveDestination(IrEditOp op, RenderState state, List<XElement> sink)
    {
        var src = SourceElement(op.RightAnchor, state.RightSource);
        if (src == null || src.Name != W.p || op.MoveGroupId is not { } gid)
        {
            EmitWholeBlock(op.RightAnchor, state.RightSource, state, sink, RevKind.Ins, fromRight: true);
            return;
        }
        string moveName = state.MoveName(gid);

        if (op.Kind == IrEditOpKind.MoveModifyBlock && op.TokenDiff is { } tokenDiff &&
            op.TextboxDiffs is null && ResolveBlock(op.LeftAnchor, state.Left) is IrParagraph)
        {
            // Build the destination paragraph from the token diff, like RenderModifiedParagraph, but with the
            // moved-and-equal spans wrapped in w:moveTo (instead of left unwrapped) so the whole relocated
            // content vanishes on reject and appears on accept. Insert spans → w:ins, Delete spans → w:del.
            var para = BuildMoveModifyDestination(op, tokenDiff, state);
            if (para != null)
            {
                MarkParagraphMark(para, RevKind.MoveTo, state);
                BracketParagraphWithMoveRange(para, isFrom: false, moveName, state);
                sink.Add(para);
                return;
            }
        }

        var dest = StripUnids(new XElement(src));
        state.RegisterMediaReferences(dest);
        MarkWholeParagraphAs(dest, RevKind.MoveTo, state);
        BracketParagraphWithMoveRange(dest, isFrom: false, moveName, state);
        sink.Add(dest);
    }

    /// <summary>Build a MoveModify destination paragraph from its token diff: Equal/FormatChanged spans →
    /// <c>w:moveTo</c> (moved-and-unchanged), Insert spans → <c>w:ins</c>, Delete spans → <c>w:del</c>. Returns
    /// null if the source elements are unexpectedly missing (caller falls back to a plain whole-paragraph
    /// moveTo).</summary>
    private static XElement? BuildMoveModifyDestination(IrEditOp op, IrTokenDiff tokenDiff, RenderState state)
    {
        var leftPara = SourceElement(op.LeftAnchor, state.Left);
        var rightPara = SourceElement(op.RightAnchor, state.RightSource);
        if (leftPara == null || rightPara == null)
            return null;

        var leftRuns = new SourceRunModel(leftPara);
        var rightRuns = new SourceRunModel(rightPara);
        var leftTokens = ParagraphTokens(op.LeftAnchor, state.Left, state.Settings);
        var rightTokens = ParagraphTokens(op.RightAnchor, state.RightSource, state.Settings);

        var newPara = new XElement(W.p);
        var rightPPr = rightPara.Element(W.pPr);
        if (rightPPr != null)
            newPara.Add(StripUnids(new XElement(rightPPr)));

        var content = new List<XElement>();
        foreach (var tokenOp in tokenDiff.Ops)
        {
            switch (tokenOp.Kind)
            {
                case IrTokenOpKind.Equal:
                case IrTokenOpKind.FormatChanged:
                {
                    var (rs, re) = RightSpanChars(rightTokens, tokenOp);
                    var (zs, ze) = ZeroWidthBoundaries(rightTokens, tokenOp.RightStart, tokenOp.RightEnd);
                    foreach (var r in rightRuns.Slice(rs, re, zs, ze))
                        content.Add(WrapRunLevel(r, RevKind.MoveTo, state));   // moved-and-unchanged
                    break;
                }
                case IrTokenOpKind.Insert:
                {
                    var (s, e) = RightSpanChars(rightTokens, tokenOp);
                    var (zs, ze) = ZeroWidthBoundaries(rightTokens, tokenOp.RightStart, tokenOp.RightEnd);
                    foreach (var r in rightRuns.Slice(s, e, zs, ze))
                        content.Add(WrapRunLevel(r, RevKind.Ins, state));
                    break;
                }
                case IrTokenOpKind.Delete:
                {
                    var (s, e) = LeftSpanChars(leftTokens, tokenOp);
                    var (zs, ze) = ZeroWidthBoundaries(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd);
                    foreach (var r in leftRuns.Slice(s, e, zs, ze))
                        content.Add(WrapRunLevel(r, RevKind.Del, state));
                    break;
                }
            }
        }
        newPara.Add(CoalesceAdjacentHyperlinks(content));
        return newPara;
    }

    /// <summary>Wrap every run-level child of a paragraph in the given move/revision kind (like
    /// <see cref="MarkWholeParagraph"/> but kind-parameterized) and mark the paragraph mark accordingly.</summary>
    private static void MarkWholeParagraphAs(XElement para, RevKind kind, RenderState state)
    {
        var pPr = para.Element(W.pPr);
        var runChildren = para.Elements().Where(e => e.Name != W.pPr).ToList();
        foreach (var child in runChildren)
            child.Remove();
        var wrapped = runChildren.Select(c => WrapRunLevel(c, kind, state)).ToList();
        if (pPr != null)
            pPr.AddAfterSelf(wrapped);
        else
            para.AddFirst(wrapped);
        MarkParagraphMark(para, kind, state);
    }

    /// <summary>Bracket a paragraph's run-level content with a move range: insert a
    /// <c>w:moveFromRangeStart</c>/<c>w:moveToRangeStart</c> (id + name + author + date) as the first run-level
    /// child (after pPr) and the matching <c>…RangeEnd</c> (same id) as the last child.</summary>
    private static void BracketParagraphWithMoveRange(XElement para, bool isFrom, string moveName, RenderState state)
    {
        int rangeId = state.NextId();
        var startName = isFrom ? W.moveFromRangeStart : W.moveToRangeStart;
        var endName = isFrom ? W.moveFromRangeEnd : W.moveToRangeEnd;
        var start = new XElement(startName,
            new XAttribute(W.id, rangeId),
            new XAttribute(W.name, moveName),
            new XAttribute(W.author, state.AuthorOverride ?? state.Settings.AuthorForRevisions),
            new XAttribute(W.date, state.Settings.DateTimeForRevisions));
        var end = new XElement(endName, new XAttribute(W.id, rangeId));

        var pPr = para.Element(W.pPr);
        if (pPr != null)
            pPr.AddAfterSelf(start);
        else
            para.AddFirst(start);
        para.Add(end);
    }

    // ----------------------------------------------------------------- paragraph emission

    /// <summary>
    /// Emit a block (paragraph or table) verbatim — no revision markup — cloned from its source element.
    /// Right-side runs may reference right-only media; their relationship ids are remapped on import.
    /// </summary>
    private static void EmitVerbatim(
        string? anchor, IrDocument doc, RenderState state, List<XElement> sink, bool fromRight)
    {
        var src = SourceElement(anchor, doc);
        if (src == null)
            return;
        var clone = new XElement(src);
        if (fromRight)
            state.RegisterMediaReferences(clone);
        sink.Add(StripUnids(clone));
    }

    /// <summary>
    /// Emit a whole block with EVERY run wrapped as a single revision kind (insert or delete), and the
    /// paragraph mark marked correspondingly. For a TABLE, the conservative fallback wraps every leaf run in
    /// the table and marks every paragraph mark — accept/reject still resolve the whole table correctly.
    /// </summary>
    private static void EmitWholeBlock(
        string? anchor, IrDocument doc, RenderState state, List<XElement> sink, RevKind kind, bool fromRight)
    {
        var src = SourceElement(anchor, doc);
        if (src == null)
            return;
        var clone = StripUnids(new XElement(src));
        if (fromRight)
            state.RegisterMediaReferences(clone);

        if (clone.Name == W.p)
        {
            MarkWholeParagraph(clone, kind, state);
            sink.Add(clone);
        }
        else if (clone.Name == W.tbl)
        {
            MarkWholeTable(clone, kind, state);
            sink.Add(clone);
        }
        else
        {
            // Opaque/section-break block: no run model. Emit verbatim — a structural change that carries no
            // text contributes nothing to the text-level invariant, and wrapping it is neither needed nor
            // schema-safe. (Reject/accept leave it in place either way; the invariant ignores it.)
            sink.Add(clone);
        }
    }

    /// <summary>
    /// Conservative whole-table revision marking (Task-3 fallback; Task 4 emits row/cell-precise markup). Mark
    /// EVERY row inserted/deleted (<c>w:trPr/w:ins</c> or <c>w:trPr/w:del</c>) AND every contained run +
    /// paragraph mark, so accept/reject toggle the whole table cleanly: accept of an all-rows-deleted table
    /// removes every row (RevisionProcessor's <c>w:tr/w:trPr/w:del</c> → remove-row rule) and the empty table
    /// is dropped; reject of an all-rows-inserted table does the same after the ins→del reversal.
    /// </summary>
    private static void MarkWholeTable(XElement tbl, RevKind kind, RenderState state)
    {
        foreach (var tr in tbl.Elements(W.tr).ToList())
        {
            // Mark the row inserted/deleted via w:trPr/w:ins|w:del.
            var trPr = tr.Element(W.trPr);
            if (trPr == null)
            {
                trPr = new XElement(W.trPr);
                tr.AddFirst(trPr);   // trPr is the first child of tr per schema order
            }
            // In w:trPr the row-revision markers w:ins/w:del come at the END of the property order (after
            // cnfStyle/trHeight/cantSplit/…, before only w:trPrChange) — so APPEND, never AddFirst, or a
            // following w:trHeight becomes schema-invalid.
            trPr.Elements().Where(e => e.Name == W.ins || e.Name == W.del).Remove();
            trPr.Add(new XElement(kind == RevKind.Ins ? W.ins : W.del, state.RevisionAttributes()));

            // Mark every paragraph in the row's cells (runs + paragraph mark).
            foreach (var p in tr.Descendants(W.p).ToList())
                MarkWholeParagraph(p, kind, state);
        }
    }

    /// <summary>
    /// Wrap every run-level child of a paragraph in <c>w:ins</c>/<c>w:del</c> (converting <c>w:t</c>→
    /// <c>w:delText</c> for deletions) and mark the paragraph mark inserted/deleted in <c>w:pPr/w:rPr</c>.
    /// </summary>
    private static void MarkWholeParagraph(XElement para, RevKind kind, RenderState state)
    {
        var pPr = para.Element(W.pPr);
        var runChildren = para.Elements().Where(e => e.Name != W.pPr).ToList();
        foreach (var child in runChildren)
            child.Remove();

        var wrapped = new List<XElement>();
        foreach (var child in runChildren)
            wrapped.Add(WrapRunLevel(child, kind, state));

        // Re-insert wrapped runs after pPr (or at the front if no pPr).
        if (pPr != null)
            pPr.AddAfterSelf(wrapped);
        else
            para.AddFirst(wrapped);

        MarkParagraphMark(para, kind, state);
    }

    /// <summary>
    /// Wrap a single run-level element (<c>w:r</c>, <c>w:hyperlink</c>, …) in a revision element. For a
    /// deletion, <c>w:t</c> descendants become <c>w:delText</c> so the markup round-trips through
    /// <see cref="RevisionProcessor"/> (accept drops the whole <c>w:del</c>; reject swaps it to <c>w:ins</c>
    /// and <c>delText</c>→<c>t</c>).
    /// </summary>
    private static XElement WrapRunLevel(XElement runLevel, RevKind kind, RenderState state)
    {
        // A w:hyperlink (and sdt/smartTag) is NOT a valid child of w:ins/w:del — the schema requires the
        // hyperlink OUTSIDE: w:hyperlink > w:ins > w:r. So for a container, keep the wrapper and wrap its inner
        // run-level children individually. For a plain run-level element (w:r, bookmark, …) wrap it directly.
        bool insGrade = !IsDeleteGrade(kind);
        if (runLevel.Name == W.hyperlink || runLevel.Name == W.sdt || runLevel.Name == W.smartTag)
        {
            var container = new XElement(runLevel.Name, runLevel.Attributes());
            if (insGrade)
                state.RegisterMediaReferences(container);   // hyperlink r:id rides on the container element
            // Wrap every run-level CHILD (descending through a w:sdtContent wrapper); structural children
            // (e.g. sdtPr) pass through untouched.
            foreach (var child in runLevel.Elements())
                container.Add(WrapContainerChild(child, kind, state));
            return container;
        }

        var clone = new XElement(runLevel);
        if (IsDeleteGrade(kind))
            ConvertTextToDelText(clone);
        var rev = new XElement(RevElementName(kind), state.RevisionAttributes(), clone);
        if (insGrade)
            state.RegisterMediaReferences(clone);   // the cloned run is the live tree node media import remaps
        return rev;
    }

    /// <summary>
    /// Wrap a child of a run-level container (see <see cref="WrapRunLevel"/>). A run-level child
    /// (<c>w:r</c>/<c>w:hyperlink</c>/<c>w:smartTag</c>/<c>w:sdt</c>) is wrapped in the revision element; a
    /// <c>w:sdtContent</c> is PRESERVED as a wrapper and its OWN run-level children wrapped. The runs of an
    /// inline content control live under <c>w:sdtContent</c>, NOT as direct <c>w:sdt</c> children, and
    /// <c>w:ins</c>/<c>w:del</c> is a valid child of <c>w:sdtContent</c>. Without this descent an
    /// inserted/deleted <c>w:sdt</c>'s content was emitted BARE (no <c>w:ins</c>/<c>w:del</c>), so
    /// <see cref="RevisionProcessor"/> reject did not strip it — the content leaked through, breaking the
    /// <c>reject ≡ left</c> contract. Structural children (<c>w:sdtPr</c>, …) pass through untouched.
    /// </summary>
    private static XElement WrapContainerChild(XElement child, RevKind kind, RenderState state)
    {
        if (child.Name == W.r || child.Name == W.hyperlink || child.Name == W.smartTag || child.Name == W.sdt)
            return WrapRunLevel(child, kind, state);
        if (child.Name == W.sdtContent)
        {
            var content = new XElement(child.Name, child.Attributes());
            foreach (var inner in child.Elements())
                content.Add(WrapContainerChild(inner, kind, state));
            return content;
        }
        return new XElement(child);
    }

    /// <summary>Mark a paragraph's end-of-paragraph mark inserted/deleted: an EMPTY <c>w:ins</c>/<c>w:del</c>
    /// inside <c>w:pPr/w:rPr</c> (the encoding <see cref="RevisionProcessor"/> recognizes — accept of a
    /// deleted mark merges the paragraph with the following one; reject restores it). The paragraph mark
    /// supports only <c>w:ins</c>/<c>w:del</c>, so a move FROM marks the mark deleted (del-grade) and a move
    /// TO marks it inserted (ins-grade).</summary>
    private static void MarkParagraphMark(XElement para, RevKind kind, RenderState state)
    {
        var pPr = para.Element(W.pPr);
        if (pPr == null)
        {
            pPr = new XElement(W.pPr);
            para.AddFirst(pPr);
        }
        var rPr = pPr.Element(W.rPr);
        if (rPr == null)
        {
            rPr = new XElement(W.rPr);
            // The paragraph-mark w:rPr is near the END of pPr's schema order — after the paragraph-level
            // properties (pStyle, numPr, spacing, …) and before only w:sectPr / w:pPrChange. Insert it there,
            // NOT at the front (AddFirst would put it before w:pStyle, which the schema rejects).
            var sectPr = pPr.Element(W.sectPr);
            var pPrChange = pPr.Element(W.pPrChange);
            if (sectPr != null)
                sectPr.AddBeforeSelf(rPr);
            else if (pPrChange != null)
                pPrChange.AddBeforeSelf(rPr);
            else
                pPr.Add(rPr);
        }
        var markName = IsDeleteGrade(kind) ? W.del : W.ins;
        // Remove any pre-existing ins/del marker (idempotence) then add the new one FIRST inside rPr.
        rPr.Elements().Where(e => e.Name == W.ins || e.Name == W.del).Remove();
        rPr.AddFirst(new XElement(markName, state.RevisionAttributes()));
    }

    /// <summary>
    /// Emit a FormatOnly paragraph: the RIGHT paragraph's text/structure with each run stamped a
    /// <c>w:rPrChange</c> carrying the LEFT run's old <c>w:rPr</c> at the aligned char position. The two
    /// paragraphs are ContentHash-equal (same text), so the left char at offset k matches the right char at
    /// offset k and the left rPr is recoverable positionally. Accept keeps the right formatting; reject
    /// restores the left. A non-paragraph FormatOnly pair (no run model) emits the right block verbatim.
    /// </summary>
    private static void EmitFormatOnlyParagraph(IrEditOp op, RenderState state, List<XElement> sink)
    {
        var rightPara = SourceElement(op.RightAnchor, state.RightSource);
        var leftPara = SourceElement(op.LeftAnchor, state.Left);
        if (rightPara == null || rightPara.Name != W.p || leftPara == null || leftPara.Name != W.p)
        {
            EmitVerbatim(op.RightAnchor, state.RightSource, state, sink, fromRight: true);
            return;
        }

        var leftRuns = new SourceRunModel(leftPara);
        var newPara = new XElement(W.p);
        var rightPPr = rightPara.Element(W.pPr);
        if (rightPPr != null)
            newPara.Add(StripUnids(new XElement(rightPPr)));

        int cursor = 0;
        foreach (var child in rightPara.Elements().Where(e => e.Name != W.pPr))
        {
            var clone = StripUnids(new XElement(child));
            state.RegisterMediaReferences(clone);
            if (clone.Name == W.r)
            {
                var oldRPr = leftRuns.RPrAtChar(cursor);
                ApplyRPrChange(clone, oldRPr, state);
                cursor += RunTextLength(clone);
            }
            newPara.Add(clone);
        }
        sink.Add(newPara);
    }

    // ------------------------------------------------------- composite (multi-author) modify path

    /// <summary>
    /// Render ONE base paragraph edited by 2+ reviewers into a single <c>w:p</c> whose run-level content
    /// composes per-span authorship: consecutive <see cref="IrAuthoredTokenOp"/>s sharing
    /// <c>(Author, SourceReviewer)</c> are grouped, and each group is emitted via the shared
    /// <see cref="BuildTokenOpContent"/> with the contributing reviewer's right paragraph as the right source
    /// (so Insert spans, whose RightStart/RightEnd index THAT reviewer's right-token list, resolve correctly)
    /// and <see cref="RenderState.AuthorOverride"/> set to that reviewer. Base-sourced groups
    /// (<c>SourceReviewer == -1</c>, Equal spans) read the BASE paragraph for both sides with no author
    /// override. The cloned <c>pPr</c> is the BASE paragraph's (the paragraph exists on every side and the
    /// composed edits are text-only, so the base pPr is the deterministic accepted-state shape).
    /// <para><paramref name="op"/> carries the MERGED token diff in <c>op.Op.TokenDiff</c> (the apply/json
    /// truth, used by the single-reviewer path) and the per-span authorship in <c>op.AuthoredTokens</c>;
    /// <c>op.SourceRightAnchors</c> maps each contributing reviewer to its right paragraph anchor. The
    /// invariant is the multi-author generalization of the two-way contract: reject-all restores the base
    /// paragraph text; accept-all yields every reviewer's accepted word edits.</para>
    /// </summary>
    internal static void RenderComposedParagraph(
        IrCompositeOp op,
        IrDocument baseIr,
        IReadOnlyList<IrDocument> reviewerIrs,
        RenderState state,
        List<XElement> sink)
    {
        var authored = op.AuthoredTokens
            ?? throw new DocxodusException("RenderComposedParagraph requires op.AuthoredTokens.");
        string? baseAnchor = op.Op.LeftAnchor;
        var basePara = SourceElement(baseAnchor, baseIr);
        if (basePara == null)
        {
            // Defensive: a composed op should always resolve its base paragraph. Fall back to the merged
            // diff via the single-reviewer path so the op is not silently dropped.
            if (op.Op.TokenDiff != null)
                RenderModifiedParagraph(op.Op, op.Op.TokenDiff, state, sink);
            // else: base paragraph AND token diff both missing — nothing to emit; skip.
            return;
        }

        // Base left tokens + run model, built once (every group's Equal/Delete spans read these).
        var leftTokens = ParagraphTokens(baseAnchor, baseIr, state.Settings);
        var leftRuns = new SourceRunModel(basePara);

        // Per-reviewer right paragraph: tokens + run model, resolved from op.SourceRightAnchors (each
        // contributing reviewer's OWN right paragraph for this base block) and cached so a reviewer with
        // several disjoint word edits builds its model once.
        var rightAnchorByReviewer = new Dictionary<int, string>();
        if (op.SourceRightAnchors != null)
            foreach (var sra in op.SourceRightAnchors)
                rightAnchorByReviewer[sra.Reviewer] = sra.Anchor;
        var rightTokensCache = new Dictionary<int, IReadOnlyList<IrDiffToken>>();
        var rightRunsCache = new Dictionary<int, SourceRunModel>();

        // Clone the BASE pPr (deterministic; composed edits are text-only).
        var newPara = new XElement(W.p);
        var basePPr = basePara.Element(W.pPr);
        if (basePPr != null)
            newPara.Add(StripUnids(new XElement(basePPr)));

        // Save/restore the shared state's per-op fields so the renderer's outer loop is unaffected.
        var savedAuthor = state.AuthorOverride;
        var savedRightSource = state.RightSource;
        var savedRightSourceId = state.RightSourceId;

        int i = 0;
        var count = authored.Count;
        while (i < count)
        {
            // Coalesce the maximal run of consecutive authored ops sharing (Author, SourceReviewer).
            int reviewer = authored[i].SourceReviewer;
            string author = authored[i].Author;
            int groupStart = i;
            while (i < count &&
                   authored[i].SourceReviewer == reviewer &&
                   string.Equals(authored[i].Author, author, StringComparison.Ordinal))
                i++;

            var groupOps = new IrTokenOp[i - groupStart];
            for (int k = 0; k < groupOps.Length; k++)
                groupOps[k] = authored[groupStart + k].Op;
            var subDiff = new IrTokenDiff(IrNodeList.From(groupOps));

            if (reviewer < 0)
            {
                // Base-sourced group (Equal spans): both sides read the base paragraph; no author override.
                state.AuthorOverride = null;
                state.RightSource = baseIr;
                state.RightSourceId = -1;
                newPara.Add(BuildTokenOpContent(subDiff, leftTokens, leftTokens, leftRuns, leftRuns, state));
            }
            else
            {
                // Reviewer-sourced group: point the right side at THAT reviewer's right paragraph so Insert
                // spans (indexing the reviewer's right-token list) resolve to its runs; Delete spans still
                // read the base (left) model. Author override attributes every emitted revision.
                state.AuthorOverride = author;
                state.RightSource = reviewerIrs[reviewer];
                state.RightSourceId = reviewer;

                if (!rightTokensCache.TryGetValue(reviewer, out var rightTokens))
                {
                    string? ra = rightAnchorByReviewer.TryGetValue(reviewer, out var a) ? a : null;
                    rightTokens = ParagraphTokens(ra, reviewerIrs[reviewer], state.Settings);
                    rightTokensCache[reviewer] = rightTokens;
                    var rightPara = SourceElement(ra, reviewerIrs[reviewer]);
                    rightRunsCache[reviewer] = rightPara != null
                        ? new SourceRunModel(rightPara)
                        : new SourceRunModel(new XElement(W.p));
                }
                var rightRuns = rightRunsCache[reviewer];
                newPara.Add(BuildTokenOpContent(subDiff, leftTokens, rightTokens, leftRuns, rightRuns, state));
            }
        }

        state.AuthorOverride = savedAuthor;
        state.RightSource = savedRightSource;
        state.RightSourceId = savedRightSourceId;

        sink.Add(newPara);
    }

    // ----------------------------------------------------------------- fine modify path

    /// <summary>
    /// Render a Modified paragraph from its token diff: build the new paragraph's run-level content by walking
    /// the token-op spans. Equal/FormatChanged spans contribute the RIGHT paragraph's runs over that char
    /// span (unwrapped); Insert spans contribute the RIGHT runs wrapped <c>w:ins</c>; Delete spans contribute
    /// the LEFT runs wrapped <c>w:del</c> (<c>w:t</c>→<c>delText</c>). The paragraph mark: if the two
    /// paragraphs' marks were content-equal we leave it unmarked; Task 3 treats every Modify as a same-mark
    /// edit (paragraph splits/merges are block-level Insert/Delete ops, not Modify), so the mark is never
    /// revision-marked here.
    /// </summary>
    private static void RenderModifiedParagraph(
        IrEditOp op, IrTokenDiff tokenDiff, RenderState state, List<XElement> sink)
    {
        var leftPara = SourceElement(op.LeftAnchor, state.Left);
        var rightPara = SourceElement(op.RightAnchor, state.RightSource);
        if (leftPara == null || rightPara == null)
        {
            // Defensive: fall back to whole-block del+ins if a source element is unexpectedly missing.
            if (op.LeftAnchor != null) EmitWholeBlock(op.LeftAnchor, state.Left, state, sink, RevKind.Del, false);
            if (op.RightAnchor != null) EmitWholeBlock(op.RightAnchor, state.RightSource, state, sink, RevKind.Ins, true);
            return;
        }

        var leftRuns = new SourceRunModel(leftPara);
        var rightRuns = new SourceRunModel(rightPara);

        // Resolve token char spans: a token op's left span is [left[LeftStart].StartChar, left[LeftEnd-1].EndChar)
        // and likewise right. We resolve via the tokenizers so char coordinates match the diff's exactly.
        var leftTokens = ParagraphTokens(op.LeftAnchor, state.Left, state.Settings);
        var rightTokens = ParagraphTokens(op.RightAnchor, state.RightSource, state.Settings);

        // The new paragraph: clone the RIGHT paragraph's pPr (accepted-state paragraph properties) and rebuild
        // its run-level content from the spans.
        var newPara = new XElement(W.p);
        var rightPPr = rightPara.Element(W.pPr);
        if (rightPPr != null)
            newPara.Add(StripUnids(new XElement(rightPPr)));

        newPara.Add(BuildTokenOpContent(tokenDiff, leftTokens, rightTokens, leftRuns, rightRuns, state));
        sink.Add(newPara);
    }

    /// <summary>
    /// Build the run-level content for one token diff over explicit token lists / run models — shared by
    /// <see cref="RenderModifiedParagraph"/> (whole-paragraph diff) and the M2.6 split/merge segment
    /// rendering (slice diffs). Token spans index the GIVEN lists; char spans resolve through the tokens'
    /// own absolute StartChar/EndChar, so a SLICE of a paragraph's token list (which retains the source
    /// paragraph's char positions) composes with the full paragraph's <see cref="SourceRunModel"/> unchanged.
    /// </summary>
    private static List<XElement> BuildTokenOpContent(
        IrTokenDiff tokenDiff,
        IReadOnlyList<IrDiffToken> leftTokens, IReadOnlyList<IrDiffToken> rightTokens,
        SourceRunModel leftRuns, SourceRunModel rightRuns, RenderState state)
    {
        var content = new List<XElement>();
        foreach (var tokenOp in tokenDiff.Ops)
        {
            switch (tokenOp.Kind)
            {
                case IrTokenOpKind.Equal:
                case IrTokenOpKind.FormatChanged:
                {
                    // Right-side runs as-is. BUT a span that is "Equal" by MATCH KEY can still differ in RAW
                    // text — the tokenizer conflates NBSP↔space and case-folds keys, so e.g. a left space vs
                    // right NBSP at the same position is an Equal token op whose raw bytes differ. Emitting the
                    // unwrapped right run there would make reject-all keep the RIGHT byte (NBSP) instead of
                    // restoring the LEFT (space). So when the span's raw left/right text is NOT byte-identical,
                    // fall back to del(left)+ins(right) for that span — the accept/reject text invariant then
                    // holds byte-for-byte.
                    var (rs, re) = RightSpanChars(rightTokens, tokenOp);
                    var (ls, le) = LeftSpanChars(leftTokens, tokenOp);
                    var (rzs, rze) = ZeroWidthBoundaries(rightTokens, tokenOp.RightStart, tokenOp.RightEnd);
                    var (lzs, lze) = ZeroWidthBoundaries(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd);
                    string rawRight = RawSpanText(rightTokens, tokenOp.RightStart, tokenOp.RightEnd);
                    string rawLeft = RawSpanText(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd);
                    if (!string.Equals(rawLeft, rawRight, StringComparison.Ordinal))
                    {
                        foreach (var r in leftRuns.Slice(ls, le, lzs, lze))
                            content.Add(WrapRunLevel(r, RevKind.Del, state));
                        foreach (var r in rightRuns.Slice(rs, re, rzs, rze))
                            content.Add(WrapRunLevel(r, RevKind.Ins, state));   // registers media on its clone
                    }
                    else if (tokenOp.Kind == IrTokenOpKind.FormatChanged)
                    {
                        // Text-equal, FORMAT-differing: emit the RIGHT runs (accepted-state formatting) and stamp
                        // each with a w:rPrChange carrying the LEFT (old) modeled rPr. Accept drops the rPrChange
                        // (keeps right format); reject swaps the run's rPr to the rPrChange's inner rPr (restores
                        // the left format). The old rPr is rebuilt from the LEFT token's modeled IrRunFormat at
                        // the aligned position, so the modeled-only block format signature round-trips.
                        // rawLeft == rawRight here, so the left/right char spans carry identical text and the
                        // left char at offset k matches the right char at offset k. For each emitted right run
                        // covering right chars [cursor, cursor+len), clone the LEFT run's rPr at the aligned left
                        // char (ls + (cursor-rs)) as the old formatting, preserving modeled AND unmodeled left rPr.
                        int cursor = rs;
                        foreach (var r in rightRuns.Slice(rs, re, rzs, rze))
                        {
                            state.RegisterMediaReferences(r);
                            // Only a w:r carries run formatting — bookmarks/zero-width markers pass through
                            // untouched (stamping an rPr onto them is schema-invalid). A w:hyperlink wrapper holds
                            // its w:r(s) one level down, so stamp each contained run at its aligned left char and
                            // advance the cursor per-run (descending into the wrapper preserves char alignment).
                            foreach (var innerRun in RunsForFormatStamp(r))
                            {
                                int leftChar = ls + (cursor - rs);
                                var oldRPr = leftRuns.RPrAtChar(leftChar);
                                ApplyRPrChange(innerRun, oldRPr, state);
                                cursor += RunTextLength(innerRun);
                            }
                            content.Add(r);
                        }
                    }
                    else
                    {
                        foreach (var r in rightRuns.Slice(rs, re, rzs, rze))
                        {
                            state.RegisterMediaReferences(r);
                            content.Add(r);
                        }
                    }
                    break;
                }
                case IrTokenOpKind.Insert:
                {
                    var (s, e) = RightSpanChars(rightTokens, tokenOp);
                    var (zs, ze) = ZeroWidthBoundaries(rightTokens, tokenOp.RightStart, tokenOp.RightEnd);
                    foreach (var r in rightRuns.Slice(s, e, zs, ze))
                        content.Add(WrapRunLevel(r, RevKind.Ins, state));   // registers media on its clone
                    break;
                }
                case IrTokenOpKind.Delete:
                {
                    var (s, e) = LeftSpanChars(leftTokens, tokenOp);
                    var (zs, ze) = ZeroWidthBoundaries(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd);
                    foreach (var r in leftRuns.Slice(s, e, zs, ze))
                        content.Add(WrapRunLevel(r, RevKind.Del, state));
                    break;
                }
            }
        }

        return CoalesceAdjacentHyperlinks(content);
    }

    /// <summary>
    /// Merge consecutive <c>w:hyperlink</c> siblings that are PIECES OF ONE SOURCE LINK back into a single
    /// <c>w:hyperlink</c>. The token-op walk slices a hyperlink whose anchor is edited INTERNALLY into one wrapper
    /// per overlapping op (e.g. Equal "our " then del "website" then ins "homepage"); without re-joining them the
    /// rendered structure would be N adjacent links instead of the source's one, so accept/reject would match
    /// only at the text level, not the block ContentHash level.
    ///
    /// Adjacency is grouped by SOURCE-LINK IDENTITY (the transient <see cref="SourceLinkId"/> ordinal stamped in
    /// <see cref="SourceRunModel.Slice"/>), NOT by attribute equality. Two genuinely DISTINCT adjacent source
    /// links that happen to share a target carry DIFFERENT ordinals, so they never group — fixing the regression
    /// where attribute-equality grouping folded an unedited link plus the next link's edit into one link, so
    /// reject collapsed two source links into one (ContentHash divergence). All fragments of ONE edited link
    /// share the ordinal (the LEFT and RIGHT models number their Nth hyperlink identically), so an intra-anchor
    /// edit's Equal/del/ins pieces group together.
    ///
    /// The merge is still GATED to never join two links that may be DIFFERENT targets at the SAME ordinal. A
    /// wrapper is "revision-pure" when all its run-level children are <c>w:del</c>/<c>w:ins</c> (no plain
    /// <c>w:r</c>). The whole-anchor retarget case (WC019: a link's text AND href both change → pure
    /// <c>w:del</c>-link of the OLD target followed by a pure <c>w:ins</c>-link of the NEW target, the new id
    /// remapped post-assembly) is ordinal-0 del + ordinal-0 ins with NO plain piece — those MUST stay separate
    /// so the remap + empty-shell-drop restore the right/left link on each side. An intra-anchor TEXT edit always
    /// carries an Equal (plain-run) piece that anchors the group, so its del/ins pieces fold into a link that
    /// already holds a plain run — the gate lets that through. The transient ordinal marker is stripped from
    /// every wrapper before return so it never reaches output.
    /// </summary>
    private static List<XElement> CoalesceAdjacentHyperlinks(List<XElement> content)
    {
        var merged = new List<XElement>();
        int i = 0;
        while (i < content.Count)
        {
            var el = content[i];
            if (el.Name != W.hyperlink)
            {
                merged.Add(el);
                i++;
                continue;
            }

            // Gather the maximal run of adjacent hyperlinks that came from the SAME source link (same
            // SourceLinkId ordinal). They are pieces of the same token-op walk over one source-link char span;
            // the slicer emits one wrapper per overlapping op. Distinct adjacent links carry distinct ordinals and
            // so end the run, even when their attributes (target) coincide.
            int j = i + 1;
            while (j < content.Count && content[j].Name == W.hyperlink && SameSourceLink(el, content[j]))
                j++;

            int runLen = j - i;
            // Coalesce the run into ONE w:hyperlink IFF it carries at least one plain (Equal/unchanged) run — the
            // signal that some anchor text matched on both sides (so the link's target is the SAME on both sides).
            // A run with NO plain piece is a pure del→ins retarget (WC019: text AND href both change → del-link of
            // the OLD target, ins-link of the NEW target, the new id remapped post-assembly) — those MUST stay
            // separate so the remap + empty-shell-drop restore the right/left link on each side.
            bool anyPlainRun = false;
            for (int k = i; k < j; k++)
                if (!IsRevisionPureHyperlink(content[k]))
                {
                    anyPlainRun = true;
                    break;
                }

            if (runLen > 1 && anyPlainRun)
            {
                var combined = new XElement(el);                 // clone the first (carries the shell attributes)
                for (int k = i + 1; k < j; k++)
                    combined.Add(content[k].Elements());
                merged.Add(combined);
            }
            else
            {
                for (int k = i; k < j; k++)
                    merged.Add(content[k]);
            }
            i = j;
        }

        // Strip the transient source-link ordinal marker from every emitted wrapper so it never reaches output.
        // (The body-wide pt: sweep in Render would also catch it, but stripping here keeps the marker an internal
        // detail of the coalescer and protects callers that inspect the returned content before that sweep.)
        foreach (var el in merged)
            if (el.Name == W.hyperlink)
                el.Attribute(SourceLinkId)?.Remove();
        return merged;
    }

    /// <summary>True iff two emitted <c>w:hyperlink</c> wrappers came from the SAME source link — i.e. they carry
    /// the same transient <see cref="SourceLinkId"/> ordinal. A wrapper with no ordinal (defensive: a hyperlink
    /// that bypassed the ordinal-stamping slice path) matches only another with no ordinal AND the same target
    /// shell, so it never spuriously groups with a distinct link.</summary>
    private static bool SameSourceLink(XElement a, XElement b)
    {
        var ao = a.Attribute(SourceLinkId)?.Value;
        var bo = b.Attribute(SourceLinkId)?.Value;
        if (ao != null || bo != null)
            return ao == bo;
        return SameHyperlinkShell(a, b);
    }

    /// <summary>True iff a <c>w:hyperlink</c> carries no plain (Equal/unchanged) run — every run-level child is a
    /// revision wrapper (<c>w:del</c>/<c>w:ins</c>), or it is empty. (An empty link has no anchoring plain content
    /// either, so it counts as revision-pure for the coalescing gate.)</summary>
    private static bool IsRevisionPureHyperlink(XElement hyperlink) =>
        hyperlink.Elements().All(c => c.Name == W.del || c.Name == W.ins);

    /// <summary>True iff two <c>w:hyperlink</c> elements carry the same MEANINGFUL attribute set (name+value),
    /// i.e. they target the same link — so their run content can be merged. Ignores namespace declarations and
    /// Docxodus-internal <c>pt:</c> tracking attributes (notably the per-element <c>pt:Unid</c>, which is unique
    /// per source node and stripped before output), since those don't affect link identity.</summary>
    private static bool SameHyperlinkShell(XElement a, XElement b)
    {
        static List<XAttribute> Meaningful(XElement e) =>
            e.Attributes()
             .Where(at => !at.IsNamespaceDeclaration && at.Name.Namespace != PtOpenXml.pt && at.Name != PtOpenXml.Unid)
             .ToList();

        var aa = Meaningful(a);
        var bb = Meaningful(b);
        if (aa.Count != bb.Count)
            return false;
        foreach (var attr in aa)
        {
            var other = b.Attribute(attr.Name);
            if (other == null || other.Value != attr.Value)
                return false;
        }
        return true;
    }

    /// <summary>Whether the half-open token range <c>[start,end)</c> STARTS / ENDS with a zero-width token — a
    /// tab, break, note ref, image, opaque, or textbox placeholder, each contributing 0 chars (<c>StartChar ==
    /// EndChar</c>). A boundary zero-width token sits exactly at the op's start/end char, which two adjacent ops
    /// share; the caller passes these flags to <see cref="SourceRunModel.Slice"/> so the op that OWNS the token
    /// (it is the op's first / last token) claims it and the other does not — keeping a tail footnote reference
    /// from being dropped and a shared tab from being double-counted.</summary>
    private static (bool Start, bool End) ZeroWidthBoundaries(
        IReadOnlyList<IrDiffToken> tokens, int start, int end) =>
        end <= start
            ? (false, false)
            : (tokens[start].StartChar == tokens[start].EndChar,
               tokens[end - 1].StartChar == tokens[end - 1].EndChar);

    /// <summary>Concatenate the RAW token text over a half-open token-index span (empty span ⇒ "").</summary>
    private static string RawSpanText(IReadOnlyList<IrDiffToken> tokens, int start, int end)
    {
        if (start >= end)
            return string.Empty;
        var sb = new System.Text.StringBuilder();
        for (int i = start; i < end; i++)
            sb.Append(tokens[i].Text);
        return sb.ToString();
    }

    /// <summary>Right char span of a token op: empty (zero-width at the right anchor) for a Delete op.</summary>
    private static (int Start, int End) RightSpanChars(IReadOnlyList<IrDiffToken> tokens, IrTokenOp op)
    {
        if (op.RightStart >= op.RightEnd)
        {
            // Empty right span: position is the start-char of the right anchor token (or end-of-paragraph).
            int at = op.RightStart < tokens.Count ? tokens[op.RightStart].StartChar
                   : (tokens.Count > 0 ? tokens[^1].EndChar : 0);
            return (at, at);
        }
        return (tokens[op.RightStart].StartChar, tokens[op.RightEnd - 1].EndChar);
    }

    /// <summary>Left char span of a token op: empty (zero-width) for an Insert op.</summary>
    private static (int Start, int End) LeftSpanChars(IReadOnlyList<IrDiffToken> tokens, IrTokenOp op)
    {
        if (op.LeftStart >= op.LeftEnd)
        {
            int at = op.LeftStart < tokens.Count ? tokens[op.LeftStart].StartChar
                   : (tokens.Count > 0 ? tokens[^1].EndChar : 0);
            return (at, at);
        }
        return (tokens[op.LeftStart].StartChar, tokens[op.LeftEnd - 1].EndChar);
    }

    // ----------------------------------------------------------------- text → delText

    /// <summary>Convert every <c>w:t</c> descendant of a run-level element to <c>w:delText</c> in place,
    /// preserving its text and any <c>xml:space</c>. Required for deletions: accept drops the whole
    /// <c>w:del</c>; reject swaps to <c>w:ins</c> and <c>delText</c>→<c>t</c>.</summary>
    private static void ConvertTextToDelText(XElement runLevel)
    {
        foreach (var t in runLevel.DescendantsAndSelf(W.t).ToList())
            t.Name = W.delText;
    }

    // ----------------------------------------------------------------- rPrChange (format change)

    /// <summary>The number of <c>w:t</c> characters a rebuilt run carries (for advancing the char cursor).</summary>
    private static int RunTextLength(XElement run) =>
        run.Elements(W.t).Sum(t => t.Value.Length);

    /// <summary>The <c>w:r</c> elements that should receive a <c>w:rPrChange</c> stamp for a FormatChanged span,
    /// given an emitted run-level element: the element itself when it IS a run, otherwise its descendant runs (a
    /// <c>w:hyperlink</c> wrapper holds its run(s) one level down). Non-run, run-less elements (bookmarks,
    /// zero-width markers) yield nothing — stamping an rPr onto them is schema-invalid.</summary>
    private static System.Collections.Generic.IEnumerable<XElement> RunsForFormatStamp(XElement runLevel) =>
        runLevel.Name == W.r
            ? new[] { runLevel }
            : runLevel.Descendants(W.r);

    /// <summary>
    /// Stamp <paramref name="run"/> (a RIGHT-side run rebuilt over a FormatChanged span) with a
    /// <c>w:rPrChange</c> carrying <paramref name="oldRPr"/> (the LEFT/old run properties). The rPrChange is
    /// the LAST child of the run's <c>w:rPr</c> (schema order). Accept drops it (run keeps its right rPr);
    /// reject swaps the run's rPr to the rPrChange's inner rPr (restoring the left formatting). When the run
    /// has no <c>w:rPr</c> yet, an empty one is created so the right-side (accepted) format is "no rPr".
    /// </summary>
    private static void ApplyRPrChange(XElement run, XElement? oldRPr, RenderState state)
    {
        var rPr = run.Element(W.rPr);
        if (rPr == null)
        {
            rPr = new XElement(W.rPr);
            run.AddFirst(rPr);
        }
        // Idempotence: never stack rPrChange markers.
        rPr.Elements(W.rPrChange).Remove();
        // The inner rPr is the OLD properties; an absent/empty old rPr is encoded as an empty w:rPr.
        var inner = oldRPr != null ? StripUnids(new XElement(W.rPr, oldRPr.Attributes(), oldRPr.Elements())) : new XElement(W.rPr);
        rPr.Add(new XElement(W.rPrChange, state.RevisionAttributes(), inner));
    }

    // ----------------------------------------------------------------- helpers

    /// <summary>
    /// Recreate hyperlink/external relationships referenced by RIGHT-sourced clones into the LEFT main part.
    /// A <c>w:hyperlink/@r:id</c> (or any r:id resolving to a hyperlink/external relationship, never a part)
    /// must point at a relationship that exists in the output package, or accept/reject re-reads the target as
    /// null and the framed-target content hash diverges. We recreate each missing relationship with the SAME id
    /// when that id is free in the left part (the common case — ids rarely collide across the two documents).
    /// </summary>
    private static void ImportHyperlinkAndExternalRelationships(
        List<XElement> rightClones, MainDocumentPart leftMain, MainDocumentPart rightMain)
    {
        var leftHyper = leftMain.HyperlinkRelationships.ToDictionary(r => r.Id, StringComparer.Ordinal);
        var leftExternalIds = new HashSet<string>(leftMain.ExternalRelationships.Select(r => r.Id), StringComparer.Ordinal);
        var rightHyper = rightMain.HyperlinkRelationships.ToDictionary(r => r.Id, StringComparer.Ordinal);
        var rightExternal = rightMain.ExternalRelationships.ToDictionary(r => r.Id, StringComparer.Ordinal);

        // Collect referenced r:ids across all right clones.
        var referenced = new HashSet<string>(StringComparer.Ordinal);
        foreach (var clone in rightClones)
            foreach (var attr in clone.DescendantsAndSelf().Attributes().Where(a => a.Name.Namespace == R.r))
            {
                var id = (string?)attr;
                if (!string.IsNullOrEmpty(id))
                    referenced.Add(id);
            }

        foreach (var id in referenced)
        {
            if (rightHyper.TryGetValue(id, out var hr))
            {
                if (!leftHyper.ContainsKey(id))
                {
                    // The id is FREE in the left part — recreate the relationship under the SAME id so the
                    // cloned w:hyperlink/@r:id keeps resolving (the common, no-collision case).
                    try { leftMain.AddHyperlinkRelationship(hr.Uri, hr.IsExternal, id); }
                    catch (Exception ex) when (ex is ArgumentException or InvalidOperationException) { }
                }
                else if (hr.Uri is { } hrUri &&
                         !string.Equals(leftHyper[id].Uri?.ToString(), hrUri.ToString(), StringComparison.Ordinal))
                {
                    // COLLISION (WC019): the id already names a DIFFERENT left hyperlink (Before → ericwhite.com,
                    // After → ericwhite2.com both as rId4). Reusing it would leave the cloned right hyperlink
                    // pointing at the LEFT target. True rId REMAP: mint a fresh id, recreate the right target
                    // under it, and rewrite the cloned @r:id so accept reads the RIGHT target. (Same id + same
                    // target is a no-op — the existing left relationship already resolves correctly.)
                    string fresh = FreshRelationshipId(leftMain);
                    leftMain.AddHyperlinkRelationship(hrUri, hr.IsExternal, fresh);
                    RewriteReferenceId(rightClones, id, fresh);
                }
            }
            else if (rightExternal.TryGetValue(id, out var er))
            {
                if (!leftExternalIds.Contains(id))
                {
                    try { leftMain.AddExternalRelationship(er.RelationshipType, er.Uri, id); }
                    catch (Exception ex) when (ex is ArgumentException or InvalidOperationException) { }
                }
                else
                {
                    var leftEr = leftMain.ExternalRelationships.FirstOrDefault(r => r.Id == id);
                    if (er.Uri is { } erUri &&
                        (leftEr is null || !string.Equals(leftEr.Uri?.ToString(), erUri.ToString(), StringComparison.Ordinal)))
                    {
                        // Same collision class for an external (non-hyperlink) relationship: remap to a fresh id.
                        string fresh = FreshRelationshipId(leftMain);
                        leftMain.AddExternalRelationship(er.RelationshipType, erUri, fresh);
                        RewriteReferenceId(rightClones, id, fresh);
                    }
                }
            }
        }
    }

    /// <summary>A relationship id not currently in use by any of the left main part's relationships (parts,
    /// hyperlinks, external links, and data-part references alike). Deterministic: the first free
    /// <c>rIdRemap{n}</c> (n ascending from 1). The dedicated <c>rIdRemap</c> prefix avoids colliding with the
    /// document's own <c>rId{n}</c> numbering — the very collision this remap exists to resolve.</summary>
    private static string FreshRelationshipId(MainDocumentPart leftMain)
    {
        var used = new HashSet<string>(StringComparer.Ordinal);
        foreach (var rel in leftMain.Parts) used.Add(rel.RelationshipId);
        foreach (var rel in leftMain.HyperlinkRelationships) used.Add(rel.Id);
        foreach (var rel in leftMain.ExternalRelationships) used.Add(rel.Id);
        foreach (var rel in leftMain.DataPartReferenceRelationships) used.Add(rel.Id);
        int n = 1;
        string candidate;
        do { candidate = "rIdRemap" + n++; } while (used.Contains(candidate));
        return candidate;
    }

    /// <summary>Rewrite every <c>@r:id</c> (any relationship-namespace attribute) on the right clones that
    /// currently reads <paramref name="oldId"/> to <paramref name="newId"/>, so the remapped relationship
    /// resolves. Scoped to the relationship namespace so only true r:id references are touched.</summary>
    private static void RewriteReferenceId(List<XElement> rightClones, string oldId, string newId)
    {
        foreach (var clone in rightClones)
            foreach (var attr in clone.DescendantsAndSelf().Attributes().Where(a => a.Name.Namespace == R.r))
                if (string.Equals((string?)attr, oldId, StringComparison.Ordinal))
                    attr.Value = newId;
    }

    /// <summary>True iff this op concerns a standalone section-break block (a `sec:` anchor on either side, or a
    /// resolved <see cref="IrSectionBreak"/>) — the trailing last-section metadata we never emit into the body.</summary>
    private static bool IsSectionBreakOp(IrEditOp op, RenderState state)
    {
        if ((op.RightAnchor?.StartsWith("sec:", StringComparison.Ordinal) ?? false) ||
            (op.LeftAnchor?.StartsWith("sec:", StringComparison.Ordinal) ?? false))
            return true;
        return ResolveBlock(op.RightAnchor, state.RightSource) is IrSectionBreak ||
               ResolveBlock(op.LeftAnchor, state.Left) is IrSectionBreak;
    }

    private static IrBlock? ResolveBlock(string? anchor, IrDocument doc) =>
        anchor != null && doc.AnchorIndex.TryGetValue(anchor, out var b) ? b : null;

    /// <summary>The source <c>w:p</c>/<c>w:tbl</c>/… XElement a block anchor resolves to, or null. Requires the
    /// block was read with <c>RetainSources=true</c> (the renderer's internal read does this).</summary>
    private static XElement? SourceElement(string? anchor, IrDocument doc) =>
        ResolveBlock(anchor, doc)?.Source.Element;

    private static IReadOnlyList<IrDiffToken> ParagraphTokens(string? anchor, IrDocument doc, IrDiffSettings settings)
    {
        if (anchor != null && doc.AnchorIndex.TryGetValue(anchor, out var block) && block is IrParagraph p)
            return IrDiffTokenizer.Tokenize(p, settings);
        return Array.Empty<IrDiffToken>();
    }

    /// <summary>Strip the reader-assigned <c>pt:Unid</c> bookkeeping attributes/elements from a cloned element so
    /// the output carries no engine-internal markup.</summary>
    private static XElement StripUnids(XElement el)
    {
        foreach (var attr in el.DescendantsAndSelf().Attributes()
                     .Where(a => a.Name.Namespace == PtOpenXml.pt || a.Name == PtOpenXml.Unid).ToList())
            attr.Remove();
        return el;
    }

    private enum RevKind { Ins, Del, MoveFrom, MoveTo }

    /// <summary>The OOXML revision-wrapper element name for a <see cref="RevKind"/>.</summary>
    private static XName RevElementName(RevKind kind) => kind switch
    {
        RevKind.Ins => W.ins,
        RevKind.Del => W.del,
        RevKind.MoveFrom => W.moveFrom,
        RevKind.MoveTo => W.moveTo,
        _ => W.ins,
    };

    /// <summary>True for the "delete-grade" kinds whose <c>w:t</c> must become <c>w:delText</c> (the moved-FROM
    /// content is removed on accept, like a deletion).</summary>
    private static bool IsDeleteGrade(RevKind kind) => kind is RevKind.Del or RevKind.MoveFrom;

    // ----------------------------------------------------------------- per-call state

    /// <summary>
    /// Mutable per-<see cref="Render"/> state: the two IR snapshots (with provenance), settings, the SINGLE
    /// ascending revision-id counter (no static state), and the live RIGHT-sourced clone roots whose media must
    /// be imported into the left package. One instance per call ⇒ concurrent renders never share a counter.
    /// </summary>
    internal sealed class RenderState
    {
        private int _nextId = 1;

        public RenderState(IrDocument left, IrDocument right, IrDiffSettings settings)
        {
            Left = left;
            Right = right;
            RightSource = right;   // two-way: the right source IS the right doc — never reassigned, so behavior is unchanged.
            Settings = settings;
        }

        public IrDocument Left { get; }
        public IrDocument Right { get; }
        public IrDiffSettings Settings { get; }

        /// <summary>The document the CURRENTLY-emitting op draws inserted/modified ("right-side") block elements
        /// and token text from. In a two-way render this is always <see cref="Right"/> (set once in the ctor and
        /// never reassigned), so behavior is byte-identical to before this field existed. The composite renderer
        /// switches it per op to the contributing reviewer's IR (or <see cref="Left"/>/base for a base-sourced
        /// equal/delete) so the existing emit helpers can be reused per-reviewer.</summary>
        public IrDocument RightSource { get; set; }

        /// <summary>When non-null, overrides Settings.AuthorForRevisions for emitted revision attributes
        /// (composite multi-author rendering). Null for normal two-way render → behavior unchanged.</summary>
        public string? AuthorOverride { get; set; }

        /// <summary>The bucket key of the CURRENTLY-active right source package, used to attribute media-bearing
        /// clones to the package they must be imported FROM. Two-way uses the single key 0 (the right package);
        /// the composite renderer sets it to the contributing reviewer's index per op so <see cref="Render"/>'s
        /// media-import pass (composite path) can copy each clone's parts from the correct reviewer package.</summary>
        public int RightSourceId { get; set; }

        /// <summary>RIGHT-sourced clone roots that may carry image relationship references the base package cannot
        /// resolve, BUCKETED by <see cref="RightSourceId"/> (the source package they were cloned from). After they
        /// are placed in the new body (still the same XElement instances),
        /// <see cref="WmlComparer.MoveRelatedPartsToDestination"/> walks each and remaps ids in place. Only roots
        /// actually containing an r-namespace attribute are recorded, so the common text-only case adds nothing.
        /// In a two-way render every clone lands in bucket 0 (the right package).</summary>
        public Dictionary<int, List<XElement>> RightSourcedClonesBySource { get; } = new();

        /// <summary>The two-way render's single clone bucket (bucket 0 = the right package). Preserves the original
        /// flat-list API for the two-way <see cref="Render"/> media-import pass; equivalent to the bucket-0 list.
        /// Returns a shared immutable empty sequence when bucket 0 is absent; callers only read (never mutate) the
        /// returned value, so the shared-immutable pattern is safe and allocation-free.</summary>
        public IReadOnlyList<XElement> RightSourcedClones =>
            RightSourcedClonesBySource.TryGetValue(0, out var list) ? list : Array.Empty<XElement>();

        /// <summary>Fresh (author, id, date) attribute triple for one revision element; id ascends from 1.</summary>
        public object[] RevisionAttributes() => new object[]
        {
            new XAttribute(W.author, AuthorOverride ?? Settings.AuthorForRevisions),
            new XAttribute(W.id, _nextId++),
            new XAttribute(W.date, Settings.DateTimeForRevisions),
        };

        /// <summary>A fresh revision id (for move-range markers, which carry only an id, not the full triple).</summary>
        public int NextId() => _nextId++;

        private readonly Dictionary<int, string> _moveNames = new();
        private int _nextMoveName = 1;

        /// <summary>The deterministic <c>w:name</c> ("move1", "move2", …) shared by a move group's FROM and TO
        /// halves, keyed by <see cref="IrEditOp.MoveGroupId"/>. Allocated in first-seen order per render, so the
        /// source and destination ops (which carry the same group id) resolve to the SAME name regardless of
        /// which renders first. Mirrors WmlComparer's "move{n}" convention.</summary>
        public string MoveName(int moveGroupId)
        {
            if (!_moveNames.TryGetValue(moveGroupId, out var name))
            {
                name = "move" + _nextMoveName++;
                _moveNames[moveGroupId] = name;
            }
            return name;
        }

        /// <summary>Record a RIGHT-sourced clone for media import iff it references any relationship id (an
        /// image embed/link), into the bucket for the currently-active <see cref="RightSourceId"/>. The recorded
        /// element is the live tree node; importing happens post-assembly. Two-way always records into bucket 0.</summary>
        public void RegisterMediaReferences(XElement clone)
        {
            if (clone.DescendantsAndSelf().Attributes().Any(a => a.Name.Namespace == R.r))
            {
                if (!RightSourcedClonesBySource.TryGetValue(RightSourceId, out var list))
                    RightSourcedClonesBySource[RightSourceId] = list = new List<XElement>();
                list.Add(clone);
            }
        }
    }

    /// <summary>
    /// A run-level slicer over a source paragraph: walks the paragraph's run-level children, tracks the
    /// half-open char offset of each <c>w:t</c>'s text EXACTLY as <see cref="IrDiffTokenizer"/> does (only
    /// <c>w:t</c> text — including a field's cached result — advances the counter; <c>w:tab</c>/<c>w:br</c>/
    /// note refs/drawings/etc. are zero-width), and can produce the run-level XElements covering a half-open
    /// char span, splitting a run whose text straddles a boundary (cloning it and trimming the <c>w:t</c>) so
    /// every fragment keeps the run's full <c>w:rPr</c> — modeled AND unmodeled.
    /// </summary>
    private sealed class SourceRunModel
    {
        // Each segment is a contiguous piece of run-level content with a [Start,End) char span. A text segment
        // is one w:t inside one run (so it is splittable); a zero-width segment is a non-text run child or a
        // whole run carrying no text.
        private readonly List<Segment> _segments = new();

        /// <summary>Source zero-width child elements already emitted by a previous <see cref="Slice"/> call on
        /// THIS model (one model per paragraph side), so a boundary-shared zero-width inline is never emitted
        /// twice across adjacent token ops. Keyed by reference identity.</summary>
        private readonly HashSet<XElement> _claimedZeroWidth = new();

        /// <summary>0-based ordinal of each top-level source <c>w:hyperlink</c> element within this paragraph, in
        /// document order, keyed by the source element's reference identity. The LEFT and RIGHT models walk their
        /// paragraphs in the same order, so the Nth hyperlink on each side gets the SAME ordinal — a STABLE
        /// per-source-link id that <see cref="CoalesceAdjacentHyperlinks"/> uses to rejoin ONLY the fragments of
        /// ONE source link (an intra-anchor edit emits its Equal/del/ins pieces under one ordinal), and to keep
        /// genuinely DISTINCT adjacent links (different ordinals) separate even when they share a target.</summary>
        private readonly Dictionary<XElement, int> _hyperlinkOrdinal = new(ReferenceEqualityComparer.Instance);
        private int _nextHyperlinkOrdinal;

        public SourceRunModel(XElement para)
        {
            int charOffset = 0;
            foreach (var child in para.Elements().Where(e => e.Name != W.pPr))
                WalkRunLevel(child, ref charOffset, ContainerChain.Empty);
        }

        private void WalkRunLevel(XElement runLevel, ref int charOffset, ContainerChain chain)
        {
            if (runLevel.Name == W.r)
            {
                WalkRun(runLevel, ref charOffset, chain);
            }
            else if (runLevel.Name == W.hyperlink)
            {
                // A w:hyperlink wrapping runs. We RECURSE into its run-level children rather than treating it as
                // one atomic blob, recording the hyperlink in the owning chain so its WRAPPER is reconstructed in
                // Slice exactly ONCE per contiguous run group it contributes — even when several token ops overlap
                // its char span (an intra-anchor edit, e.g. changing one word of a multi-run anchor). Before this,
                // the whole hyperlink was re-emitted per overlapping op, doubling/tripling the anchor on the
                // accept/reject paths. A wrapper shell (the element with its attributes but WITHOUT inner content)
                // rides on each leaf segment so the rebuilt runs are re-wrapped. The char span advances exactly as
                // before (sum of descendant w:t lengths), so token char coordinates are unchanged. (Other
                // containers — sdt/smartTag/ins/del — stay atomic but are now claim-tracked in Slice so they too
                // emit once across overlapping ops; only the hyperlink needs intra-anchor splitting to round-trip.)
                // Assign this hyperlink its document-order ordinal (nested hyperlinks are schema-invalid, so only
                // top-level hyperlinks are numbered). Slice stamps the ordinal onto each emitted wrapper clone so
                // the coalescer can rejoin only fragments of the SAME source link.
                _hyperlinkOrdinal[runLevel] = _nextHyperlinkOrdinal++;
                var childChain = chain.Append(runLevel);
                bool anyChild = false;
                foreach (var child in runLevel.Elements())
                {
                    anyChild = true;
                    WalkRunLevel(child, ref charOffset, childChain);
                }
                // An empty hyperlink (no run-level children) still needs its wrapper preserved: emit a zero-width
                // segment carrying the shell chain so Slice re-wraps it once.
                if (!anyChild)
                    _segments.Add(new Segment(runLevel, charOffset, charOffset, SegmentKind.ZeroWidth) { Chain = childChain });
            }
            else if (runLevel.Name == W.ins || runLevel.Name == W.del ||
                     runLevel.Name == W.sdt || runLevel.Name == W.smartTag)
            {
                // Non-hyperlink container (sdt/smartTag/accepted ins-del wrapper): one ATOMIC segment spanning its
                // full inner text, emitted whole. Its char span is the sum of its descendant w:t lengths
                // (mirroring the tokenizer's transparent recursion). Claim-tracked in Slice so multiple
                // overlapping ops emit it once.
                int start = charOffset;
                foreach (var t in runLevel.Descendants(W.t))
                    charOffset += t.Value.Length;
                _segments.Add(new Segment(runLevel, start, charOffset, SegmentKind.Container) { Chain = chain });
            }
            else
            {
                // A non-run, non-container run-level element (bookmarkStart/End, proofErr, commentRangeStart…):
                // zero-width, atomic, kept whole.
                _segments.Add(new Segment(runLevel, charOffset, charOffset, SegmentKind.ZeroWidth) { Chain = chain });
            }
        }

        private void WalkRun(XElement run, ref int charOffset, ContainerChain chain)
        {
            // A run can contain multiple w:t / w:tab / w:br / drawing children. We emit one segment per child
            // so a span boundary inside the run splits at child granularity, and a w:t segment can split inside.
            bool any = false;
            foreach (var child in run.Elements().Where(e => e.Name != W.rPr))
            {
                any = true;
                if (child.Name == W.t)
                {
                    string text = child.Value;
                    int start = charOffset;
                    charOffset += text.Length;
                    _segments.Add(new Segment(run, start, charOffset, SegmentKind.RunText) { TextChild = child, Chain = chain });
                }
                else if (child.Name == W.fldSimple || IsContainer(child.Name))
                {
                    // A simple field's cached result advances the offset by its text (tokenizer recurses too).
                    int start = charOffset;
                    foreach (var t in child.Descendants(W.t))
                        charOffset += t.Value.Length;
                    _segments.Add(new Segment(run, start, charOffset, SegmentKind.RunOther) { OtherChild = child, Chain = chain });
                }
                else
                {
                    // tab/break/drawing/noteref/sym/… — zero-width run child.
                    _segments.Add(new Segment(run, charOffset, charOffset, SegmentKind.RunOther) { OtherChild = child, Chain = chain });
                }
            }
            if (!any)
                _segments.Add(new Segment(run, charOffset, charOffset, SegmentKind.RunOther) { Chain = chain });
        }

        private static bool IsContainer(XName n) =>
            n == W.hyperlink || n == W.ins || n == W.del || n == W.sdt || n == W.smartTag;

        /// <summary>Produce run-level XElements covering the half-open char span [start,end). Run children that
        /// fall (partly) inside the span are grouped back into per-source-run <c>w:r</c> clones carrying the
        /// original <c>w:rPr</c>; a straddling <c>w:t</c> is split.
        /// <para><b>Boundary zero-width ownership.</b> A zero-width inline (footnote/endnote reference, drawing,
        /// tab, break, …) occupies a single char position that two adjacent token ops SHARE (one ends there, the
        /// next starts there). It belongs to whichever op's token range holds it as its FIRST or LAST token — the
        /// diff already decided this. The caller passes <paramref name="includeStartZeroWidth"/> (its first token
        /// is zero-width) / <paramref name="includeEndZeroWidth"/> (its last token is zero-width); a STRICTLY
        /// interior zero-width is always taken. Without this, a half-open char span both DROPS a trailing
        /// zero-width (it sits at <c>end</c>, which no op's interior and no later op covers — the footnote-ref
        /// corruption) and DOUBLE-COUNTS a boundary one (an equal tab claimed both as equal by the op that owns it
        /// AND as deleted by the next op's start). Defaults <c>(true,false)</c> reproduce the original
        /// always-start-inclusive / end-exclusive rule for callers that don't pass token boundaries.</para></summary>
        public List<XElement> Slice(int start, int end, bool includeStartZeroWidth = true, bool includeEndZeroWidth = false)
        {
            var result = new List<XElement>();

            // Two-level grouping. INNER: consecutive RunText/RunOther segments sharing the SAME source run are
            // rebuilt into one w:r (so a split w:t and its siblings keep one run + its rPr). OUTER: consecutive
            // run-level pieces sharing the SAME owning container chain (e.g. the same w:hyperlink, by reference
            // identity) are collected and emitted under a SINGLE clone of that chain's wrapper shells. So a
            // hyperlink's wrapper is reconstructed EXACTLY ONCE per contiguous run group it contributes to this
            // slice — even when its anchor spans several source runs, and even when several token ops overlap its
            // char span (an intra-anchor edit). Before this, the whole hyperlink was re-emitted per overlapping
            // op, doubling/tripling the anchor on accept/reject.
            ContainerChain groupChain = ContainerChain.Empty;
            var groupChildren = new List<XElement>();          // run-level children to wrap in groupChain
            XElement? currentRun = null;
            XElement? rebuilt = null;

            void FlushRun()
            {
                if (rebuilt != null && rebuilt.Elements().Any(e => e.Name != W.rPr))
                    groupChildren.Add(rebuilt);
                rebuilt = null;
                currentRun = null;
            }

            void FlushGroup()
            {
                FlushRun();
                if (groupChildren.Count > 0)
                {
                    // Wrap the collected children in clones of each container shell (outermost first), so e.g.
                    // <w:hyperlink …><w:r>our </w:r><w:r>website</w:r></w:hyperlink> with the original attributes.
                    object content = groupChildren.ToArray();
                    for (int i = groupChain.Count - 1; i >= 0; i--)
                    {
                        var shell = groupChain[i];
                        var wrapper = new XElement(shell.Name, shell.Attributes(), content);
                        // Tag a hyperlink wrapper with its source-link ordinal (a TRANSIENT pt: marker the
                        // coalescer reads and then strips) so adjacent emitted fragments are rejoined ONLY when
                        // they came from the SAME source w:hyperlink — never two distinct links sharing a target.
                        if (shell.Name == W.hyperlink && _hyperlinkOrdinal.TryGetValue(shell, out int ord))
                            wrapper.SetAttributeValue(SourceLinkId, ord);
                        content = wrapper;
                    }
                    if (content is XElement single)
                        result.Add(single);
                    else
                        foreach (var c in (XElement[])content)
                            result.Add(c);
                }
                groupChildren = new List<XElement>();
                groupChain = ContainerChain.Empty;
            }

            void StartGroupIfNeeded(ContainerChain chain)
            {
                // A change of owning chain ends the current wrapper group (so a run leaving/entering a hyperlink
                // starts a fresh wrapper). The top-level (empty) chain groups plain runs together too.
                if (groupChildren.Count > 0 || currentRun != null)
                {
                    if (!groupChain.SameAs(chain))
                        FlushGroup();
                }
                groupChain = chain;
            }

            foreach (var seg in _segments)
            {
                bool overlaps;
                if (start == end)
                    overlaps = seg.Start == start && seg.IsZeroWidth;     // empty span: only zero-width at the point
                else if (seg.IsZeroWidth)
                    overlaps = (seg.Start > start && seg.Start < end)     // strictly interior: always taken
                            || (seg.Start == start && includeStartZeroWidth)  // leading: only if this op owns it
                            || (seg.Start == end && includeEndZeroWidth);     // trailing: only if this op owns it
                else
                    overlaps = seg.Start < end && seg.End > start;        // text overlap

                // A zero-width inline (note ref, drawing, tab, …) sits at ONE char position two adjacent
                // token-ops can SHARE (prev op's end char == this op's start char), so a char-span slice would
                // emit it twice. De-duplicate by the specific source CHILD element identity across the
                // paragraph: a given zero-width source node is sliced into exactly one output op (the first to
                // claim it). A standalone ZeroWidth segment keys on its own Element; a RunOther zero-width keys
                // on its OtherChild (so two distinct zero-width children of one run are not conflated).
                if (overlaps && seg.IsZeroWidth && seg.Kind != SegmentKind.Container)
                {
                    var key = seg.OtherChild ?? seg.Element;
                    if (key != null && !_claimedZeroWidth.Add(key))
                        overlaps = false;
                }

                // An ATOMIC Container segment (sdt/smartTag/ins/del) spans a char range several token ops can
                // overlap (an intra-container edit). Emitting it per op would double it, so claim it exactly once
                // (the first op to overlap it wins); later overlapping ops skip it. Keyed by element identity.
                if (overlaps && seg.Kind == SegmentKind.Container)
                {
                    if (!_claimedZeroWidth.Add(seg.Element))
                        overlaps = false;
                }

                if (!overlaps)
                {
                    if (seg.Kind == SegmentKind.Container || seg.Kind == SegmentKind.ZeroWidth)
                        FlushRun();
                    continue;
                }

                switch (seg.Kind)
                {
                    case SegmentKind.ZeroWidth:
                        // A zero-width marker (incl. an empty hyperlink shell) is its own run-level piece; it joins
                        // its owning chain group so a bookmark inside a hyperlink stays inside the wrapper.
                        StartGroupIfNeeded(seg.Chain);
                        FlushRun();
                        groupChildren.Add(new XElement(seg.Element));
                        break;

                    case SegmentKind.Container:
                        // Atomic non-hyperlink container: emitted whole under its own (parent) chain.
                        StartGroupIfNeeded(seg.Chain);
                        FlushRun();
                        groupChildren.Add(new XElement(seg.Element));
                        break;

                    case SegmentKind.RunText:
                    case SegmentKind.RunOther:
                    {
                        StartGroupIfNeeded(seg.Chain);
                        if (!ReferenceEquals(currentRun, seg.Element))
                        {
                            FlushRun();
                            currentRun = seg.Element;
                            rebuilt = new XElement(W.r);
                            var rPr = seg.Element.Element(W.rPr);
                            if (rPr != null)
                                rebuilt.Add(new XElement(rPr));
                        }
                        if (seg.Kind == SegmentKind.RunText && seg.TextChild != null)
                        {
                            // Possibly-split text: take the overlap [max(start,seg.Start), min(end,seg.End)).
                            int s = Math.Max(start, seg.Start);
                            int e = Math.Min(end, seg.End);
                            string full = seg.TextChild.Value;
                            string piece = full.Substring(s - seg.Start, e - s);
                            var t = new XElement(W.t, piece);
                            if (PreserveSpace(piece))
                                t.Add(new XAttribute(XNamespace.Xml + "space", "preserve"));
                            rebuilt!.Add(t);
                        }
                        else if (seg.OtherChild != null)
                        {
                            rebuilt!.Add(new XElement(seg.OtherChild));
                        }
                        break;
                    }
                }
            }
            FlushGroup();
            return result;
        }

        /// <summary>The <c>w:rPr</c> of the source run whose text covers char position <paramref name="at"/>,
        /// cloned (or null if that run has none / no segment covers the position). Used to recover the LEFT/old
        /// run properties for a FormatChanged span's <c>w:rPrChange</c>.</summary>
        public XElement? RPrAtChar(int at)
        {
            // Prefer a RunText segment that strictly contains [at, at+1); fall back to a segment starting at `at`.
            Segment? hit = null;
            foreach (var seg in _segments)
            {
                if (seg.Kind == SegmentKind.RunText && seg.Start <= at && at < seg.End)
                {
                    hit = seg;
                    break;
                }
            }
            hit ??= _segments.FirstOrDefault(s => s.Start <= at && (at < s.End || (s.IsZeroWidth && s.Start == at)));
            var rPr = hit?.Element.Element(W.rPr);
            return rPr != null ? new XElement(rPr) : null;
        }

        private static bool PreserveSpace(string s) =>
            s.Length > 0 && (char.IsWhiteSpace(s[0]) || char.IsWhiteSpace(s[^1]));

        private enum SegmentKind { RunText, RunOther, ZeroWidth, Container }

        /// <summary>The (possibly empty) chain of run-level container wrappers — outermost first — owning a
        /// segment, e.g. the <c>w:hyperlink</c> a run sits inside. Reference-identity based: two segments share a
        /// chain iff they came from the SAME wrapper element(s), so Slice re-wraps runs from one hyperlink
        /// together and starts a fresh wrapper for a different (or no) hyperlink. Immutable; <see cref="Empty"/>
        /// is the no-wrapper chain. Cheap shells (only the chain depth matters; clones are minted in Slice).</summary>
        private sealed class ContainerChain
        {
            public static readonly ContainerChain Empty = new(System.Array.Empty<XElement>());

            private readonly XElement[] _wrappers;
            private ContainerChain(XElement[] wrappers) => _wrappers = wrappers;

            public int Count => _wrappers.Length;
            public XElement this[int i] => _wrappers[i];

            public ContainerChain Append(XElement wrapper)
            {
                var next = new XElement[_wrappers.Length + 1];
                System.Array.Copy(_wrappers, next, _wrappers.Length);
                next[^1] = wrapper;
                return new ContainerChain(next);
            }

            /// <summary>True iff the two chains are the same length and reference the same wrapper elements in
            /// order (reference identity) — so runs nested in the identical hyperlink group together.</summary>
            public bool SameAs(ContainerChain other)
            {
                if (ReferenceEquals(this, other))
                    return true;
                if (_wrappers.Length != other._wrappers.Length)
                    return false;
                for (int i = 0; i < _wrappers.Length; i++)
                    if (!ReferenceEquals(_wrappers[i], other._wrappers[i]))
                        return false;
                return true;
            }
        }

        private sealed class Segment
        {
            public Segment(XElement element, int start, int end, SegmentKind kind)
            {
                Element = element;
                Start = start;
                End = end;
                Kind = kind;
            }

            public XElement Element { get; }
            public int Start { get; }
            public int End { get; }
            public SegmentKind Kind { get; }
            public XElement? TextChild { get; init; }
            public XElement? OtherChild { get; init; }
            public ContainerChain Chain { get; init; } = ContainerChain.Empty;
            public bool IsZeroWidth => Start == End;
        }
    }
}
