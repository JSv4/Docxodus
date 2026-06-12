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
/// <item><see cref="IrEditOpKind.FormatOnlyBlock"/> → the right block verbatim (Task 3 has no <c>w:rPrChange</c>;
/// the block is content-equal so accept/reject both yield correct TEXT — see the FormatChanged gap below).</item>
/// <item>A TABLE ModifyBlock, a non-paragraph Modified pair, moves, and notes → a conservative whole-block
/// <c>w:del</c> of the LEFT block immediately followed by a <c>w:ins</c> of the RIGHT block. Accept keeps the
/// right (correct), reject keeps the left (correct); the text-level invariant holds. Task 4 replaces these
/// with native table/move/note markup.</item>
/// </list></para>
///
/// <para><b>FormatChanged-span gap (precise).</b> A <see cref="IrTokenOpKind.FormatChanged"/> span is
/// TEXT-equal on both sides but FORMAT-differing. Task 3 renders it as the RIGHT-side runs with NO
/// <c>w:rPrChange</c>. Consequence: reject-all then restores the right-side FORMATTING on those runs (the
/// LEFT formatting is lost) while restoring the correct TEXT. Because THE INVARIANT compares per-block
/// <c>ContentHash</c> (which is text/structure, NOT modeled run format), it still holds. Task 4 closes this
/// by emitting <c>w:rPrChange</c> carrying the old rPr so reject restores the left formatting too. Likewise a
/// FormatOnly block reject yields right formatting; same Task-4 gap, same invariant safety.</para>
///
/// <para><b>Note scopes (Task 4).</b> <see cref="IrEditScript.NoteOps"/> are NOT rendered into footnote/
/// endnote part markup yet. The body still round-trips; note-scope markup + id uniqueness across scopes is
/// Task 4.</para>
/// </remarks>
internal static class IrMarkupRenderer
{
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
                if (rightMain != null && state.RightSourcedClones.Count > 0)
                {
                    // (1) Import hyperlink/external relationships (e.g. w:hyperlink/@r:id targets) the right
                    // clones reference but the left package lacks — these are NOT parts, so the part-copy path
                    // below skips them; recreate them with the SAME id where free so the cloned r:id resolves.
                    ImportHyperlinkAndExternalRelationships(state.RightSourcedClones, main, rightMain);

                    // (2) Import media PARTS (image embeds, diagram data) and remap their r:ids in place, using
                    // the stream documents' own packages directly (the wrapper's package is the authoritative
                    // writable one — not the reflection-based OpenXmlPackage.GetPackage()).
                    var leftPkgPart = streamDoc.GetPackage().GetPart(main.Uri);
                    var rightPkgPart = rightStream.GetPackage().GetPart(rightMain.Uri);
                    foreach (var clone in state.RightSourcedClones)
                        WmlComparer.MoveRelatedPartsToDestination(rightPkgPart, leftPkgPart, clone);
                }

                // Strip ALL engine-internal pt:Unid bookkeeping attributes from the assembled body (cloned runs
                // inside ins/del wrappers carry them too; a single sweep here catches every nested occurrence).
                foreach (var attr in bodyEl.DescendantsAndSelf().Attributes()
                             .Where(a => a.Name.Namespace == PtOpenXml.pt).ToList())
                    attr.Remove();

                main.PutXDocument();

                // Carry right-only styles + numbering into the left-based package.
                if (main.StyleDefinitionsPart != null &&
                    wDocRight.MainDocumentPart?.StyleDefinitionsPart != null)
                    WmlComparer.CopyMissingStylesFromOneDocToAnother(wDocRight, wDoc);
                WmlComparer.CopyMissingNumberingFromOneDocToAnother(wDocRight, wDoc);
            }
            return streamDoc.GetModifiedWmlDocument();
        }
    }

    // ----------------------------------------------------------------- block-op dispatch

    private static void RenderBlockOp(IrEditOp op, RenderState state, List<XElement> sink)
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
                // Content-equal: emit the RIGHT block verbatim (accepted-state continuity).
                EmitVerbatim(op.RightAnchor, state.Right, state, sink, fromRight: true);
                break;

            case IrEditOpKind.FormatOnlyBlock:
                // Text-equal, format-differing. Task 3 has no rPrChange: emit the right block verbatim.
                EmitVerbatim(op.RightAnchor, state.Right, state, sink, fromRight: true);
                break;

            case IrEditOpKind.InsertBlock:
                EmitWholeBlock(op.RightAnchor, state.Right, state, sink, RevKind.Ins, fromRight: true);
                break;

            case IrEditOpKind.DeleteBlock:
                EmitWholeBlock(op.LeftAnchor, state.Left, state, sink, RevKind.Del, fromRight: false);
                break;

            case IrEditOpKind.ModifyBlock:
                RenderModifyBlock(op, state, sink);
                break;

            case IrEditOpKind.MoveBlock:
            case IrEditOpKind.MoveModifyBlock:
                // Task 4 emits native w:moveFrom/w:moveTo. Task 3 conservative fallback: a move is a
                // delete-here + insert-there pair that round-trips identically. The SOURCE op (left anchor)
                // emits a whole-block del; the DESTINATION op (right anchor) a whole-block ins. They are
                // already placed separately in script order by the builder, so each half renders on its own.
                if (op.IsMoveSource == true)
                    EmitWholeBlock(op.LeftAnchor, state.Left, state, sink, RevKind.Del, fromRight: false);
                else
                    EmitWholeBlock(op.RightAnchor, state.Right, state, sink, RevKind.Ins, fromRight: true);
                break;
        }
    }

    /// <summary>
    /// A Modified pair. A PARAGRAPH pair with a token diff renders finely (per-span run wrapping). Any other
    /// Modified pair (table, opaque, section break, or a paragraph that somehow lacks a token diff) falls back
    /// to a conservative whole-block del(left)+ins(right) that keeps the invariant — Task 4 refines tables.
    /// </summary>
    private static void RenderModifyBlock(IrEditOp op, RenderState state, List<XElement> sink)
    {
        bool leftIsPara = ResolveBlock(op.LeftAnchor, state.Left) is IrParagraph;
        bool rightIsPara = ResolveBlock(op.RightAnchor, state.Right) is IrParagraph;

        if (op.TokenDiff is { } tokenDiff && leftIsPara && rightIsPara &&
            op.TextboxDiffs is null)   // textbox-interior diffs are not finely rendered in Task 3
        {
            RenderModifiedParagraph(op, tokenDiff, state, sink);
            return;
        }

        // Conservative fallback: delete the left block, insert the right block. Order matters only for human
        // reading; accept→right, reject→left both hold. A missing side (shouldn't happen for Modify) is skipped.
        if (op.LeftAnchor != null)
            EmitWholeBlock(op.LeftAnchor, state.Left, state, sink, RevKind.Del, fromRight: false);
        if (op.RightAnchor != null)
            EmitWholeBlock(op.RightAnchor, state.Right, state, sink, RevKind.Ins, fromRight: true);
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
        if (runLevel.Name == W.hyperlink || runLevel.Name == W.sdt || runLevel.Name == W.smartTag)
        {
            var container = new XElement(runLevel.Name, runLevel.Attributes());
            if (kind == RevKind.Ins)
                state.RegisterMediaReferences(container);   // hyperlink r:id rides on the container element
            // Wrap every run-level CHILD; structural children (e.g. sdtPr) pass through untouched.
            foreach (var child in runLevel.Elements())
            {
                if (child.Name == W.r || child.Name == W.hyperlink || child.Name == W.smartTag)
                    container.Add(WrapRunLevel(child, kind, state));
                else
                    container.Add(new XElement(child));
            }
            return container;
        }

        var clone = new XElement(runLevel);
        if (kind == RevKind.Del)
            ConvertTextToDelText(clone);
        var rev = new XElement(kind == RevKind.Ins ? W.ins : W.del, state.RevisionAttributes(), clone);
        if (kind == RevKind.Ins)
            state.RegisterMediaReferences(clone);   // the cloned run is the live tree node media import remaps
        return rev;
    }

    /// <summary>Mark a paragraph's end-of-paragraph mark inserted/deleted: an EMPTY <c>w:ins</c>/<c>w:del</c>
    /// inside <c>w:pPr/w:rPr</c> (the encoding <see cref="RevisionProcessor"/> recognizes — accept of a
    /// deleted mark merges the paragraph with the following one; reject restores it).</summary>
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
            pPr.AddFirst(rPr);   // rPr is the first child of pPr per schema order
        }
        // Remove any pre-existing ins/del marker (idempotence) then add the new one FIRST inside rPr.
        rPr.Elements().Where(e => e.Name == W.ins || e.Name == W.del).Remove();
        rPr.AddFirst(new XElement(kind == RevKind.Ins ? W.ins : W.del, state.RevisionAttributes()));
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
        var rightPara = SourceElement(op.RightAnchor, state.Right);
        if (leftPara == null || rightPara == null)
        {
            // Defensive: fall back to whole-block del+ins if a source element is unexpectedly missing.
            if (op.LeftAnchor != null) EmitWholeBlock(op.LeftAnchor, state.Left, state, sink, RevKind.Del, false);
            if (op.RightAnchor != null) EmitWholeBlock(op.RightAnchor, state.Right, state, sink, RevKind.Ins, true);
            return;
        }

        var leftRuns = new SourceRunModel(leftPara);
        var rightRuns = new SourceRunModel(rightPara);

        // Resolve token char spans: a token op's left span is [left[LeftStart].StartChar, left[LeftEnd-1].EndChar)
        // and likewise right. We resolve via the tokenizers so char coordinates match the diff's exactly.
        var leftTokens = ParagraphTokens(op.LeftAnchor, state.Left, state.Settings);
        var rightTokens = ParagraphTokens(op.RightAnchor, state.Right, state.Settings);

        // The new paragraph: clone the RIGHT paragraph's pPr (accepted-state paragraph properties) and rebuild
        // its run-level content from the spans.
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
                    // Right-side runs as-is (FormatChanged: no rPrChange in Task 3 — documented gap). BUT a
                    // span that is "Equal" by MATCH KEY can still differ in RAW text — the tokenizer conflates
                    // NBSP↔space and case-folds keys, so e.g. a left space vs right NBSP at the same position is
                    // an Equal token op whose raw bytes differ. Emitting the unwrapped right run there would make
                    // reject-all keep the RIGHT byte (NBSP) instead of restoring the LEFT (space). So when the
                    // span's raw left/right text is NOT byte-identical, fall back to del(left)+ins(right) for
                    // that span — the accept/reject invariant then holds byte-for-byte. (This subsumes the
                    // FormatChanged case's text too; only its FORMAT is the documented Task-4 gap, not its text.)
                    var (rs, re) = RightSpanChars(rightTokens, tokenOp);
                    var (ls, le) = LeftSpanChars(leftTokens, tokenOp);
                    string rawRight = RawSpanText(rightTokens, tokenOp.RightStart, tokenOp.RightEnd);
                    string rawLeft = RawSpanText(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd);
                    if (string.Equals(rawLeft, rawRight, StringComparison.Ordinal))
                    {
                        foreach (var r in rightRuns.Slice(rs, re))
                        {
                            state.RegisterMediaReferences(r);
                            content.Add(r);
                        }
                    }
                    else
                    {
                        foreach (var r in leftRuns.Slice(ls, le))
                            content.Add(WrapRunLevel(r, RevKind.Del, state));
                        foreach (var r in rightRuns.Slice(rs, re))
                            content.Add(WrapRunLevel(r, RevKind.Ins, state));   // registers media on its clone
                    }
                    break;
                }
                case IrTokenOpKind.Insert:
                {
                    var (s, e) = RightSpanChars(rightTokens, tokenOp);
                    foreach (var r in rightRuns.Slice(s, e))
                        content.Add(WrapRunLevel(r, RevKind.Ins, state));   // registers media on its clone
                    break;
                }
                case IrTokenOpKind.Delete:
                {
                    var (s, e) = LeftSpanChars(leftTokens, tokenOp);
                    foreach (var r in leftRuns.Slice(s, e))
                        content.Add(WrapRunLevel(r, RevKind.Del, state));
                    break;
                }
            }
        }

        newPara.Add(content);
        sink.Add(newPara);
    }

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
            if (rightHyper.TryGetValue(id, out var hr) && !leftHyper.ContainsKey(id))
            {
                // AddHyperlinkRelationship with the explicit id keeps the cloned w:hyperlink/@r:id resolving. A
                // duplicate-id collision (ArgumentException/InvalidOperationException from the packaging API — the
                // same id already names a DIFFERENT left relationship) is the only expected failure; we leave that
                // reference dangling because the hyperlink TEXT is unchanged, so the ContentHash round-trip still
                // holds (precise rId remap is the Task-4 item). Any OTHER exception propagates.
                try { leftMain.AddHyperlinkRelationship(hr.Uri, hr.IsExternal, id); }
                catch (Exception ex) when (ex is ArgumentException or InvalidOperationException) { }
            }
            else if (rightExternal.TryGetValue(id, out var er) && !leftExternalIds.Contains(id))
            {
                try { leftMain.AddExternalRelationship(er.RelationshipType, er.Uri, id); }
                catch (Exception ex) when (ex is ArgumentException or InvalidOperationException) { }
            }
        }
    }

    /// <summary>True iff this op concerns a standalone section-break block (a `sec:` anchor on either side, or a
    /// resolved <see cref="IrSectionBreak"/>) — the trailing last-section metadata we never emit into the body.</summary>
    private static bool IsSectionBreakOp(IrEditOp op, RenderState state)
    {
        if ((op.RightAnchor?.StartsWith("sec:", StringComparison.Ordinal) ?? false) ||
            (op.LeftAnchor?.StartsWith("sec:", StringComparison.Ordinal) ?? false))
            return true;
        return ResolveBlock(op.RightAnchor, state.Right) is IrSectionBreak ||
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

    private enum RevKind { Ins, Del }

    // ----------------------------------------------------------------- per-call state

    /// <summary>
    /// Mutable per-<see cref="Render"/> state: the two IR snapshots (with provenance), settings, the SINGLE
    /// ascending revision-id counter (no static state), and the live RIGHT-sourced clone roots whose media must
    /// be imported into the left package. One instance per call ⇒ concurrent renders never share a counter.
    /// </summary>
    private sealed class RenderState
    {
        private int _nextId = 1;

        public RenderState(IrDocument left, IrDocument right, IrDiffSettings settings)
        {
            Left = left;
            Right = right;
            Settings = settings;
        }

        public IrDocument Left { get; }
        public IrDocument Right { get; }
        public IrDiffSettings Settings { get; }

        /// <summary>RIGHT-sourced clone roots (in document order) that may carry image relationship references
        /// the LEFT package cannot resolve. After they are placed in the new body (still the same XElement
        /// instances), <see cref="WmlComparer.MoveRelatedPartsToDestination"/> walks each and remaps ids in
        /// place. Only roots actually containing an r-namespace attribute are recorded, so the common
        /// text-only case adds nothing.</summary>
        public List<XElement> RightSourcedClones { get; } = new();

        /// <summary>Fresh (author, id, date) attribute triple for one revision element; id ascends from 1.</summary>
        public object[] RevisionAttributes() => new object[]
        {
            new XAttribute(W.author, Settings.AuthorForRevisions),
            new XAttribute(W.id, _nextId++),
            new XAttribute(W.date, Settings.DateTimeForRevisions),
        };

        /// <summary>Record a RIGHT-sourced clone for media import iff it references any relationship id (an
        /// image embed/link). The recorded element is the live tree node; importing happens post-assembly.</summary>
        public void RegisterMediaReferences(XElement clone)
        {
            if (clone.DescendantsAndSelf().Attributes().Any(a => a.Name.Namespace == R.r))
                RightSourcedClones.Add(clone);
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

        public SourceRunModel(XElement para)
        {
            int charOffset = 0;
            foreach (var child in para.Elements().Where(e => e.Name != W.pPr))
                WalkRunLevel(child, ref charOffset);
        }

        private void WalkRunLevel(XElement runLevel, ref int charOffset)
        {
            if (runLevel.Name == W.r)
            {
                WalkRun(runLevel, ref charOffset);
            }
            else if (runLevel.Name == W.hyperlink || runLevel.Name == W.ins || runLevel.Name == W.del ||
                     runLevel.Name == W.sdt || runLevel.Name == W.smartTag)
            {
                // Container of runs (hyperlink/sdt/smartTag/accepted ins-del wrapper): one ATOMIC segment spanning
                // its full inner text. We do NOT recurse into separate inner segments — the container is emitted
                // whole (so its wrapper survives) and a span boundary never splits inside it. Its char span is
                // the sum of its descendant w:t lengths (mirroring the tokenizer's transparent recursion).
                int start = charOffset;
                foreach (var t in runLevel.Descendants(W.t))
                    charOffset += t.Value.Length;
                _segments.Add(new Segment(runLevel, start, charOffset, SegmentKind.Container));
            }
            else
            {
                // A non-run, non-container run-level element (bookmarkStart/End, proofErr, commentRangeStart…):
                // zero-width, atomic, kept whole.
                _segments.Add(new Segment(runLevel, charOffset, charOffset, SegmentKind.ZeroWidth));
            }
        }

        private void WalkRun(XElement run, ref int charOffset)
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
                    _segments.Add(new Segment(run, start, charOffset, SegmentKind.RunText) { TextChild = child });
                }
                else if (child.Name == W.fldSimple || IsContainer(child.Name))
                {
                    // A simple field's cached result advances the offset by its text (tokenizer recurses too).
                    int start = charOffset;
                    foreach (var t in child.Descendants(W.t))
                        charOffset += t.Value.Length;
                    _segments.Add(new Segment(run, start, charOffset, SegmentKind.RunOther) { OtherChild = child });
                }
                else
                {
                    // tab/break/drawing/noteref/sym/… — zero-width run child.
                    _segments.Add(new Segment(run, charOffset, charOffset, SegmentKind.RunOther) { OtherChild = child });
                }
            }
            if (!any)
                _segments.Add(new Segment(run, charOffset, charOffset, SegmentKind.RunOther));
        }

        private static bool IsContainer(XName n) =>
            n == W.hyperlink || n == W.ins || n == W.del || n == W.sdt || n == W.smartTag;

        /// <summary>Produce run-level XElements covering the half-open char span [start,end). Run children that
        /// fall (partly) inside the span are grouped back into per-source-run <c>w:r</c> clones carrying the
        /// original <c>w:rPr</c>; a straddling <c>w:t</c> is split. Zero-width segments are included iff their
        /// position lies within [start,end) (so a zero-width boundary token attaches to exactly one side).</summary>
        public List<XElement> Slice(int start, int end)
        {
            var result = new List<XElement>();
            // Group consecutive RunText/RunOther segments that share the same source run into one rebuilt w:r.
            XElement? currentRun = null;
            XElement? rebuilt = null;

            void FlushRun()
            {
                if (rebuilt != null && rebuilt.Elements().Any(e => e.Name != W.rPr))
                    result.Add(rebuilt);
                rebuilt = null;
                currentRun = null;
            }

            foreach (var seg in _segments)
            {
                bool overlaps = start == end
                    ? (seg.Start == start && seg.IsZeroWidth)         // empty span: only zero-width at the point
                    : seg.Start < end && seg.End > start ||           // text overlap
                      (seg.IsZeroWidth && seg.Start >= start && seg.Start < end);

                if (!overlaps)
                {
                    if (seg.Kind == SegmentKind.Container || seg.Kind == SegmentKind.ZeroWidth)
                        FlushRun();
                    continue;
                }

                switch (seg.Kind)
                {
                    case SegmentKind.ZeroWidth:
                        FlushRun();
                        result.Add(new XElement(seg.Element));
                        break;

                    case SegmentKind.Container:
                        FlushRun();
                        result.Add(new XElement(seg.Element));
                        break;

                    case SegmentKind.RunText:
                    case SegmentKind.RunOther:
                    {
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
            FlushRun();
            return result;
        }

        private static bool PreserveSpace(string s) =>
            s.Length > 0 && (char.IsWhiteSpace(s[0]) || char.IsWhiteSpace(s[^1]));

        private enum SegmentKind { RunText, RunOther, ZeroWidth, Container }

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
            public bool IsZeroWidth => Start == End;
        }
    }
}
