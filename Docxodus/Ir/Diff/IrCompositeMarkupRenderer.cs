#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Ir;

namespace Docxodus.Ir.Diff;

/// <summary>
/// Renders an <see cref="IrCompositeScript"/> (the N-way merge of several reviewers' pairwise edit scripts
/// against one shared base) into a SINGLE native tracked-revisions <see cref="WmlDocument"/>. The diff-side
/// counterpart to <see cref="IrMarkupRenderer"/> for the multi-reviewer consolidate path.
/// <para><b>Output base + per-reviewer sourcing.</b> The output package is assembled on the BASE document's
/// package (styles/numbering/fonts/settings/theme/base media carry over by reuse — exactly as the two-way
/// renderer builds on LEFT). Inserted/modified ("right-side") block elements and token text are cloned from
/// the CONTRIBUTING REVIEWER's package, selected per op via <see cref="IrCompositeOp.SourceReviewer"/>: the
/// renderer points the shared <see cref="IrMarkupRenderer.RenderState.RightSource"/> at that reviewer's IR
/// (or the base, for a base-sourced equal/delete) and sets
/// <see cref="IrMarkupRenderer.RenderState.AuthorOverride"/> to the reviewer's author so every emitted
/// revision is attributed correctly. The reused emit helpers therefore behave per-reviewer with no change to
/// the two-way path (where <c>RightSource == Right</c> always).</para>
/// <para><b>Scope (Task 2.2).</b> EqualBlock / InsertBlock / DeleteBlock / single-reviewer ModifyBlock — plus
/// single-reviewer Move/Split/Merge passthrough (these arise from one reviewer in v1). The COMPOSED
/// multi-reviewer ModifyBlock branch (<see cref="IrCompositeOp.AuthoredTokens"/> non-null) is handled by
/// <see cref="IrMarkupRenderer.RenderComposedParagraph"/>, which reuses the two-way token-span emit helpers
/// to attribute each span to its contributing reviewer. Conflict-policy body shaping is Task 2.4.</para>
/// <para><b>Invariant.</b> Reject-all yields the BASE document's content; accept-all yields the per-reviewer
/// accepted edits — the multi-author generalization of the two-way accept≡right / reject≡left contract.</para>
/// </summary>
internal static class IrCompositeMarkupRenderer
{
    /// <summary>
    /// Render <paramref name="script"/> into a tracked-revisions <see cref="WmlDocument"/> on
    /// <paramref name="baseDoc"/>'s package. <paramref name="reviewers"/> supplies, in
    /// <see cref="IrCompositeOp.SourceReviewer"/> index order, each reviewer's author + original document
    /// (the source the reviewer's inserted/modified content is cloned from). <paramref name="settings"/>
    /// supplies date/granularity (per-op author comes from the reviewer, not <c>settings</c>).
    /// </summary>
    public static WmlDocument Render(
        IrCompositeScript script,
        WmlDocument baseDoc,
        IReadOnlyList<(string Author, WmlDocument Doc)> reviewers,
        IrDiffSettings settings)
    {
        ArgumentNullException.ThrowIfNull(script);
        ArgumentNullException.ThrowIfNull(baseDoc);
        ArgumentNullException.ThrowIfNull(reviewers);
        ArgumentNullException.ThrowIfNull(settings);

        // v1 limitation: note-scope (footnote/endnote) merging is not yet implemented in the composite path.
        // The merger never populates NoteOps today (IrCompositeMerger fails fast on a reviewer note edit), so
        // this guard never fires; it is a RUNTIME check (not Debug.Assert — which is stripped in the Release
        // build this library ships in and CI tests under) so a future change that starts populating NoteOps
        // without renderer support fails loudly instead of silently dropping note diffs.
        if (script.NoteOps is { Count: > 0 })
            throw new System.InvalidOperationException(
                "IrCompositeMarkupRenderer does not yet render composite NoteOps; note merging is a follow-on.");

        // Re-read base + each reviewer WITH provenance (RetainSources=true) + Accept view — the SAME options the
        // two-way renderer uses — so block anchors in the script resolve to source w:p/w:tbl elements to clone.
        var readOpts = new IrReaderOptions { RetainSources = true, RevisionView = RevisionView.Accept };
        var baseIr = IrReader.Read(baseDoc, readOpts);
        var reviewerIrs = reviewers.Select(r => IrReader.Read(r.Doc, readOpts)).ToList();

        // One shared RenderState (single ascending id counter, single move-name allocator). Left = base. The
        // RightSource / AuthorOverride / RightSourceId are switched per op below. RightSourcedClones are bucketed
        // by RightSourceId so each reviewer's media-bearing clones import from that reviewer's package.
        var state = new IrMarkupRenderer.RenderState(baseIr, baseIr, settings);

        var bodyBlocks = new List<XElement>();
        foreach (var op in script.Operations)
            EmitCompositeOp(op, baseIr, reviewerIrs, state, bodyBlocks);

        // SimplifyMoveMarkup post-pass (mirrors the two-way renderer): rewrite native move markup to del/ins.
        if (settings is { RenderMoves: true, SimplifyMoveMarkup: true })
            foreach (var block in bodyBlocks)
                IrMarkupRenderer.SimplifyMoveMarkup(block);

        // Assemble on a clone of the BASE package, preserving its trailing top-level w:sectPr, then copy each
        // reviewer's missing styles/numbering for continuity (the multi-source analogue of the two-way pass).
        var result = new WmlDocument(baseDoc);
        using (var streamDoc = new OpenXmlMemoryStreamDocument(result))
        {
            using (var wDoc = streamDoc.GetWordprocessingDocument())
            {
                var main = wDoc.MainDocumentPart
                    ?? throw new DocxodusException("Base document has no MainDocumentPart.");
                var mainXDoc = main.GetXDocument();
                var bodyEl = mainXDoc.Root?.Element(W.body)
                    ?? throw new DocxodusException("Base document has no w:body.");

                var trailingSectPr = bodyEl.Elements(W.sectPr).LastOrDefault();
                bodyEl.Elements().Where(e => e.Name != W.sectPr).Remove();
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

                // Per-reviewer media import: each reviewer's media-bearing clones live in their own bucket
                // (keyed by reviewer index); import each from THAT reviewer's package. Opening every reviewer
                // package only when it actually has clones to import keeps the text-only common case cheap.
                foreach (var (sourceId, clones) in state.RightSourcedClonesBySource.OrderBy(kvp => kvp.Key))
                {
                    if (sourceId < 0 || sourceId >= reviewers.Count || clones.Count == 0)
                        continue;   // base-sourced (-1) clones reference base parts already present; nothing to import.
                    using var revStream = new OpenXmlMemoryStreamDocument(reviewers[sourceId].Doc);
                    using var wDocRev = revStream.GetWordprocessingDocument();
                    IrMarkupRenderer.ImportRightSourcedMedia(clones, main, wDocRev.MainDocumentPart, streamDoc, revStream);
                }

                // Strip ALL engine-internal pt bookkeeping from the assembled body.
                foreach (var attr in bodyEl.DescendantsAndSelf().Attributes()
                             .Where(a => a.Name.Namespace == PtOpenXml.pt).ToList())
                    attr.Remove();

                main.PutXDocument();

                // Carry each reviewer's missing styles + numbering into the base-based package (continuity for
                // right-only styles / legal numbering referenced by cloned content). Order is reviewer index;
                // first-writer-wins matches the two-way single-right behavior for a one-reviewer consolidate.
                // Same loop merges each reviewer's right-only COMMENT definitions (+ commentsExtended/commentsIds
                // threading) into the base-based comments part — comment markers are AlwaysKeep, so they ride the
                // composite token diff exactly as in two-way Compare and must be reconciled the same way.
                foreach (var reviewer in reviewers)
                {
                    using var revStream = new OpenXmlMemoryStreamDocument(reviewer.Doc);
                    using var wDocRev = revStream.GetWordprocessingDocument();
                    if (main.StyleDefinitionsPart != null &&
                        wDocRev.MainDocumentPart?.StyleDefinitionsPart != null)
                        WmlComparer.CopyMissingStylesFromOneDocToAnother(wDocRev, wDoc);
                    WmlComparer.CopyMissingNumberingFromOneDocToAnother(wDocRev, wDoc);
                    IrMarkupRenderer.MergeRightCommentDefinitions(main, wDocRev.MainDocumentPart);
                }

                // Reconcile the assembled body's comment markers (unique ids, 1:1 range pairing, exactly-one
                // resolved definition per reference) — the multi-source analogue of the two-way NormalizeComments.
                // Left = base; "right" = the union of every reviewer's anchored comment ids (accept ≡ reviewers).
                var reviewerCommentIds = new HashSet<string>();
                foreach (var rir in reviewerIrs)
                    foreach (var id in IrMarkupRenderer.BodyCommentIds(rir))
                        reviewerCommentIds.Add(id);
                IrMarkupRenderer.NormalizeComments(main, IrMarkupRenderer.BodyCommentIds(baseIr),
                    reviewerCommentIds, state);
            }
            return streamDoc.GetModifiedWmlDocument();
        }
    }

    /// <summary>
    /// Dispatch one composite op: point the shared <see cref="IrMarkupRenderer.RenderState"/> at the
    /// contributing reviewer (author override + right source IR + media bucket), then reuse the two-way
    /// renderer's per-op emit dispatch (<see cref="IrMarkupRenderer.RenderBlockOp"/>). A composed
    /// multi-reviewer ModifyBlock is handled by <see cref="IrMarkupRenderer.RenderComposedParagraph"/>.
    /// </summary>
    private static void EmitCompositeOp(
        IrCompositeOp compositeOp,
        IrDocument baseIr,
        IReadOnlyList<IrDocument> reviewerIrs,
        IrMarkupRenderer.RenderState state,
        List<XElement> sink)
    {
        // Composed multi-reviewer paragraph: per-span authorship from AuthoredTokens, each contributing
        // reviewer's right paragraph resolved via SourceRightAnchors. Delegates to the two-way renderer's
        // composed-paragraph builder so it can reuse the private token-span emit helpers (Task 2.3).
        if (compositeOp.AuthoredTokens != null)
        {
            IrMarkupRenderer.RenderComposedParagraph(compositeOp, baseIr, reviewerIrs, state, sink);
            return;
        }

        // Composed multi-reviewer TABLE (FOLLOW-ON B): per-row/per-cell authorship from AuthoredRows.
        // Delegates to the two-way renderer's composed-table builder, which recursively renders each composed
        // cell-block via the shared RenderOneCompositeBlock helper below.
        if (compositeOp.AuthoredRows != null)
        {
            IrMarkupRenderer.RenderComposedTable(compositeOp, baseIr, reviewerIrs, state, sink,
                RenderOneCompositeBlock);
            return;
        }

        RenderOneCompositeBlock(compositeOp, baseIr, reviewerIrs, state, sink);
    }

    /// <summary>
    /// Render ONE composite op into <paramref name="sink"/> by pointing the shared
    /// <see cref="IrMarkupRenderer.RenderState"/> at the contributing reviewer (author override + right source
    /// IR + media bucket) and reusing the two-way per-op dispatch — OR, for a composed paragraph
    /// (<see cref="IrCompositeOp.AuthoredTokens"/> non-null), delegating to
    /// <see cref="IrMarkupRenderer.RenderComposedParagraph"/>. Shared by the body loop
    /// (<see cref="EmitCompositeOp"/>) and the per-cell recursion inside
    /// <see cref="IrMarkupRenderer.RenderComposedTable"/>: a composed table cell's block ops are themselves
    /// composite ops (single-source or composed paragraph), rendered here exactly as a body block would be.
    /// </summary>
    internal static void RenderOneCompositeBlock(
        IrCompositeOp compositeOp,
        IrDocument baseIr,
        IReadOnlyList<IrDocument> reviewerIrs,
        IrMarkupRenderer.RenderState state,
        List<XElement> sink)
    {
        // A composed paragraph nested in a cell: per-span authorship (recursion proof for same-cell-paragraph
        // token composition).
        if (compositeOp.AuthoredTokens != null)
        {
            IrMarkupRenderer.RenderComposedParagraph(compositeOp, baseIr, reviewerIrs, state, sink);
            return;
        }

        var op = compositeOp.Op;
        int sourceReviewer = compositeOp.SourceReviewer;

        // A base-sourced op (SourceReviewer -1): EqualBlock emits the base block verbatim; the right source is the
        // base IR and there is no author override. A reviewer-sourced op points RightSource at that reviewer's IR
        // and overrides the author to the reviewer; media-bearing clones land in that reviewer's bucket.
        bool baseSourced = sourceReviewer < 0;
        var savedAuthor = state.AuthorOverride;
        var savedRightSource = state.RightSource;
        var savedRightSourceId = state.RightSourceId;

        state.AuthorOverride = baseSourced ? null : compositeOp.Author;
        state.RightSource = baseSourced ? baseIr : reviewerIrs[sourceReviewer];
        state.RightSourceId = sourceReviewer;

        IrMarkupRenderer.RenderBlockOp(op, state, sink);

        // Restore the shared state so a CELL recursion does not leak its per-op source into the surrounding
        // table-row render (the body loop reassigns these every op, so the restore is inert there).
        state.AuthorOverride = savedAuthor;
        state.RightSource = savedRightSource;
        state.RightSourceId = savedRightSourceId;
    }
}
