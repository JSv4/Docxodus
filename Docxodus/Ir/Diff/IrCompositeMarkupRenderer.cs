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

                // NOTE SCOPES (N-way): rewrite reviewer-sourced body note references into the base-anchored
                // output id space, apply the composed per-note diffs inside the footnotes/endnotes parts
                // (creating reviewer-inserted definitions under fresh ids), then run the SAME body-order
                // renumber + cross-kind nested-reference sweep the two-way renderer runs.
                RenderCompositeNoteScopes(script, state, main, baseIr, reviewerIrs, settings);

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

    // ------------------------------------------------------------------ N-way note scopes

    /// <summary>
    /// The composite note-scope phase (the N-way analogue of the two-way renderer's note pipeline):
    /// <list type="number">
    /// <item><b>Reference rewrite.</b> Reviewer-sourced cloned body content carries footnote/endnote
    /// references in THAT reviewer's id space (which diverges from the base space whenever the reviewer's
    /// note numbering shifted, and can collide across reviewers). Using the merger's per-reviewer id maps
    /// (<see cref="IrCompositeScript.NoteIdMaps"/>) plus fresh output ids allocated for reviewer-INSERTED
    /// notes, every reviewer-sourced reference (tracked per reviewer via
    /// <see cref="IrMarkupRenderer.RenderState.NoteRefClonesBySource"/>) is rewritten to the base-anchored
    /// output id space. References inside <c>w:del</c>/<c>w:moveFrom</c> content are skipped — deleted
    /// content clones from the BASE side and already carries base ids.</item>
    /// <item><b>Definition apply.</b> Each composed note diff renders inside the output part: a base-matched
    /// note's blocks are replaced by its composed ops (per-op reviewer sourcing/attribution via
    /// <see cref="RenderOneCompositeBlock"/>); a reviewer-inserted note's definition is cloned from the
    /// reviewer's part under its fresh output id, its ins-marked content rendered from the reviewer.</item>
    /// <item><b>Renumber.</b> The SAME body-order renumber + cross-kind nested-reference sweep the two-way
    /// renderer runs (<see cref="IrMarkupRenderer.RenumberNoteIds"/>/<see cref="IrMarkupRenderer.RemapNestedNoteReferences"/>).</item>
    /// </list>
    /// No-op when no reviewer has any note diff.
    /// </summary>
    private static void RenderCompositeNoteScopes(
        IrCompositeScript script,
        IrMarkupRenderer.RenderState state,
        MainDocumentPart main,
        IrDocument baseIr,
        IReadOnlyList<IrDocument> reviewerIrs,
        IrDiffSettings settings)
    {
        if (script.NoteIdMaps is null && script.NoteOps is null)
            return;

        // ---- 1. Allocate output ids: matched -> base id; inserted -> fresh (deterministic order). ----
        // outputIdByReviewer[(reviewer, kind, reviewerId)] -> output id.
        var outputId = new Dictionary<(int Reviewer, IrNoteKind Kind, string ReviewerId), string>();
        var freshIdByInserted = new Dictionary<(IrNoteKind Kind, int Reviewer, string ReviewerId), string>();
        long nextFresh = MaxPositiveNoteId(main) + 1;
        if (script.NoteIdMaps is { } maps)
        {
            foreach (var m in maps)
            {
                if (m.BaseNoteId is { } baseId)
                {
                    outputId[(m.Reviewer, m.Kind, m.ReviewerNoteId)] = baseId;
                }
                else
                {
                    var fresh = nextFresh.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    nextFresh++;
                    outputId[(m.Reviewer, m.Kind, m.ReviewerNoteId)] = fresh;
                    freshIdByInserted[(m.Kind, m.Reviewer, m.ReviewerNoteId)] = fresh;
                }
            }
        }

        // ---- 2. Rewrite reviewer-sourced body references to the output id space. ----
        foreach (var (reviewer, clones) in state.NoteRefClonesBySource.OrderBy(kvp => kvp.Key))
        {
            if (reviewer < 0)
                continue;   // base-sourced clones already carry base ids
            foreach (var clone in clones)
            {
                foreach (var refEl in clone.DescendantsAndSelf()
                             .Where(e => e.Name == W.footnoteReference || e.Name == W.endnoteReference))
                {
                    // Deleted/moved-from content is BASE-sourced (cloned from the left side) even inside a
                    // reviewer-registered clone — its references already carry base ids.
                    if (refEl.Ancestors().Any(a => a.Name == W.del || a.Name == W.moveFrom))
                        continue;
                    var kind = refEl.Name == W.footnoteReference ? IrNoteKind.Footnote : IrNoteKind.Endnote;
                    var id = (string?)refEl.Attribute(W.id);
                    if (id != null && outputId.TryGetValue((reviewer, kind, id), out var mapped) && mapped != id)
                        refEl.SetAttributeValue(W.id, mapped);
                }
            }
        }
        main.PutXDocument();

        // ---- 3. Apply the composed note diffs inside the parts. ----
        if (script.NoteOps is { Count: > 0 } noteOps)
        {
            ApplyCompositeNoteDiffs(
                noteOps.Where(n => n.Kind == IrNoteKind.Footnote).ToList(), isFootnote: true,
                main, baseIr, reviewerIrs, state, freshIdByInserted, settings);
            ApplyCompositeNoteDiffs(
                noteOps.Where(n => n.Kind == IrNoteKind.Endnote).ToList(), isFootnote: false,
                main, baseIr, reviewerIrs, state, freshIdByInserted, settings);
        }

        // ---- 4. Body-order renumber + cross-kind nested-reference sweep (the two-way pipeline). ----
        var footnoteRemap = IrMarkupRenderer.RenumberNoteIds(main, W.footnoteReference, W.footnote, W.footnotes,
            main.FootnotesPart, null);
        var endnoteRemap = IrMarkupRenderer.RenumberNoteIds(main, W.endnoteReference, W.endnote, W.endnotes,
            main.EndnotesPart, null);
        IrMarkupRenderer.RemapNestedNoteReferences(main, footnoteRemap, endnoteRemap);
    }

    /// <summary>The highest positive <c>w:id</c> across the output package's footnotes AND endnotes parts
    /// (0 when there are none) — the floor for fresh reviewer-inserted note ids, kept above both kinds so a
    /// fresh id can never collide with any base id of either part.</summary>
    private static long MaxPositiveNoteId(MainDocumentPart main)
    {
        long max = 0;
        foreach (var part in new OpenXmlPart?[] { main.FootnotesPart, main.EndnotesPart })
        {
            var root = part?.GetXDocument().Root;
            if (root == null)
                continue;
            foreach (var el in root.Elements())
            {
                if (long.TryParse((string?)el.Attribute(W.id), out var id) && id > max)
                    max = id;
            }
        }
        return max;
    }

    /// <summary>
    /// Apply the composed diffs of ONE note kind inside the output part (creating the part from the
    /// inserting reviewer's boilerplate when the base lacks it): a base-matched diff replaces the note's
    /// block children with its rendered composite ops; a reviewer-inserted diff clones the reviewer's note
    /// shell under its fresh output id and renders the reviewer's (ins-marked) content into it.
    /// </summary>
    private static void ApplyCompositeNoteDiffs(
        List<IrCompositeNoteDiff> diffs, bool isFootnote,
        MainDocumentPart main, IrDocument baseIr, IReadOnlyList<IrDocument> reviewerIrs,
        IrMarkupRenderer.RenderState state,
        IReadOnlyDictionary<(IrNoteKind Kind, int Reviewer, string ReviewerId), string> freshIdByInserted,
        IrDiffSettings settings)
    {
        if (diffs.Count == 0)
            return;
        var noteName = isFootnote ? W.footnote : W.endnote;
        var kind = isFootnote ? IrNoteKind.Footnote : IrNoteKind.Endnote;

        var part = EnsureCompositeNotePart(main, isFootnote, diffs, reviewerIrs);
        var root = part?.GetXDocument().Root;
        if (root == null)
            return;

        bool changed = false;
        foreach (var diff in diffs)
        {
            XElement? noteEl;
            if (diff.BaseNoteId is { } baseId)
            {
                // Base-matched: locate the base-cloned definition by its base id.
                noteEl = root.Elements(noteName).FirstOrDefault(e => (string?)e.Attribute(W.id) == baseId);
                if (noteEl == null)
                    continue;
            }
            else
            {
                // Reviewer-inserted: clone the reviewer's note shell under the fresh output id.
                var shell = ReviewerNoteShell(reviewerIrs, diff.SourceReviewer, kind, diff.ReviewerNoteId);
                if (shell == null || diff.ReviewerNoteId is not { } revId
                    || !freshIdByInserted.TryGetValue((kind, diff.SourceReviewer, revId), out var freshId))
                    continue;
                noteEl = new XElement(noteName, shell.Attributes());
                foreach (var pre in shell.Elements().Where(e => e.Name != W.p && e.Name != W.tbl))
                    noteEl.Add(IrMarkupRenderer.StripUnids(new XElement(pre)));
                noteEl.SetAttributeValue(W.id, freshId);
                root.Add(noteEl);
            }

            // Render the composed ops with the SAME per-op reviewer sourcing the body uses.
            var noteBlocks = new List<XElement>();
            foreach (var op in diff.Ops)
                RenderOneCompositeBlock(op, baseIr, reviewerIrs, state, noteBlocks);
            if (settings is { RenderMoves: true, SimplifyMoveMarkup: true })
                foreach (var b in noteBlocks)
                    IrMarkupRenderer.SimplifyMoveMarkup(b);
            foreach (var b in noteBlocks)
                IrMarkupRenderer.StripUnids(b);

            noteEl.Elements().Where(e => e.Name == W.p || e.Name == W.tbl).Remove();
            noteEl.Add(noteBlocks);
            changed = true;
        }

        if (changed)
        {
            foreach (var attr in root.DescendantsAndSelf().Attributes()
                         .Where(a => a.Name.Namespace == PtOpenXml.pt).ToList())
                attr.Remove();
            part!.PutXDocument();
        }
    }

    /// <summary>Return the output's footnotes/endnotes part, creating one seeded with the FIRST inserting
    /// reviewer's reserved boilerplate notes (ids ≤ 0) when the base package lacks the part but a reviewer
    /// inserted a note of that kind (the composite analogue of the two-way <c>EnsureNotePart</c>). Null only
    /// when there is genuinely no part to apply to.</summary>
    private static OpenXmlPart? EnsureCompositeNotePart(
        MainDocumentPart main, bool isFootnote,
        List<IrCompositeNoteDiff> diffs, IReadOnlyList<IrDocument> reviewerIrs)
    {
        var existing = isFootnote ? (OpenXmlPart?)main.FootnotesPart : main.EndnotesPart;
        if (existing != null)
            return existing;

        // Base has no part: only reviewer-INSERTED notes can require one. Seed boilerplate from the first
        // inserting reviewer's note part root (reached through the reviewer IR's source elements).
        var kind = isFootnote ? IrNoteKind.Footnote : IrNoteKind.Endnote;
        XElement? reviewerRoot = null;
        foreach (var diff in diffs.Where(d => d.BaseNoteId == null))
        {
            reviewerRoot = ReviewerNoteShell(reviewerIrs, diff.SourceReviewer, kind, diff.ReviewerNoteId)
                ?.Document?.Root;
            if (reviewerRoot != null)
                break;
        }
        if (reviewerRoot == null)
            return null;

        var rootName = isFootnote ? W.footnotes : W.endnotes;
        var noteName = isFootnote ? W.footnote : W.endnote;
        var newPart = isFootnote ? (OpenXmlPart)main.AddNewPart<FootnotesPart>() : main.AddNewPart<EndnotesPart>();
        var newRoot = new XElement(rootName, reviewerRoot.Attributes());
        foreach (var note in reviewerRoot.Elements(noteName)
                     .Where(n => int.TryParse((string?)n.Attribute(W.id), out var id) && id <= 0))
            newRoot.Add(IrMarkupRenderer.StripUnids(new XElement(note)));
        var xDoc = newPart.GetXDocument();
        if (xDoc.Root == null)
            xDoc.Add(newRoot);
        else
            xDoc.Root.ReplaceWith(newRoot);
        newPart.PutXDocument();
        return newPart;
    }

    /// <summary>The <c>w:footnote</c>/<c>w:endnote</c> SOURCE element of reviewer
    /// <paramref name="reviewer"/>'s note <paramref name="reviewerNoteId"/>, reached through the reviewer
    /// IR's note store (the note's first block's source element's parent IS the note element, since note
    /// blocks are direct children of the definition). Null when unresolvable.</summary>
    private static XElement? ReviewerNoteShell(
        IReadOnlyList<IrDocument> reviewerIrs, int reviewer, IrNoteKind kind, string? reviewerNoteId)
    {
        if (reviewer < 0 || reviewer >= reviewerIrs.Count || reviewerNoteId == null)
            return null;
        var store = kind == IrNoteKind.Footnote ? reviewerIrs[reviewer].Footnotes : reviewerIrs[reviewer].Endnotes;
        if (!store.Notes.TryGetValue(reviewerNoteId, out var scope) || scope.Blocks.Count == 0)
            return null;
        return scope.Blocks[0].Source.Element?.Parent;
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
