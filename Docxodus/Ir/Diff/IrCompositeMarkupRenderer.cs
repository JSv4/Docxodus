#nullable enable

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Ir;
using Docxodus.Internal;

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
        IrDiffSettings settings,
        IrDocument? sourceBaseIr = null,
        IReadOnlyList<IrDocument>? sourceReviewerIrs = null)
    {
        ArgumentNullException.ThrowIfNull(script);
        ArgumentNullException.ThrowIfNull(baseDoc);
        ArgumentNullException.ThrowIfNull(reviewers);
        ArgumentNullException.ThrowIfNull(settings);

        // Consolidate v1 ceiling (block-format-change family, 2026-07-03): mirror IrCompositeMerger's
        // forcing so the shared two-way emit helpers (which consult state.Settings) never stamp *PrChange
        // markup on a composite render — e.g. on a conflict-path winner op. B1 (sub-project B) turns the
        // PARAGRAPH slice ON so a single-source pPr FormatOnly op stamps w:pPrChange authored to its reviewer;
        // the table-shell/section slices stay OFF (B2). Mirrors IrCompositeMerger's forcing.
        settings = settings with { TrackBlockFormatChanges = false, TrackParagraphFormatChanges = true, TrackTableFormatChanges = true, TrackSectionFormatChanges = true };

        // Base + each reviewer are needed WITH provenance (RetainSources=true) + Accept view — the SAME options
        // the two-way renderer uses — so block anchors in the script resolve to source w:p/w:tbl elements to
        // clone. A caller that already read these exact documents that way passes them in and the reads are
        // skipped; that matters more here than in the two-way case, because re-reading meant 2*(N+1) reads of
        // heavyweight packages for an N-reviewer consolidate. Provenance is equality-neutral (see
        // IrProvenance), so a supplied snapshot renders identically to one read here. A partial or
        // wrong-length hand-off is ignored rather than mixed with fresh reads.
        IrDocument baseIr;
        List<IrDocument> reviewerIrs;
        if (sourceBaseIr is not null && sourceReviewerIrs is not null && sourceReviewerIrs.Count == reviewers.Count)
        {
            baseIr = sourceBaseIr;
            reviewerIrs = sourceReviewerIrs.ToList();
        }
        else
        {
            var readOpts = new IrReaderOptions
            {
                RetainSources = true,
                RevisionView = RevisionView.Accept,
                UnidAssignment = UnidAssignment.IdentityBearing,
            };
            baseIr = IrReader.Read(baseDoc, readOpts);
            reviewerIrs = reviewers.Select(r => IrReader.Read(r.Doc, readOpts)).ToList();
        }

        // One shared RenderState (single ascending id counter, single move-name allocator). Left = base. The
        // RightSource / AuthorOverride / RightSourceId are switched per op below. RightSourcedClones are bucketed
        // by RightSourceId so each reviewer's media-bearing clones import from that reviewer's package.
        var state = new IrMarkupRenderer.RenderState(baseIr, baseIr, settings)
        {
            IsComposite = true,
            CompositeBookmarks = script.Bookmarks,
        };

        // Bookmarks and comment ranges between body blocks go back around the blocks they bracket: the base's
        // where the base has them, each reviewer's where that reviewer has them (issue #1149).
        var authors = reviewers.Select(r => r.Author).ToList();
        var bodyMarkers = IrMarkupRenderer.PlanCompositeBlockLevelMarkers(script.Operations, baseIr, reviewerIrs,
            script.ReviewerComparisons?.Select(o => (IReadOnlyList<IrEditOp>)o.Operations).ToList(), authors, state);
        var bodyBlocks = new List<XElement>();
        for (int i = 0; i < script.Operations.Count; i++)
        {
            if (bodyMarkers?.GetValueOrDefault(i) is { } before)
                bodyBlocks.AddRange(before);
            EmitCompositeOp(script.Operations[i], baseIr, reviewerIrs, state, bodyBlocks);
        }
        if (bodyMarkers?.GetValueOrDefault(script.Operations.Count) is { } trailing)
            bodyBlocks.AddRange(trailing);

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
                // B2: stamp w:sectPrChange on the trailing sectPr for a composed section change (winner's
                // properties applied, base captured into the marker) — the document-level section merge.
                if (trailingSectPr != null)
                    IrMarkupRenderer.ApplyComposedTrailingSectPr(trailingSectPr, script.TrailingSectPr, reviewerIrs, state);
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

                // Whole-block and token clones use the same reviewer choice as explicit marker placement.
                // These are live body nodes; note parts have their own id space and are rendered separately.
                IrMarkupRenderer.ReconcileCompositeBookmarkClones(state, authors);
                state.CompositeBookmarks = null;
                main.PutXDocument();

                // NOTE SCOPES (N-way): rewrite reviewer-sourced body note references into the base-anchored
                // output id space, apply the composed per-note diffs inside the footnotes/endnotes parts
                // (creating reviewer-inserted definitions under fresh ids), then run the SAME body-order
                // renumber + cross-kind nested-reference sweep the two-way renderer runs.
                ReconcileUnchangedCompositeNoteMarkers(script, main, baseIr, reviewerIrs, authors, state);
                RenderCompositeNoteScopes(script, state, main, baseIr, reviewerIrs, authors, settings);

                // Carry each reviewer's missing styles + numbering into the base-based package (continuity for
                // right-only styles / legal numbering referenced by cloned content). Order is reviewer index;
                // first-writer-wins matches the two-way single-right behavior for a one-reviewer consolidate.
                // Same loop merges each reviewer's right-only COMMENT definitions (+ commentsExtended/commentsIds
                // threading) into the base-based comments part — comment markers are AlwaysKeep, so they ride the
                // composite token diff exactly as in two-way Compare and must be reconciled the same way.
                var styleIdRemap = new StyleIdRemap();
                foreach (var reviewer in reviewers)
                {
                    using var revStream = new OpenXmlMemoryStreamDocument(reviewer.Doc);
                    using var wDocRev = revStream.GetWordprocessingDocument();
                    if (main.StyleDefinitionsPart != null &&
                        wDocRev.MainDocumentPart?.StyleDefinitionsPart != null)
                        PackageMerge.CopyMissingStylesFromOneDocToAnother(wDocRev, wDoc, styleIdRemap);
                    // The reviewer's numbering levels and comments, merged after its styles, may name a style
                    // copied under a fresh id (issue #1184).
                    var numberingBeforeCopy = StyleIdRemap.ChildrenSnapshot(main.NumberingDefinitionsPart);
                    PackageMerge.CopyMissingNumberingFromOneDocToAnother(wDocRev, wDoc);
                    styleIdRemap.ApplyToChildrenAddedSince(main.NumberingDefinitionsPart, numberingBeforeCopy);
                    var commentsBeforeMerge = StyleIdRemap.ChildrenSnapshot(main.WordprocessingCommentsPart);
                    IrMarkupRenderer.MergeRightCommentDefinitions(main, wDocRev.MainDocumentPart, streamDoc, revStream);
                    styleIdRemap.ApplyToChildrenAddedSince(main.WordprocessingCommentsPart, commentsBeforeMerge);
                }

                // Reconcile the assembled body's comment markers (unique ids, 1:1 range pairing, exactly-one
                // resolved definition per reference) — the multi-source analogue of the two-way NormalizeComments.
                // Left = base; "right" = the union of every reviewer's anchored comment ids (accept ≡ reviewers).
                var commentAuthors = new Dictionary<string, string>();
                for (int r = 0; r < reviewerIrs.Count; r++)
                    foreach (var id in IrMarkupRenderer.BodyCommentIds(reviewerIrs[r]))
                        commentAuthors.TryAdd(id, authors[r]);
                IrMarkupRenderer.NormalizeComments(main, IrMarkupRenderer.BodyCommentIds(baseIr),
                    commentAuthors.Keys.ToHashSet(), state, commentAuthors);
                IrMarkupRenderer.CompleteCommentThreadingEntries(main);

                // The base's and the reviewers' surviving bookmarks can share ids; renumber them unique,
                // each start kept with its own end (issue #840).
                BookmarkIds.MakeUnique(main);

                // Declare the namespaces the reviewers' cloned content uses on the output part roots —
                // the two-way renderer's closing pass (issue #836), reviewers' prefixes first.
                PartNamespaces.Of(reviewers.Select(r => r.Doc).Append(baseDoc)).DeclareIn(wDoc);

                // Schema child order for every table, row, pPr and rPr — the two-way renderer's closing
                // pass (issue #837).
                WordprocessingMLUtil.OrderChildrenPerSchema(wDoc);

                // Drawings from the base and every reviewer share the story parts — the two-way renderer's
                // drawing id pass (issue #860), after the namespace pass for the same reason.
                DrawingIds.MakeUnique(wDoc);

                // Paragraph ids likewise: a paragraph replaced whole keeps the source paraId in both its deleted and
                // its inserted copy, and merged comments bring ids from another package. Last, after every pass that
                // clones or mints one (issue #1170).
                ParagraphIds.MakeUnique(wDoc);
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
        IReadOnlyList<string> authors,
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
            // A whole-block insertion registers both its runs and its paragraph. Rewrite each live
            // reference once, or a shifted id can be mistaken for another id in the reviewer's map.
            foreach (var refEl in clones.SelectMany(c => c.DescendantsAndSelf())
                         .Where(e => e.Name == W.footnoteReference || e.Name == W.endnoteReference).Distinct())
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
        main.PutXDocument();

        // ---- 3. Apply the composed note diffs inside the parts. ----
        if (script.NoteOps is { Count: > 0 } noteOps)
        {
            ApplyCompositeNoteDiffs(
                noteOps.Where(n => n.Kind == IrNoteKind.Footnote).ToList(), isFootnote: true,
                main, script, baseIr, reviewerIrs, authors, state, freshIdByInserted, outputId, settings);
            ApplyCompositeNoteDiffs(
                noteOps.Where(n => n.Kind == IrNoteKind.Endnote).ToList(), isFootnote: false,
                main, script, baseIr, reviewerIrs, authors, state, freshIdByInserted, outputId, settings);
        }

        // ---- 4. Body-order renumber + cross-kind nested-reference sweep (the two-way pipeline). ----
        var footnoteRemap = IrMarkupRenderer.RenumberNoteIds(main, W.footnoteReference, W.footnote, W.footnotes,
            main.FootnotesPart, null);
        var endnoteRemap = IrMarkupRenderer.RenumberNoteIds(main, W.endnoteReference, W.endnote, W.endnotes,
            main.EndnotesPart, null);
        IrMarkupRenderer.RemapNestedNoteReferences(main, footnoteRemap, endnoteRemap);
    }

    /// <summary>
    /// The output keeps the base's copy of every note no reviewer changed, so a reviewer's bookmark or comment
    /// range between such a note's blocks (invisible to the IR, rule N3) was lost (issue #1175). Each such note is
    /// reconciled against every reviewer's copy of it, as an equal container is
    /// (<see cref="IrMarkupRenderer.ReconcileCompositeEqualContainerMarkers"/>): a reviewer's note is the one its
    /// note-id map pairs with the base note, or the one with the same id when that reviewer's comparison changed
    /// no note. Runs before the note diffs apply and renumber, while the output holds the base's ids.
    /// </summary>
    private static void ReconcileUnchangedCompositeNoteMarkers(
        IrCompositeScript script, MainDocumentPart main, IrDocument baseIr, IReadOnlyList<IrDocument> reviewerIrs,
        IReadOnlyList<string> authors, IrMarkupRenderer.RenderState state)
    {
        foreach (var (part, kind, noteName) in new (OpenXmlPart?, IrNoteKind, XName)[]
                 {
                     (main.FootnotesPart, IrNoteKind.Footnote, W.footnote),
                     (main.EndnotesPart, IrNoteKind.Endnote, W.endnote),
                 })
        {
            if (part?.GetXDocument().Root is not { } root)
                continue;
            var touched = (script.NoteOps ?? IrNodeList.Empty<IrCompositeNoteDiff>())
                .Where(d => d.Kind == kind && d.BaseNoteId is not null).Select(d => d.BaseNoteId!)
                .ToHashSet(StringComparer.Ordinal);
            bool changed = false;
            foreach (var note in root.Elements(noteName).ToList())
            {
                if ((string?)note.Attribute(W.id) is not { } id || touched.Contains(id) ||
                    NoteShell(baseIr, kind, id) is not { } baseNote)
                    continue;
                var counterparts = new List<(int Reviewer, XElement Counterpart)>();
                for (int r = 0; r < reviewerIrs.Count; r++)
                    if (ReviewerNoteId(script, r, kind, id) is { } reviewerId && NoteShell(reviewerIrs[r], kind, reviewerId) is { } reviewerNote &&
                        IrMarkupRenderer.BlockLevelMarkersDiffer(baseNote, reviewerNote))
                        counterparts.Add((r, reviewerNote));
                if (counterparts.Count == 0)
                    continue;
                IrMarkupRenderer.ReconcileCompositeEqualContainerMarkers(note, baseNote, counterparts, authors, state);
                changed = true;
            }
            if (changed)
                part.PutXDocument();
        }
    }

    /// <summary>The id of reviewer <paramref name="reviewer"/>'s copy of base note <paramref name="baseId"/>: the
    /// one its note-id map pairs with it, or the same id when the map leaves that id alone (an id absent from a
    /// reviewer's map is identity, <see cref="IrReviewerNoteIdMap"/>). Null when the reviewer has no copy.</summary>
    private static string? ReviewerNoteId(IrCompositeScript script, int reviewer, IrNoteKind kind, string baseId)
    {
        if (!NoteIdLookups.GetValue(script, BuildNoteIdLookup).TryGetValue((reviewer, kind), out var map))
            return baseId;
        if (map.ByBase.TryGetValue(baseId, out var mapped))
            return mapped;
        return map.ReviewerIds.Contains(baseId) ? null : baseId;
    }

    /// <summary>Each reviewer's note-id map by kind: base id → reviewer id, and every reviewer id the map names.</summary>
    private static Dictionary<(int Reviewer, IrNoteKind Kind), (Dictionary<string, string> ByBase, HashSet<string> ReviewerIds)>
        BuildNoteIdLookup(IrCompositeScript script)
    {
        var lookup = new Dictionary<(int, IrNoteKind), (Dictionary<string, string>, HashSet<string>)>();
        foreach (var m in script.NoteIdMaps ?? IrNodeList.Empty<IrReviewerNoteIdMap>())
        {
            if (!lookup.TryGetValue((m.Reviewer, m.Kind), out var map))
                lookup[(m.Reviewer, m.Kind)] = map = (new Dictionary<string, string>(StringComparer.Ordinal),
                    new HashSet<string>(StringComparer.Ordinal));
            if (m.BaseNoteId is { } baseId)
                map.Item1.TryAdd(baseId, m.ReviewerNoteId);
            map.Item2.Add(m.ReviewerNoteId);
        }
        return lookup;
    }

    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<IrCompositeScript,
        Dictionary<(int Reviewer, IrNoteKind Kind), (Dictionary<string, string> ByBase, HashSet<string> ReviewerIds)>>
        NoteIdLookups = new();

    /// <summary>Reviewer <paramref name="reviewer"/>'s own comparison of base note <paramref name="baseId"/>, for
    /// placing that reviewer's markers between the note's blocks: the reviewer's note diff, or, for a note the
    /// reviewer left equal but whose markers differ, equal ops pairing the two notes block for block; empty
    /// otherwise.</summary>
    private static IReadOnlyList<IrEditOp> ReviewerNoteOps(IrCompositeScript script, int reviewer, IrNoteKind kind,
        string baseId, IrDocument baseIr, IrDocument reviewerIr)
    {
        if (script.ReviewerComparisons is not { } scripts || reviewer >= scripts.Count)
            return Array.Empty<IrEditOp>();
        if (scripts[reviewer].NoteOps?.FirstOrDefault(n => n.Kind == kind && n.LeftNoteId == baseId) is { } own)
            return own.Ops;
        var baseStore = kind == IrNoteKind.Footnote ? baseIr.Footnotes : baseIr.Endnotes;
        var store = kind == IrNoteKind.Footnote ? reviewerIr.Footnotes : reviewerIr.Endnotes;
        if (ReviewerNoteId(script, reviewer, kind, baseId) is not { } reviewerId ||
            !baseStore.Notes.TryGetValue(baseId, out var baseNote) || !store.Notes.TryGetValue(reviewerId, out var note) ||
            baseNote.Blocks.Count != note.Blocks.Count ||
            !IrMarkupRenderer.BlockLevelMarkersDiffer(NoteShell(baseIr, kind, baseId), NoteShell(reviewerIr, kind, reviewerId)))
            return Array.Empty<IrEditOp>();
        return baseNote.Blocks.Zip(note.Blocks, (b, r) =>
            new IrEditOp(IrEditOpKind.EqualBlock, b.Anchor.ToString(), r.Anchor.ToString(), null, null, null)).ToList();
    }

    /// <summary>The source <c>w:footnote</c>/<c>w:endnote</c> of note <paramref name="id"/> in
    /// <paramref name="ir"/>: its first block's parent. Null for a note with no block or no source.</summary>
    private static XElement? NoteShell(IrDocument ir, IrNoteKind kind, string id)
    {
        var store = kind == IrNoteKind.Footnote ? ir.Footnotes : ir.Endnotes;
        return store.Notes.TryGetValue(id, out var scope) && scope.Blocks.Count > 0 ? scope.Blocks[0].Source.Element?.Parent : null;
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
        MainDocumentPart main, IrCompositeScript script, IrDocument baseIr, IReadOnlyList<IrDocument> reviewerIrs,
        IReadOnlyList<string> authors,
        IrMarkupRenderer.RenderState state,
        IReadOnlyDictionary<(IrNoteKind Kind, int Reviewer, string ReviewerId), string> freshIdByInserted,
        IReadOnlyDictionary<(int Reviewer, IrNoteKind Kind, string ReviewerId), string> outputId,
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
                noteEl = root.Elements(noteName).FirstOrDefault(e => (string?)e.Attribute(W.id) == baseId
                    && !WmlToMarkdownConverter.IsBoilerplateNote(e));
                if (noteEl == null)
                    continue;
            }
            else
            {
                // Reviewer-inserted: clone the reviewer's note shell under the fresh output id. A note with no
                // blocks has no shell to reach (<w:footnote w:id="1"/>); it still gets its definition, or its
                // inserted reference dangles (issue #840).
                if (diff.ReviewerNoteId is not { } revId
                    || !freshIdByInserted.TryGetValue((kind, diff.SourceReviewer, revId), out var freshId))
                    continue;
                var shell = ReviewerNoteShell(reviewerIrs, diff.SourceReviewer, kind, revId);
                noteEl = new XElement(noteName, shell?.Attributes());
                // The shell's bookmarks and comment ranges between blocks go among the blocks, below.
                foreach (var pre in shell?.Elements().Where(e => !IrMarkupRenderer.IsBlockContent(e) &&
                             !IrMarkupRenderer.IsBlockLevelMarker(e)) ?? Enumerable.Empty<XElement>())
                    noteEl.Add(IrMarkupRenderer.StripUnids(new XElement(pre)));
                noteEl.SetAttributeValue(W.id, freshId);
                root.Add(noteEl);
            }

            // Render the composed ops with the SAME per-op reviewer sourcing the body uses. A base note's
            // bookmarks and comment ranges between blocks stay where the base has them (issue #1149), and each
            // reviewer's go in where that reviewer's comparison of the note puts them; an inserted note's are its
            // reviewer's, inserted with its blocks: that note's ops are the reviewer's own comparison of the
            // note (issue #1175).
            var noteMarkers = IrMarkupRenderer.PlanCompositeBlockLevelMarkers(diff.Ops, baseIr, reviewerIrs,
                Enumerable.Range(0, reviewerIrs.Count).Select(r => diff.BaseNoteId is { } noteId
                    ? ReviewerNoteOps(script, r, kind, noteId, baseIr, reviewerIrs[r])
                    : r == diff.SourceReviewer ? diff.Ops.Select(o => o.Op).ToList() : (IReadOnlyList<IrEditOp>)Array.Empty<IrEditOp>())
                    .ToList(),
                authors, state);
            if (noteMarkers is not null)
                noteEl.Elements().Where(IrMarkupRenderer.IsBlockLevelMarker).Remove();
            var noteBlocks = new List<XElement>();
            for (int i = 0; i < diff.Ops.Count; i++)
            {
                if (noteMarkers?.GetValueOrDefault(i) is { } before)
                    noteBlocks.AddRange(before);
                RenderOneCompositeBlock(diff.Ops[i], baseIr, reviewerIrs, state, noteBlocks);
            }
            if (noteMarkers?.GetValueOrDefault(diff.Ops.Count) is { } trailing)
                noteBlocks.AddRange(trailing);
            if (settings is { RenderMoves: true, SimplifyMoveMarkup: true })
                foreach (var b in noteBlocks)
                    IrMarkupRenderer.SimplifyMoveMarkup(b);
            foreach (var b in noteBlocks)
                IrMarkupRenderer.StripUnids(b);

            noteEl.Elements().Where(IrMarkupRenderer.IsBlockContent).Remove();
            noteEl.Add(noteBlocks);

            // A reviewer-INSERTED note is entirely reviewer-sourced (all-ins content, no base/del side), so every
            // nested note reference in its body carries THAT reviewer's id space. Rewrite each to the output id
            // space (matched -> base id, inserted -> fresh id) so the body-order renumber sweep below then carries
            // it to the final id. Without this, a nested reference to ANOTHER note the same reviewer inserted keeps
            // the reviewer id and dangles: step 2's reference rewrite only visits body clones, not inserted-note
            // definition bodies, and step 4's remap is keyed on the output-old id, which the reviewer id is not.
            if (diff.BaseNoteId == null)
            {
                foreach (var refEl in noteEl.Descendants()
                             .Where(e => e.Name == W.footnoteReference || e.Name == W.endnoteReference))
                {
                    var refKind = refEl.Name == W.footnoteReference ? IrNoteKind.Footnote : IrNoteKind.Endnote;
                    var refId = (string?)refEl.Attribute(W.id);
                    if (refId != null
                        && outputId.TryGetValue((diff.SourceReviewer, refKind, refId), out var mapped)
                        && mapped != refId)
                        refEl.SetAttributeValue(W.id, mapped);
                }
            }

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
    /// reviewer's reserved boilerplate notes when the base package lacks the part but a reviewer
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
        if (reviewerRoot == null && !diffs.Any(d => d.BaseNoteId == null))
            return null;

        var rootName = isFootnote ? W.footnotes : W.endnotes;
        var noteName = isFootnote ? W.footnote : W.endnote;
        var newPart = isFootnote
            ? (OpenXmlPart)main.AddDeterministicPart<FootnotesPart>("rIdFootnotes")
            : main.AddDeterministicPart<EndnotesPart>("rIdEndnotes");
        // Only notes without blocks were inserted: no reviewer root is reachable, so start from a bare one.
        var newRoot = new XElement(rootName,
            reviewerRoot?.Attributes() ?? new[] { new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName) });
        foreach (var note in (reviewerRoot?.Elements(noteName) ?? Enumerable.Empty<XElement>())
                     .Where(WmlToMarkdownConverter.IsBoilerplateNote))
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

        int emittedFrom = sink.Count;
        if (compositeOp.AuthoredRows is not null)
            IrMarkupRenderer.RenderComposedTable(compositeOp, baseIr, reviewerIrs, state, sink, RenderOneCompositeBlock);
        else if (compositeOp.AuthoredWrapperOps is { } wrapperOps)
        {
            var children = new List<XElement>();
            foreach (var child in wrapperOps)
                RenderOneCompositeBlock(child, baseIr, reviewerIrs, state, children);
            IrMarkupRenderer.EmitModifiedWrapper(op, op.WrapperOps!, state, sink, children);
        }
        else if (compositeOp.AuthoredTokens is not null)
            IrMarkupRenderer.RenderComposedParagraph(compositeOp, baseIr, reviewerIrs, state, sink);
        else
            IrMarkupRenderer.RenderBlockOp(op, state, sink);
        if (compositeOp.EqualReviewerMarkers is { } reviewerMarkers && sink.Count == emittedFrom + 1)
        {
            if (baseSourced)
                foreach (var reviewer in reviewerMarkers)
                    if (reviewer.Reviewer >= 0 && reviewer.Reviewer < reviewerIrs.Count)
                        IrMarkupRenderer.CarryMissingComments(sink[emittedFrom],
                            IrMarkupRenderer.SourceElement(reviewer.Anchor, reviewerIrs[reviewer.Reviewer]));
            if (sink[emittedFrom].Name == W.p && (baseSourced || op.TokenDiff is not null))
                IrMarkupRenderer.CarryReviewerOnlyBookmarks(sink[emittedFrom], reviewerMarkers, baseIr, reviewerIrs, state);
            else if (baseSourced && IrMarkupRenderer.SourceElement(op.LeftAnchor, baseIr) is { } baseSource)
                IrMarkupRenderer.ReconcileCompositeEqualContainerMarkers(sink[emittedFrom], baseSource,
                    reviewerMarkers.Where(b => b.Reviewer >= 0 && b.Reviewer < reviewerIrs.Count)
                        .Select(b => (b.Reviewer, IrMarkupRenderer.SourceElement(b.Anchor, reviewerIrs[b.Reviewer])))
                        .Where(c => c.Item2 is not null).Select(c => (c.Reviewer, c.Item2!)).ToList(),
                    AuthorsByReviewer(reviewerMarkers, reviewerIrs.Count), state);
        }

        // Restore the shared state so a CELL recursion does not leak its per-op source into the surrounding
        // table-row render (the body loop reassigns these every op, so the restore is inert there).
        state.AuthorOverride = savedAuthor;
        state.RightSource = savedRightSource;
        state.RightSourceId = savedRightSourceId;
    }

    /// <summary>Each reviewer's author by reviewer index, from the entries that name them.</summary>
    private static IReadOnlyList<string> AuthorsByReviewer(IEnumerable<IrEqualReviewerMarkers> entries, int reviewers)
    {
        var authors = new string[reviewers];
        foreach (var entry in entries)
            if (entry.Reviewer >= 0 && entry.Reviewer < reviewers)
                authors[entry.Reviewer] = entry.Author;
        return authors;
    }
}
