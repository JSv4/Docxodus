#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
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

    /// <summary>TRANSIENT marker attribute carrying a source <c>w:hyperlink</c>'s RESOLVED target (external URI,
    /// or <c>"#" + anchor</c> for an internal link) onto each emitted wrapper clone. It lets
    /// <see cref="CoalesceAdjacentHyperlinks"/> merge a fully-replaced single link's pure <c>w:del</c>+<c>w:ins</c>
    /// fragments (same target on both sides — #232) while keeping the WC019 whole-anchor RETARGET separate (the
    /// del/ins fragments carry the SAME <c>r:id</c> STRING at coalesce time but DIFFERENT resolved targets, so a
    /// string compare cannot tell them apart — the resolved URI can). Del fragments come from the LEFT model and
    /// carry the LEFT target; ins/plain fragments come from the RIGHT model and carry the RIGHT target; a run
    /// merges only when every fragment's target agrees. In the <c>pt:</c> namespace (blanket-stripped), and
    /// removed explicitly before output regardless.</summary>
    private static readonly XName SourceLinkTarget = PtOpenXml.pt + "SourceLinkTarget";

    /// <summary>Identifies a style definition independently of the source <see cref="XElement"/>.
    /// Style ids are typed in OOXML, so the type participates in right-import provenance.</summary>
    private readonly record struct StyleIdentity(string Type, string Id);

    /// <summary>
    /// A deliberately scoped, reversible package-presentation plan. OOXML has no native tracked revision for
    /// <c>docDefaults</c>, theme, numbering, or settings parts, so the output keeps the left package parts and
    /// materializes only a proven docDefaults-only delta into the CURRENT payload of shared style definitions.
    /// The old effective payload is carried by the standard style-level pPrChange/rPrChange markers, making
    /// accept/reject switch presentation without a hidden package swap.
    /// </summary>
    private sealed record DocDefaultsStyleProjection(HashSet<StyleIdentity> Styles)
    {
        public bool Includes(string? type, string? id) =>
            type is not null && id is not null && Styles.Contains(new StyleIdentity(type, id));
    }

    // Mirrors the tail of DocxSession.PPrChildOrder after w:spacing. When a right-only style
    // delegates paragraph spacing to defaults, the materialized w:spacing must precede the first
    // present CT_PPrBase child in this set (for example w:ind or w:jc), not merely w:outlineLvl.
    private static readonly HashSet<XName> PPrChildrenAfterSpacing = new()
    {
        W.ind,
        W.contextualSpacing,
        W.mirrorIndents,
        W.suppressOverlap,
        W.jc,
        W.textDirection,
        W.textAlignment,
        W.textboxTightWrap,
        W.outlineLvl,
        W.divId,
        W.cnfStyle,
        W.rPr,
        W.sectPr,
        W.pPrChange,
    };

    // Mirrors the tail of PtOpenXmlUtil.Order_rPr after w:kern. Synthesized kerning must remain
    // between w:w and w:position even when a copied style has no explicit size; appending it after
    // w:lang, for example, produces schema-invalid CT_RPr ordering.
    private static readonly HashSet<XName> RPrChildrenAfterKern = new()
    {
        W.position,
        W.sz,
        W14.wShadow,
        W14.wTextOutline,
        W14.wTextFill,
        W14.wScene3d,
        W14.wProps3d,
        W.szCs,
        W.highlight,
        W.u,
        W.effect,
        W.bdr,
        W.shd,
        W.fitText,
        W.vertAlign,
        W.rtl,
        W.cs,
        W.em,
        W.lang,
        W.eastAsianLayout,
        W.specVanish,
        W.oMath,
    };

    // Schema order for the properties produced by the reversible docDefaults projection. Effective formatting
    // is assembled from several cascade layers, so a later override can otherwise be appended after an earlier
    // property with a higher schema position (for example an inherited w:rFonts after w:b). Keep unknown/
    // extension children stable at the tail; the v1 projection admits only literal WordprocessingML properties.
    private static readonly string[] StylePPrChildOrder =
    {
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
        "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
        "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
        "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
        "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
        "sectPr", "pPrChange",
    };

    private static readonly string[] StyleRPrChildOrder =
    {
        "moveFrom", "moveTo", "ins", "del", "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps",
        "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
        "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz",
        "wShadow", "wTextOutline", "wTextFill", "wScene3d", "wProps3d", "szCs", "highlight", "u",
        "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
        "specVanish", "oMath", "rPrChange",
    };

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
        state.LeftStyleIds = ReadStyleIds(left);

        // Word-parity input-revision preservation (PreserveInputRevisions): map each accepted working-copy
        // body block back to its ORIGINAL source element. A revision-free LEFT keeps the established path:
        // right Equal/Insert emissions carry the RIGHT input's foreign markup verbatim. When LEFT itself is
        // dirty, do NOT preserve the RIGHT half alone (that is asymmetric); instead retain a narrowly safe
        // LEFT map for delete-side projection. A pre-existing left w:ins that the comparison deletes must be
        // emitted as deletion-grade markup, not flattened then re-deleted as a fresh unrelated change. More
        // complex two-sided cases (left moves/property revisions and Modify spans) deliberately stay on the
        // accepted-view renderer until they have source-provenance-aware handling.
        if (settings.PreserveInputRevisions)
        {
            if (HasTrackedRevisionMarkup(left))
            {
                state.LeftPreservedOriginals = BuildPreservedOriginalIndex(irLeft, left);
                // A raw LEFT deletion can be the historical counterpart of a right-only accepted block.
                // Do not try to infer this generally: only direct body ordinal matches with a fully deleted
                // source block are safe enough to project as the source author's insertion.
                state.LeftDeletedInsertionOriginals = BuildLeftDeletedInsertionIndex(irLeft, irRight, left);
            }
            else
                state.PreservedOriginals = BuildPreservedOriginalIndex(irRight, right);
        }

        // Assemble the new body's block-level children (w:p / w:tbl), in script order with Word's
        // replace-gap arrangement (inserted blocks before deleted ones inside each gap).
        var bodyBlocks = new List<XElement>();
        RenderBlockOpsWordShaped(script.Operations, state, bodyBlocks);

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

                // Capture this predicate while MAIN still is the untouched LEFT package.  Once
                // the body is assembled below, the output deliberately contains clones from both
                // sides and can no longer serve as evidence that a font list was shared by the
                // original inputs.
                var sharedCssFontStacks = DirectCssFontStacks(main);
                if (wDocRight.MainDocumentPart is { } rightMainForFontStacks)
                    sharedCssFontStacks.IntersectWith(DirectCssFontStacks(rightMainForFontStacks));
                else
                    sharedCssFontStacks.Clear();

                // A very narrow malformed-input compatibility shape: a complete body replacement can
                // introduce an entirely new paragraph-style universe into a left package that has no
                // defaults or default paragraph style of its own. Word projects the USED inserted styles
                // against the left package's stock defaults (rather than raw-copying them), which keeps
                // their line metrics compact. The helper proves all of the safety preconditions while the
                // package still contains the original LEFT stories; outside that shape it returns null and
                // style treatment remains the established general path.
                var leftHadTheme = main.ThemePart is not null;
                var insertedStyleNormalization = TryCreateInsertedStyleNormalization(
                    script, state, main, wDocRight.MainDocumentPart);
                var docDefaultsStyleProjection = TryCreateDocDefaultsStyleProjection(wDoc, wDocRight, state);

                // Preserve the trailing top-level sectPr (a direct child of w:body that is NOT inside a pPr).
                var trailingSectPr = bodyEl.Elements(W.sectPr).LastOrDefault();

                // Trailing-section property tracking (block-format-change family, Phase 3): when the left/right
                // trailing sectPr differ in their PROPERTIES (page size/margins/columns/…, ignoring header/footer
                // references + rsids), stamp native w:sectPrChange — the right properties are applied and the left
                // properties preserved in the marker. References (owned by the header/footer machinery, which runs
                // later and mutates them) and any mid-document sectPr inside a pPr are untouched (v1 ceilings).
                if (settings.TrackSectionFormatChanges && trailingSectPr != null)
                {
                    var rightTrailingSectPr = wDocRight.MainDocumentPart?.GetXDocument().Root?
                        .Element(W.body)?.Elements(W.sectPr).LastOrDefault();
                    if (rightTrailingSectPr != null && SectPrPropsDiffer(trailingSectPr, rightTrailingSectPr))
                        ApplySectPrChange(trailingSectPr, trailingSectPr, rightTrailingSectPr, state);
                }

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
                    RenderNoteScopes(script.NoteOps, state, main, wDocRight, settings, streamDoc, rightStream);

                // Header/footer story markup (2026-07-03 campaign): apply each changed story's edit ops
                // INSIDE its header/footer part (the output package carries the LEFT parts; a changed
                // story is rebuilt from its ops with native w:ins/w:del markup, so accept/reject
                // round-trips header/footer content too). Unchanged stories keep the verbatim carry-over.
                if (script.HeaderFooterOps is { Count: > 0 })
                    RenderHeaderFooterScopes(script.HeaderFooterOps, state, main,
                        wDocRight.MainDocumentPart, settings, streamDoc, rightStream);

                // Story-reference reconciliation — MUST run after the story machinery above so its
                // part mapping is known. Cloned RIGHT-side content (a paired paragraph's adopted pPr,
                // a seam, a whole-block insert) can carry inline w:sectPr story references under the
                // RIGHT package's r:ids, which in this left-based package are dangling or name a
                // relationship of the wrong kind — LibreOffice refuses such a package outright.
                // Each reference is REBOUND to the output part carrying that story (the left part the
                // story diff merged into for a matched pair, the freshly-inserted part for a
                // right-only story, or a wholesale import as a last resort); only references no
                // package can resolve are dropped (absent reference ⇒ OOXML section inheritance).
                RebindOrStripStoryReferences(state, main, wDocRight.MainDocumentPart);

                // Note-id renumber pass (M2.6 Task 1): mirror the oracle's ChangeFootnoteEndnoteReferencesToUniqueRange.
                // Walk the produced body in document order; every footnote/endnote reference gets a sequential id
                // (body-reference order, base 1), each note DEFINITION is renumbered + reordered to match, and the
                // reserved separator/continuation boilerplate notes keep their ids. Runs for EVERY render (cheap and
                // idempotent when ids already coincide) so accept-by-right-order / reject-by-left-order both hold.
                var footnoteRemap = RenumberNoteIds(main, W.footnoteReference, W.footnote, W.footnotes,
                    main.FootnotesPart, wDocRight.MainDocumentPart?.FootnotesPart);
                var endnoteRemap = RenumberNoteIds(main, W.endnoteReference, W.endnote, W.endnotes,
                    main.EndnotesPart, wDocRight.MainDocumentPart?.EndnotesPart);
                RemapNestedNoteReferences(main, footnoteRemap, endnoteRemap);

                // Comment fidelity passes (the comment analogue of NormalizeBookmarks). A commented paragraph now
                // renders FINELY (token-granular), carrying its range markers + reference run through the diff.
                // First merge any RIGHT-only comment definitions (+ commentsExtended/commentsIds threading) the
                // emitted right-sourced content references, so the LEFT-based comments part resolves every
                // reference. Then reconcile the body's markers to unique ids, 1:1 range pairing, and exactly-one
                // resolved definition per reference — collapsing an unchanged comment to a single bare range
                // (survives accept AND reject) and renumber-deduping a rewritten comment's del/ins copies.
                MergeRightCommentDefinitions(main, wDocRight.MainDocumentPart, streamDoc, rightStream);
                NormalizeComments(main, BodyCommentIds(state.Left), BodyCommentIds(state.RightSource), state);

                // Bookmark normalization pass: an edit straddling a bookmark range endpoint, or a dense
                // overlapping content-region layout, can leave the rendered body with a duplicate bookmark id
                // (schema-invalid Sem_UniqueAttributeValue), an unpaired marker, or a COMMON bookmark surviving
                // only one of accept/reject. Reconcile, identity-aware: a bookmark present in BOTH sources is
                // collapsed to a single BARE pair (survives accept AND reject); a genuinely inserted/deleted one
                // keeps its w:ins/w:del context; ids are made unique and every start↔end re-paired — so reject ≡
                // left / accept ≡ right at the bookmark-structure level and every REF/PAGEREF/NOTEREF/HYPERLINK\l
                // + internal hyperlink anchor still resolves.
                NormalizeBookmarks(main, BodyBookmarkNames(state.Left), BodyBookmarkNames(state.RightSource));

                // Field-context normalization: field plumbing (w:fldChar/w:instrText) is kept across edit
                // boundaries (AlwaysKeep) so a REF/PAGEREF field is never orphaned, but a boundary may leave the
                // plumbing in the wrong revision wrapper (e.g. a begin/separate run wrapped in w:ins after the
                // text before the field was edited — the field would then vanish on reject). Re-home each field's
                // plumbing to the field's own context in every rendered story (body, headers/footers, and notes):
                // bare for an unchanged or result-edited field (survives accept AND reject), left in w:del/w:ins
                // for a wholly deleted/inserted field.
                NormalizeFields(main);

                // Style-definition provenance (decoded from the Word-compare oracle corpus): the result
                // keeps the LEFT document's styles part — docDefaults/theme/latentStyles byte-identical
                // to the left — while each style whose RAW definition formatting differs between the
                // sides has its CURRENT payload updated to the RIGHT document's EFFECTIVE formatting,
                // with the left's effective payload archived in a tracked rPrChange/pPrChange INSIDE
                // the style definition. Eligible raw-equal styles also receive a reversible projection
                // when a literal docDefaults delta is the only safe presentation difference. Right-only
                // styles are copied in. Numbering keeps the existing missing-copy treatment (numId
                // collisions are remapped there, not overwritten).
                var rightImportedStyles = TrackStyleDefinitionChanges(
                    wDoc, wDocRight, state, insertedStyleNormalization, leftHadTheme, docDefaultsStyleProjection);
                // Equal blocks and archived pPr values are cloned verbatim, so neither path passes
                // through the paired-paragraph style guard. Once the final styles part is known,
                // discard only paragraph-style references which cannot resolve there.
                DropDanglingParagraphStyleRefs(main);
                // The output's surviving body is RIGHT-sourced (equal/inserted/modified blocks emit
                // the right document's XML) while the numbering part is seeded from the LEFT. When
                // a numId collides across the sides with different content, the copy renumbers the
                // right's definition to a fresh id — rebind every surviving reference to it, or
                // right-sourced lists silently resolve against the left's definition (decimal
                // rendering as the left's bullets). Deleted (left-sourced) paragraphs — the ones
                // whose paragraph mark carries w:del — keep the left id untouched, as do archived
                // left properties inside *Change elements.
                var numIdMap = WmlComparer.CopyMissingNumberingFromOneDocToAnother(wDocRight, wDoc);
                RebindRightNumberingReferences(main, numIdMap, state);
                RebindRightImportedStyleNumberingReferences(
                    main.StyleDefinitionsPart, rightImportedStyles, numIdMap);
                // Word-parity repair: a body numPr referencing a numId with NO definition (tool-made
                // corpus inputs ship this) renders as a plain paragraph in LibreOffice, while Word
                // synthesizes a decimal multilevel definition on open — its compare oracle carries
                // exactly that. Mirror the repair so numbered lists survive into the redline.
                RepairDanglingNumberingReferences(main);
                // Word always writes a settings part; a package without one makes LibreOffice fall
                // back to its own default tab stop (≈709 twips) instead of Word's 720, drifting every
                // tab-positioned run cumulatively. Backfill a minimal settings part with the 720-twip
                // default so tab metrics match the oracle (tool-generated corpus inputs ship without).
                if (main.DocumentSettingsPart is null)
                {
                    var settingsPart = main.AddNewPart<DocumentSettingsPart>("rIdSettingsBackfill");
                    settingsPart.GetXDocument().Add(new XElement(W.settings,
                        new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName),
                        new XElement(W.defaultTabStop, new XAttribute(W.val, 720))));
                    settingsPart.PutXDocument();
                }
                // Theme backfill: without a theme part, LibreOffice resolves scheme colors
                // (bg1/tx1/accentN) to BLACK — right-sourced charts and shapes render as black
                // boxes. Word's compare output always carries a theme, and when the ORIGINAL (left)
                // document has none Word supplies its STOCK default — NOT the revised document's
                // theme (adopting the right's shifts every theme-referencing cloned run's font away
                // from the oracle). Backfill Word's stock theme byte-for-byte (WordStockTheme:
                // Aptos fonts, 2023+ palette — verified identical across 164 Word outputs).
                if (!leftHadTheme)
                    BackfillDefaultTheme(main);
                // docDefaults backfill (same provenance rule as the theme): when the left's styles
                // part carries no w:docDefaults, Word's compare output backfills Word's STOCK
                // docDefaults — never the revised document's. Which stock depends on whether the
                // left had a theme: themeless lefts are seeded like a new document (modern stock:
                // sz 24, spacing after=160 line=278, kern, ligatures — the era of the stock theme
                // above), while a left with its own theme gets the classic-era stock (sz 22,
                // line=259). Verified across the six corpus oracles whose lefts lack docDefaults
                // (two byte-identical groups keyed exactly on theme presence).
                BackfillStockDocDefaults(main, leftHadTheme);

                // Some HTML-to-DOCX producers write CSS font-family lists directly into w:rFonts
                // (for example "Roboto, sans-serif").  Word's comparison output keeps that raw
                // string, but LibreOffice does not apply Word's fallback behavior when rendering a
                // tracked document.  A tightly-scoped compatibility projection is warranted only
                // when BOTH original documents carry the identical CSS-shaped list in direct run
                // formatting AND it is not archived by a tracked run-format change: together,
                // those prove it is shared formatting rather than one side's actual edit.  The
                // recognized syntax is either an unquoted comma-bearing list or one quoted primary
                // followed only by the CSS generic <c>sans-serif</c>; quoted lists with a concrete
                // fallback remain untouched because that fallback is semantically meaningful.  As
                // before, styles, themes, east-Asian fonts, and one-sided lists are excluded.
                ProjectSharedCssFontStacks(main, sharedCssFontStacks);
            }
            return streamDoc.GetModifiedWmlDocument();
        }
    }

    /// <summary>
    /// Projects the narrowly malformed, shared CSS-like direct font-list shapes described at the
    /// call site to LibreOffice's reliable sans-serif fallback.  This runs after all renderer
    /// package work has completed so it cannot influence edit alignment, source provenance, or
    /// style/numbering reconciliation.
    /// </summary>
    private static void ProjectSharedCssFontStacks(
        MainDocumentPart outputMain,
        HashSet<string> sharedStacks)
    {
        if (sharedStacks.Count == 0)
            return;

        var output = outputMain.GetXDocument();
        // A run-format revision's current rPr becomes the accepted document.  Projecting its font
        // would therefore leak the renderer fallback into that accepted result (and empirically
        // breaks the accepted-output oracle). Exclude only the exact stack recorded by such a
        // revision; unrelated format revisions remain harmless.
        sharedStacks.RemoveWhere(stack => output.Descendants(W.rPrChange)
            .Descendants(W.rFonts)
            .Any(fonts => IsExactFontTriplet(fonts, stack)));
        if (sharedStacks.Count == 0)
            return;

        var changed = false;
        foreach (var fonts in output.Descendants(W.rFonts))
        {
            // Limit the projection to direct run properties.  A style/default/theme carrying a
            // similarly-shaped value is semantically broader and has not earned this workaround.
            if (fonts.Parent?.Name != W.rPr || fonts.Parent.Parent?.Name != W.r)
                continue;

            var ascii = (string?)fonts.Attribute(W.ascii);
            if (ascii is null || !sharedStacks.Contains(ascii) || !IsExactFontTriplet(fonts, ascii))
                continue;

            fonts.SetAttributeValue(W.ascii, "Arial");
            fonts.SetAttributeValue(W.hAnsi, "Arial");
            fonts.SetAttributeValue(W.cs, "Arial");
            changed = true;
        }

        if (changed)
            outputMain.PutXDocument();
    }

    /// <summary>Returns exact direct-run font triplets that use one of the narrowly supported
    /// CSS-like font-list syntaxes. Exact matching is deliberate: different fallback ordering,
    /// a concrete fallback after a quoted primary, or a one-sided occurrence must not opt a
    /// document into the renderer compatibility projection.</summary>
    private static HashSet<string> DirectCssFontStacks(MainDocumentPart main)
    {
        var stacks = new HashSet<string>(StringComparer.Ordinal);
        foreach (var run in main.GetXDocument().Descendants(W.r))
        {
            var fonts = run.Element(W.rPr)?.Element(W.rFonts);
            var ascii = (string?)fonts?.Attribute(W.ascii);
            if (ascii is not null && fonts is not null && IsExactFontTriplet(fonts, ascii) &&
                IsCssFontStackWithArialFallback(ascii))
                stacks.Add(ascii);
        }
        return stacks;
    }

    private static bool IsExactFontTriplet(XElement fonts, string face)
        => (string?)fonts.Attribute(W.ascii) == face &&
           (string?)fonts.Attribute(W.hAnsi) == face &&
           (string?)fonts.Attribute(W.cs) == face;

    private static bool IsUnquotedCssFontStack(string value)
    {
        var first = value.TrimStart();
        return first.Length > 1 && first.IndexOf(',') > 0 && first[0] is not '\'' and not '\"';
    }

    /// <summary>
    /// Recognizes the one quoted CSS-family shape that has no concrete fallback to preserve:
    /// <c>"Primary", sans-serif</c> (or its single-quoted equivalent).  It deliberately declines
    /// multiple quoted names, explicit secondary faces such as <c>"Calibri", Arial, sans-serif</c>,
    /// malformed quoting, and non-sans generic families.  Those are not interchangeable with the
    /// renderer's Arial/Liberation Sans compatibility fallback.
    /// </summary>
    private static bool IsQuotedPrimaryWithOnlyGenericSansFallback(string value)
    {
        var trimmed = value.Trim();
        if (trimmed.Length < 5 || trimmed[0] is not '\'' and not '\"')
            return false;

        var quote = trimmed[0];
        int closingQuote = trimmed.IndexOf(quote, 1);
        if (closingQuote <= 1)
            return false;

        var remainder = trimmed[(closingQuote + 1)..].TrimStart();
        if (!remainder.StartsWith(",", StringComparison.Ordinal))
            return false;

        return string.Equals(remainder[1..].Trim(), "sans-serif", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsCssFontStackWithArialFallback(string value)
        => IsUnquotedCssFontStack(value) || IsQuotedPrimaryWithOnlyGenericSansFallback(value);

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

    /// <summary>
    /// Render a sequence of block ops with Microsoft Word's replace-gap arrangement (the grammar Word's
    /// own compare output uses at every site where old blocks are deleted and new blocks inserted):
    /// <list type="number">
    /// <item>INSERTED blocks render before the deleted ones (Word emits new content first, struck old
    /// content after it) — deletes are buffered until the run of pure Delete/Insert ops ends.</item>
    /// <item>The SEAM: the last inserted paragraph and the first deleted paragraph share one
    /// <c>w:p</c> — Word renders the old text inline right after the new text. The seam paragraph keeps
    /// the deleted paragraph's <c>pPr</c> (whose tracked paragraph mark makes reject restore the old
    /// paragraph exactly); the inserted paragraph's own mark-ins is dropped (the new side contributes
    /// n−1 inserted marks, exactly as Word does).</item>
    /// <item>The TERMINATOR: the last deleted paragraph of a seam-merged gap keeps a LIVE (untracked)
    /// mark, so accepting ends the inserted text at it instead of bleeding into the following block,
    /// and rejecting still restores that paragraph under its own mark.</item>
    /// </list>
    /// Every other op kind (Equal/FormatOnly/Modify/Move/Split/Merge) flushes the buffer and renders in
    /// script order, so single-sided gaps and non-gap ops render exactly as before. The seam mutates the
    /// last inserted paragraph IN PLACE (it may be registered as a right-sourced clone for the
    /// post-assembly media/relationship import — the registered instance must stay the live tree node);
    /// deleted blocks are left-sourced and safe to consume. The accept ≡ right / reject ≡ left contract
    /// is preserved in both directions (proof: DocxDiffWordShapeTests).
    /// </summary>
    internal static void RenderBlockOpsWordShaped(
        IEnumerable<IrEditOp> ops, RenderState state, List<XElement> sink)
    {
        // A MoveModify destination deliberately owns only its RIGHT anchor: its paired LEFT anchor lives on
        // the separately emitted source op.  Resolve that pairing for THIS operation list before rendering.
        // Scoping matters because note/header/cell projections allocate move-group ids locally; a nested cell
        // render must not replace the body scope's source lookup once it returns.
        var opList = ops as IReadOnlyList<IrEditOp> ?? ops.ToList();
        var previousMoveSources = state.ActiveMoveSourceAnchors;
        var moveSources = new Dictionary<int, string>();
        foreach (var candidate in opList)
            if (candidate.IsMoveSource == true && candidate.MoveGroupId is { } groupId &&
                candidate.LeftAnchor is { } leftAnchor)
                moveSources[groupId] = leftAnchor;
        state.ActiveMoveSourceAnchors = moveSources;

        try
        {
        var pendingDeletes = new List<IrEditOp>();
        int gapInsertStart = -1;   // sink index where the current gap's inserted elements begin
        int? lastInsertedBodyFullRewriteGroupId = null;

        void FlushGap()
        {
            if (pendingDeletes.Count == 0)
            {
                gapInsertStart = -1;
                lastInsertedBodyFullRewriteGroupId = null;
                return;
            }
            int? firstDeletedBodyFullRewriteGroupId = pendingDeletes[0].BodyFullRewriteGroupId;
            var delEls = new List<XElement>();
            foreach (var d in pendingDeletes)
                RenderBlockOp(d, state, delEls);
            pendingDeletes.Clear();

            var hasInserts = gapInsertStart >= 0 && sink.Count > gapInsertStart;
            var lastIns = hasInserts ? sink[^1] : null;
            var firstDel = delEls.Count > 0 ? delEls[0] : null;
            // Seam-merge guard: both boundary blocks must be plain paragraphs; neither paragraph's
            // pPr may carry an inline w:sectPr (swapping the pPr would silently move a section
            // break); and the deleted paragraph must not carry a PAGE BREAK (pageBreakBefore or a
            // w:br type="page" run) — Word keeps a deleted page break PAGINATING, so the deleted
            // paragraph stays standalone and the following struck content still starts its own page.
            static bool CarriesPageBreak(XElement p) =>
                p.Element(W.pPr)?.Element(W.pageBreakBefore) is not null ||
                p.Descendants(W.br).Any(b => (string?)b.Attribute(W.type) == "page");
            // This is intentionally explicit alignment provenance, never a renderer guess based on
            // a 1×1 cardinality or textual heuristic. The builder can set it only for a body-level,
            // non-tail full lexical rewrite, where Word Compare keeps two physical marked paragraphs.
            // Cell/textbox/note/header/footer ops remain unmarked and retain the normal seam.
            bool suppressSeam = lastInsertedBodyFullRewriteGroupId is { } groupId &&
                firstDeletedBodyFullRewriteGroupId == groupId;
            if (lastIns is not null && firstDel is not null &&
                !suppressSeam &&
                lastIns.Name == W.p && firstDel.Name == W.p &&
                lastIns.Element(W.pPr)?.Element(W.sectPr) is null &&
                firstDel.Element(W.pPr)?.Element(W.sectPr) is null &&
                !CarriesPageBreak(firstDel) && !CarriesPageBreak(lastIns))
            {
                // Capture the ins-side pPr before it is dropped: Word's seam-terminator carries the
                // RIGHT side's DIRECT paragraph props (real spacing/indents survive) but never a
                // style reference the left universe can't resolve, and never format-change bars.
                XElement? insPPr = null;
                if (lastIns.Element(W.pPr) is { } insSrc)
                {
                    insPPr = StripUnids(new XElement(insSrc));
                    insPPr.Elements(W.rPr).Elements(W.ins).Remove();
                    insPPr.Elements(W.rPr).Where(r => !r.HasElements && !r.HasAttributes).Remove();
                    DropUnresolvableStyleRef(insPPr, state);
                    if (!insPPr.HasElements && !insPPr.HasAttributes)
                        insPPr = null;
                }

                // Mutate lastIns in place into the seam: drop its pPr (and with it the mark-ins),
                // adopt the deleted paragraph's pPr (carrying the tracked mark + old paragraph props),
                // and move the deleted paragraph's content in AFTER the inserted runs.
                lastIns.Element(W.pPr)?.Remove();
                if (firstDel.Element(W.pPr) is { } delPPr)
                {
                    delPPr.Remove();
                    lastIns.AddFirst(delPPr);
                }
                foreach (var child in firstDel.Elements().ToList())
                {
                    child.Remove();
                    lastIns.Add(child);
                }
                delEls.RemoveAt(0);

                // Live terminator: the last deleted paragraph of the CONTIGUOUS paragraph chain that
                // starts at the seam (the seam itself when no deleted paragraph follows it directly)
                // keeps an untracked mark — accept coalesces the chain into it, ending the inserted
                // text there. A deleted TABLE breaks the chain: paragraphs beyond it are disconnected
                // from the seam's accept-time coalescing, so they keep their tracked marks (accept
                // removes them via the deleted-range rules; a live mark there would leave a stray
                // empty paragraph behind).
                var terminator = lastIns;
                foreach (var el in delEls)
                {
                    if (el.Name != W.p)
                        break;
                    terminator = el;
                }
                terminator.Element(W.pPr)?.Element(W.rPr)?.Elements(W.del).Remove();

                // Word's seam-terminator shape (decoded across the oracle corpus): the surviving
                // paragraph carries the INS side's DIRECT props — real spacing/indent survive —
                // but never a style reference the left universe can't resolve (a Title/Heading on
                // the right renders plain in Word's own compare output when the left lacks the
                // style), never the del side's pPr (the left's style must not outlive accept),
                // and never a w:pPrChange (Word puts no format-change bars on seam lines).
                // Skipped when the del-side pPr carries an inline w:sectPr — replacing it would
                // silently move a section break (the seam guard above only vets firstDel/lastIns,
                // not a chain terminator).
                if (terminator.Element(W.pPr) is not { } termPPr || termPPr.Element(W.sectPr) is null)
                {
                    terminator.Element(W.pPr)?.Remove();
                    if (insPPr is not null)
                        terminator.AddFirst(new XElement(insPPr));
                }
            }
            sink.AddRange(delEls);
            gapInsertStart = -1;
            lastInsertedBodyFullRewriteGroupId = null;
        }

        foreach (var op in opList)
        {
            if (op.Kind == IrEditOpKind.DeleteBlock || op.Kind == IrEditOpKind.InsertBlock)
            {
                if (gapInsertStart < 0)
                {
                    gapInsertStart = sink.Count;
                    lastInsertedBodyFullRewriteGroupId = null;
                }
                if (op.Kind == IrEditOpKind.DeleteBlock)
                    pendingDeletes.Add(op);       // buffered: renders after the gap's inserts
                else
                {
                    RenderBlockOp(op, state, sink); // insert leapfrogs the buffered deletes of its gap
                    lastInsertedBodyFullRewriteGroupId = op.BodyFullRewriteGroupId;
                }
                continue;
            }
            FlushGap();
            RenderBlockOp(op, state, sink);
        }
        FlushGap();
        }
        finally
        {
            state.ActiveMoveSourceAnchors = previousMoveSources;
        }
    }

    internal static void RenderBlockOp(IrEditOp op, RenderState state, List<XElement> sink)
    {
        // A block-level content control owns a non-run OOXML envelope (`w:sdtPr`, `w:sdtEndPr`, nesting,
        // bindings, locks, etc.). It cannot be reconstructed safely from independent paragraph/table ops.
        // Route every operation through the atomic envelope renderer before the generic block dispatch.
        if (IsBlockSdtOp(op, state))
        {
            RenderBlockSdtOp(op, state, sink);
            return;
        }

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
                // Text-equal, block-format-differing. A paragraph stamps per-run w:rPrChange + w:pPrChange;
                // a TABLE stamps the native table-shell property markers (block-format-change family); any
                // other non-paragraph pair falls through to a verbatim right emit.
                if (ResolveBlock(op.RightAnchor, state.RightSource) is IrTable)
                    EmitFormatOnlyTable(op, state, sink);
                else
                    EmitFormatOnlyParagraph(op, state, sink);
                break;

            case IrEditOpKind.InsertBlock:
                // A very narrow dirty-left projection is supported here only. Its raw-left candidate is
                // keyed to this exact main-body InsertBlock target; other right-side insert emissions (a
                // replacement half, move destination, split member, note/header block, etc.) must remain on
                // the normal accepted-view path.
                EmitWholeBlock(op.RightAnchor, state.RightSource, state, sink, RevKind.Ins, fromRight: true,
                    projectLeftDeletionAsInsertion: true);
                break;

            case IrEditOpKind.DeleteBlock:
                EmitWholeBlock(op.LeftAnchor, state.Left, state, sink, RevKind.Del, fromRight: false);
                break;

            case IrEditOpKind.ModifyBlock:
                RenderModifyBlock(op, state, sink);
                break;

            case IrEditOpKind.MoveBlock:
            case IrEditOpKind.MoveModifyBlock:
                // An inline envelope or non-hyperlink field carrier change is structurally inseparable from its
                // paragraph. Do not disguise it as a native move whose destination can slice the carrier: a full
                // delete/insert pair is the only representation that makes both Accept and Reject exact.
                if (op.RequiresWholeParagraphReplace)
                {
                    if (op.IsMoveSource == true)
                        EmitWholeBlock(op.LeftAnchor, state.Left, state, sink, RevKind.Del, fromRight: false);
                    else
                        EmitWholeBlock(op.RightAnchor, state.RightSource, state, sink, RevKind.Ins, fromRight: true);
                    break;
                }
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

    /// <summary>Whether either side of an operation resolves to an atomic block-level content control.</summary>
    private static bool IsBlockSdtOp(IrEditOp op, RenderState state) =>
        (op.LeftAnchor is { } left && ResolveBlock(left, state.Left) is IrSdtBlock) ||
        (op.RightAnchor is { } right && ResolveBlock(right, state.RightSource) is IrSdtBlock);

    /// <summary>
    /// Render a block-level <c>w:sdt</c> as a single ownership unit.  Native content-control range revisions
    /// toggle the wrapper itself, while normal whole-block run/paragraph/table marking toggles its payload.
    /// Both layers are required: Word's custom-XML deletion range accepts by COLLAPSING an SDT to its content,
    /// not by deleting that content, so a range-only representation would leak a deleted control's text.
    /// </summary>
    private static void RenderBlockSdtOp(IrEditOp op, RenderState state, List<XElement> sink)
    {
        switch (op.Kind)
        {
            case IrEditOpKind.EqualBlock:
            case IrEditOpKind.FormatOnlyBlock:
                EmitVerbatim(op.RightAnchor, state.RightSource, state, sink, fromRight: true);
                return;

            case IrEditOpKind.InsertBlock:
                EmitWholeSdt(op.RightAnchor, state.RightSource, state, sink, RevKind.Ins, fromRight: true);
                return;

            case IrEditOpKind.DeleteBlock:
                EmitWholeSdt(op.LeftAnchor, state.Left, state, sink, RevKind.Del, fromRight: false);
                return;

            case IrEditOpKind.ModifyBlock:
                EmitWholeSdt(op.LeftAnchor, state.Left, state, sink, RevKind.Del, fromRight: false);
                EmitWholeSdt(op.RightAnchor, state.RightSource, state, sink, RevKind.Ins, fromRight: true);
                return;

            case IrEditOpKind.MoveBlock:
            case IrEditOpKind.MoveModifyBlock:
                // The aligner deliberately lowers SDT relocations to delete+insert. Keep this defensive
                // projection too so an old/corrupt script can never emit a native move around an SDT envelope.
                if (op.IsMoveSource == true)
                    EmitWholeSdt(op.LeftAnchor, state.Left, state, sink, RevKind.Del, fromRight: false);
                else
                    EmitWholeSdt(op.RightAnchor, state.RightSource, state, sink, RevKind.Ins, fromRight: true);
                return;

            default:
                // Split/merge operations are paragraph-only by construction. A malformed SDT-bearing op
                // degrades to the same reversible old/new envelope pair as a generic structural replacement.
                if (op.LeftAnchor is not null)
                    EmitWholeSdt(op.LeftAnchor, state.Left, state, sink, RevKind.Del, fromRight: false);
                if (op.RightAnchor is not null)
                    EmitWholeSdt(op.RightAnchor, state.RightSource, state, sink, RevKind.Ins, fromRight: true);
                return;
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

        if (!op.RequiresWholeParagraphReplace && op.TokenDiff is { } tokenDiff && leftIsPara && rightIsPara &&
            op.TextboxDiffs is null)             // textbox-interior diffs are not finely rendered in Task 3
        {
            // Commented paragraphs render finely too: comment range markers + the commentReference run ride
            // through the token diff as AlwaysKeep zero-width markers (IsAlwaysKeepMarker / WalkRun), then
            // NormalizeComments reconciles them to unique/paired/resolved markup. (Was: bailed to whole-block.)
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

        // A right-only cell is reversible in place only when the renderer can also emit tblGridChange /
        // tcPrChange histories. With table-format tracking disabled, cloning the accepted (right) grid and
        // merely marking the cell w:cellIns would make Reject remove the cell but retain the widened grid and
        // widths. Bail before touching render state so the caller emits the conservative whole-table pair.
        if (!state.Settings.TrackTableFormatChanges && tableDiff.RowOps.Any(row =>
                row.Kind == IrRowOpKind.ModifyRow && row.CellOps is { } cells &&
                cells.Any(cell => cell.LeftCellAnchor == null || cell.RightCellAnchor == null)))
            return false;

        // Index the source rows by anchor so a row op resolves to its source w:tr.
        var leftRowsByAnchor = IndexRows(ResolveBlock(op.LeftAnchor, state.Left) as IrTable);
        var rightRowsByAnchor = IndexRows(ResolveBlock(op.RightAnchor, state.RightSource) as IrTable);

        var newTbl = new XElement(W.tbl);
        // Carry the table's non-row prelude (tblPr, tblGrid, …) from the right shell.
        foreach (var pre in rightTbl.Elements().Where(e => e.Name != W.tr))
            newTbl.Add(StripUnids(new XElement(pre)));
        // Table-level shell changes (block-format-change family): tblPr / tblGrid.
        ApplyTableLevelShellChanges(newTbl, leftTbl, state);

        foreach (var rowOp in tableDiff.RowOps)
        {
            switch (rowOp.Kind)
            {
                case IrRowOpKind.EqualRow:
                {
                    if (!rightRowsByAnchor.TryGetValue(rowOp.RightRowAnchor ?? "", out var src)) return false;
                    var row = StripUnids(new XElement(src));
                    state.RegisterMediaReferences(row);
                    // An EqualRow (content-equal) may still carry a trPr/tcPr shell change — track it.
                    if (rowOp.LeftRowAnchor is { } ela && leftRowsByAnchor.TryGetValue(ela, out var eleft))
                        ApplyRowAndCellShellChanges(row, eleft, state);
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
                    XElement? leftRowSrc = rowOp.LeftRowAnchor is { } mla && leftRowsByAnchor.TryGetValue(mla, out var ml) ? ml : null;
                    if (!RenderModifyRow(rowOp, rightSrc, leftRowSrc, state, newTbl))
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
    /// content per its block ops, and whole-marking an unpaired (column-surplus) cell. When
    /// <paramref name="leftRowSrc"/> is supplied, the row's <c>w:trPr</c> and each paired cell's <c>w:tcPr</c>
    /// shell change is tracked (block-format-change family). Returns false to bail to the caller's fallback if
    /// the structure can't be resolved.</summary>
    private static bool RenderModifyRow(IrRowOp rowOp, XElement rightRowSrc, XElement? leftRowSrc, RenderState state, XElement newTbl)
    {
        // Without a per-cell op list, emit the right row as-is (content-equal row that fell into a ModifyRow by
        // row-property change only) — still round-trips.
        if (rowOp.CellOps == null)
        {
            var row0 = StripUnids(new XElement(rightRowSrc));
            state.RegisterMediaReferences(row0);
            if (leftRowSrc != null)
                ApplyRowAndCellShellChanges(row0, leftRowSrc, state);
            newTbl.Add(row0);
            return true;
        }

        var rightCells = rightRowSrc.Elements(W.tc).ToList();
        var leftCells = leftRowSrc?.Elements(W.tc).ToList();
        // A monotone cell-op sequence is renderable whenever every output (right) cell is represented once
        // and no left-only deletion is present.  This includes an ordinary-grid insertion at the head or in
        // the middle: build from the accepted right grid, mark only the right-only cell w:cellIns, and let
        // tblGridChange restore the old grid on reject.  Left-only cells remain a conservative whole-table
        // fallback until the delete/merge topology path is made grid-aware.
        if (rowOp.CellOps.Count(c => c.RightCellAnchor != null) != rightCells.Count ||
            rowOp.CellOps.Any(c => c.RightCellAnchor == null) ||
            (leftCells != null && rowOp.CellOps.Count(c => c.LeftCellAnchor != null) != leftCells.Count))
            return false;

        var newRow = new XElement(W.tr);
        foreach (var pre in rightRowSrc.Elements().Where(e => e.Name != W.tc))
            newRow.Add(StripUnids(new XElement(pre)));

        int rightIndex = 0;
        int leftIndex = 0;
        foreach (var cellOp in rowOp.CellOps)
        {
            if (rightIndex >= rightCells.Count)
                return false;
            var cellSrc = rightCells[rightIndex++];
            if (cellOp.LeftCellAnchor == null)
            {
                var insertedCell = StripUnids(new XElement(cellSrc));
                state.RegisterMediaReferences(insertedCell);
                MarkWholeCell(insertedCell, RevKind.Ins, state);
                newRow.Add(insertedCell);
                continue;
            }

            XElement? leftCellSrc = null;
            if (leftCells != null)
            {
                if (leftIndex >= leftCells.Count)
                    return false;
                leftCellSrc = leftCells[leftIndex++];
            }

            var newCell = new XElement(W.tc);
            foreach (var pre in cellSrc.Elements().Where(e => e.Name != W.p && e.Name != W.tbl && e.Name != W.sdt))
                newCell.Add(StripUnids(new XElement(pre)));

            if (cellOp.BlockOps != null)
            {
                // Render the cell's block ops with the same dispatch the body uses (paragraph token diffs, etc.).
                var cellSink = new List<XElement>();
                RenderBlockOpsWordShaped(cellOp.BlockOps, state, cellSink);
                // A cell must contain at least one block-level child; if the ops produced none, keep the right
                // cell's content verbatim so the table stays schema-valid.
                if (cellSink.Count == 0)
                    foreach (var b in cellSrc.Elements().Where(e => e.Name == W.p || e.Name == W.tbl || e.Name == W.sdt))
                        cellSink.Add(StripUnids(new XElement(b)));
                newCell.Add(cellSink);
            }
            else
            {
                foreach (var b in cellSrc.Elements().Where(e => e.Name == W.p || e.Name == W.tbl || e.Name == W.sdt))
                    newCell.Add(StripUnids(new XElement(b)));
            }
            // Do not compare by output ordinal: a middle cellIns shifts every later right cell.  The
            // monotone differ guarantees leftCellSrc is exactly this paired operation's source cell.
            if (leftCellSrc != null)
                ApplyPairedCellShellChange(newCell, leftCellSrc, state);
            newRow.Add(newCell);
        }
        state.RegisterMediaReferences(newRow);
        // The paired tcPr histories were applied per cell above; row shells remain one per row.
        if (leftRowSrc != null)
            ApplyRowShellChanges(newRow, leftRowSrc, state);
        newTbl.Add(newRow);
        return true;
    }

    /// <summary>Emit a composed INSERTED cell (column add): clone the inserting reviewer's whole <c>w:tc</c>
    /// under that reviewer's state (author attribution + media bucket), mark it <c>w:tcPr/w:cellIns</c> with
    /// ins-marked content, and append it to <paramref name="newRow"/>. Unresolvable sources emit nothing
    /// (defensive; the merger only records resolvable reviewer cells).</summary>
    private static void EmitComposedInsertedCell(
        IrAuthoredCellOp cellOp, IReadOnlyList<IrDocument> reviewerIrs, RenderState state, XElement newRow)
    {
        if (cellOp.ShellSourceReviewer < 0 || cellOp.ShellSourceReviewer >= reviewerIrs.Count
            || cellOp.ShellRightCellAnchor is not { } anchor
            || FindCellSource(reviewerIrs[cellOp.ShellSourceReviewer], anchor) is not { } src)
            return;

        var newCell = StripUnids(new XElement(src));
        var savedAuthor = state.AuthorOverride;
        var savedSource = state.RightSource;
        var savedId = state.RightSourceId;
        state.AuthorOverride = cellOp.Author;
        state.RightSource = reviewerIrs[cellOp.ShellSourceReviewer];
        state.RightSourceId = cellOp.ShellSourceReviewer;
        state.RegisterMediaReferences(newCell);
        MarkWholeCell(newCell, RevKind.Ins, state);
        state.AuthorOverride = savedAuthor;
        state.RightSource = savedSource;
        state.RightSourceId = savedId;
        newRow.Add(newCell);
    }

    /// <summary>Mark a whole table cell inserted/deleted with Word's native cell-revision marks: a
    /// <c>w:tcPr/w:cellIns</c>|<c>w:cellDel</c> marker (appended in tcPr — the cell-revision marks sit at the
    /// end of the property order) plus every paragraph in the cell run-and-mark wrapped. Accept then removes a
    /// <c>cellDel</c> cell / keeps a <c>cellIns</c> cell bare; reject restores / removes it —
    /// <see cref="RevisionProcessor"/> implements both sides.</summary>
    private static void MarkWholeCell(XElement tc, RevKind kind, RenderState state)
    {
        var tcPr = tc.Element(W.tcPr);
        if (tcPr == null)
        {
            tcPr = new XElement(W.tcPr);
            tc.AddFirst(tcPr);
        }
        tcPr.Elements().Where(e => e.Name == W.cellIns || e.Name == W.cellDel).Remove();
        tcPr.Add(new XElement(kind == RevKind.Ins ? W.cellIns : W.cellDel, state.RevisionAttributes()));
        foreach (var p in tc.Descendants(W.p).ToList())
            MarkWholeParagraph(p, kind, state);
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

        // B2: the table-level shells (tblPr/tblGrid) were base-cloned above; swap in the composed winner's and
        // stamp native w:tblPrChange/w:tblGridChange (inner = base) so a table-shell edit round-trips.
        ApplyComposedTableShell(newTbl, baseTbl, op.TableShell, reviewerIrs, state);

        foreach (var rowOp in authoredRows)
        {
            switch (rowOp.Kind)
            {
                case IrRowOpKind.EqualRow:
                {
                    if (rowOp.BaseRowAnchor is { } ra && baseRowsByAnchor.TryGetValue(ra, out var src))
                    {
                        // A content-Equal row may still carry a composed trPr/tblPrEx shell edit (B2): build it
                        // and stamp the row-level marker; otherwise it is the base row verbatim.
                        var newRow = StripUnids(new XElement(src));
                        ApplyComposedRowShell(newRow, src, rowOp, reviewerIrs, state);
                        newTbl.Add(newRow);
                    }
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

    /// <summary>The source <c>w:tc</c> a cell anchor resolves to in <paramref name="ir"/> (cells are not in
    /// AnchorIndex): walk every indexed table's rows/cells, recursing into nested tables. Null if unknown.</summary>
    private static XElement? FindCellSource(IrDocument ir, string cellAnchor)
    {
        foreach (var block in ir.AnchorIndex.Values)
            if (block is IrTable tbl && FindCellSourceInTable(tbl, cellAnchor) is { } found)
                return found;
        return null;
    }

    private static XElement? FindCellSourceInTable(IrTable tbl, string cellAnchor)
    {
        foreach (var row in tbl.Rows)
            foreach (var cell in row.Cells)
            {
                if (cell.Anchor.ToString() == cellAnchor)
                    return cell.Source.Element;
                foreach (var b in cell.Blocks)
                    if (b is IrTable nested && FindCellSourceInTable(nested, cellAnchor) is { } found)
                        return found;
            }
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

        // B2: swap in the composed winner's trPr/tblPrEx and stamp the row-level marker (inner = base), so a
        // row-shell edit that rides alongside cell edits round-trips.
        ApplyComposedRowShell(newRow, baseRowSrc, rowOp, reviewerIrs, state);

        if (rowOp.ComposedCells is not { } cells)
        {
            // No per-cell view: keep the base row verbatim (defensive).
            newTbl.Add(StripUnids(new XElement(baseRowSrc)));
            return;
        }

        foreach (var cellOp in cells)
        {
            // A reviewer-INSERTED cell (column add): the whole cell clones from the reviewer, marked
            // w:tcPr/w:cellIns + ins-marked content — accept keeps it, reject removes it.
            if (cellOp.Kind == IrAuthoredCellKind.InsertCell)
            {
                EmitComposedInsertedCell(cellOp, reviewerIrs, state, newRow);
                continue;
            }

            XElement? baseCellSrc = cellOp.BaseCellAnchor != null
                && baseCellsByAnchor.TryGetValue(cellOp.BaseCellAnchor, out var bc) ? bc : null;
            if (baseCellSrc == null)
                continue;

            // A reviewer-DELETED base cell (column remove): the base cell marked w:tcPr/w:cellDel +
            // del-marked content — accept removes it, reject restores it.
            if (cellOp.Kind == IrAuthoredCellKind.DeleteCell)
            {
                var deletedCell = StripUnids(new XElement(baseCellSrc));
                var savedAuthor0 = state.AuthorOverride;
                state.AuthorOverride = cellOp.Author;
                MarkWholeCell(deletedCell, RevKind.Del, state);
                state.AuthorOverride = savedAuthor0;
                newRow.Add(deletedCell);
                continue;
            }

            // The cell SHELL (tcPr etc.) is cloned from the base cell by default; when the merger attributed
            // the shell to a reviewer (ShellSourceReviewer/ShellRightCellAnchor — a changed cell, so a
            // width/merge-only edit composes instead of silently reverting to the base shell), clone that
            // reviewer's right-cell shell instead. Falls back to base if the reviewer cell is unresolvable.
            var shellSrc = baseCellSrc;
            if (cellOp.ShellSourceReviewer >= 0 && cellOp.ShellSourceReviewer < reviewerIrs.Count
                && cellOp.ShellRightCellAnchor is { } shellAnchor
                && FindCellSource(reviewerIrs[cellOp.ShellSourceReviewer], shellAnchor) is { } reviewerCellSrc)
            {
                shellSrc = reviewerCellSrc;
            }

            var newCell = new XElement(W.tc);
            foreach (var pre in shellSrc.Elements().Where(e => e.Name != W.p && e.Name != W.tbl && e.Name != W.sdt))
                newCell.Add(StripUnids(new XElement(pre)));

            // The winner's shell was swapped in above; stamp a native w:tcPrChange (inner = BASE tcPr)
            // attributed to the shell winner, so accept keeps the winner's shell and reject restores the base
            // shell BYTES (not just the text). No-op when the reviewer shell is canonically equal to base
            // (ShellDiffers short-circuits) or when the base shell was kept (shellSrc == baseCellSrc). B2.
            if (!ReferenceEquals(shellSrc, baseCellSrc) && state.Settings.TrackTableFormatChanges)
            {
                var savedShellAuthor = state.AuthorOverride;
                state.AuthorOverride = cellOp.ShellAuthor;
                ApplyShellChange(newCell, W.tcPr, W.tcPrChange, baseCellSrc.Element(W.tcPr), state,
                    idOnly: false, TcPrInnerExclude);
                state.AuthorOverride = savedShellAuthor;
            }

            if (cellOp.ComposedBlockOps is { } blockOps)
            {
                var cellSink = new List<XElement>();
                foreach (var cellBlock in blockOps)
                    renderOneCompositeBlock(cellBlock, baseIr, reviewerIrs, state, cellSink);
                if (cellSink.Count == 0)
                    foreach (var b in baseCellSrc.Elements().Where(e => e.Name == W.p || e.Name == W.tbl || e.Name == W.sdt))
                        cellSink.Add(StripUnids(new XElement(b)));
                newCell.Add(cellSink);
            }
            else
            {
                // Base passthrough: the base cell's content verbatim.
                foreach (var b in baseCellSrc.Elements().Where(e => e.Name == W.p || e.Name == W.tbl || e.Name == W.sdt))
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
        IReadOnlyList<IrNoteDiff> noteOps, RenderState state, MainDocumentPart main,
        WordprocessingDocument? wDocRight, IrDiffSettings settings,
        OpenXmlMemoryStreamDocument leftStreamDoc, OpenXmlMemoryStreamDocument rightStreamDoc)
    {
        // Group note diffs by their target part so each part is loaded/saved once.
        var footnoteDiffs = noteOps.Where(n => n.Kind == IrNoteKind.Footnote).ToList();
        var endnoteDiffs = noteOps.Where(n => n.Kind == IrNoteKind.Endnote).ToList();

        var rightMain = wDocRight?.MainDocumentPart;
        ApplyNoteDiffsToPart(footnoteDiffs, EnsureNotePart(main, isFootnote: true, rightMain),
            rightMain?.FootnotesPart, W.footnote, W.footnotes, state, settings, leftStreamDoc, rightStreamDoc);
        ApplyNoteDiffsToPart(endnoteDiffs, EnsureNotePart(main, isFootnote: false, rightMain),
            rightMain?.EndnotesPart, W.endnote, W.endnotes, state, settings, leftStreamDoc, rightStreamDoc);
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
        RenderState state, IrDiffSettings settings, OpenXmlMemoryStreamDocument leftStreamDoc,
        OpenXmlMemoryStreamDocument rightStreamDoc)
    {
        if (diffs.Count == 0 || part == null)
            return;
        var xDoc = part.GetXDocument();
        var root = xDoc.Root;
        if (root == null)
            return;
        var rightRoot = rightPart?.GetXDocument().Root;

        // The registry is shared by body, notes, and stories. Take a per-note-part watermark so the
        // relationships for clones rendered below are imported from THIS right note part, rather than the
        // main document part (OOXML relationship ids are part-scoped).
        int clonesBefore = state.RightSourcedClones.Count;
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
                foreach (var pre in rightNote.Elements().Where(e => e.Name != W.p && e.Name != W.tbl && e.Name != W.sdt))
                    noteEl.Add(StripUnids(new XElement(pre)));
                root.Add(noteEl);
            }

            // Render the note's block ops to a fresh block list (same dispatch as the body).
            var noteBlocks = new List<XElement>();
            RenderBlockOpsWordShaped(diff.Ops, state, noteBlocks);
            if (settings is { RenderMoves: true, SimplifyMoveMarkup: true })
                foreach (var b in noteBlocks)
                    SimplifyMoveMarkup(b);

            // Strip engine-internal pt bookkeeping from the rendered blocks.
            foreach (var b in noteBlocks)
                StripUnids(b);

            // Replace the note's block-level children (w:p / w:tbl / w:sdt), keeping any non-block prelude.
            noteEl.Elements().Where(e => e.Name == W.p || e.Name == W.tbl || e.Name == W.sdt).Remove();
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
            ImportNoteSourcedRelationships(state, clonesBefore, part, rightPart,
                leftStreamDoc, rightStreamDoc);

            // A note part should never carry pt bookkeeping in the output.
            foreach (var attr in root.DescendantsAndSelf().Attributes()
                         .Where(a => a.Name.Namespace == PtOpenXml.pt).ToList())
                attr.Remove();
            part.PutXDocument();
        }
    }

    /// <summary>
    /// Import media parts plus hyperlink/external relationships referenced by clones registered while one
    /// footnote/endnote part was rendered. The body import cannot serve these clones: relationship ids are
    /// scoped to the note part that owns them.
    /// </summary>
    private static void ImportNoteSourcedRelationships(
        RenderState state, int clonesBefore, OpenXmlPart outputPart, OpenXmlPart? rightPart,
        OpenXmlMemoryStreamDocument leftStreamDoc, OpenXmlMemoryStreamDocument rightStreamDoc)
    {
        if (rightPart is null)
            return;
        var noteClones = state.RightSourcedClones.Skip(clonesBefore).ToList();
        if (noteClones.Count == 0)
            return;

        ImportHyperlinkAndExternalRelationships(noteClones, outputPart, rightPart);

        var outPkgPart = leftStreamDoc.GetPackage().GetPart(outputPart.Uri);
        var rightPkgPart = rightStreamDoc.GetPackage().GetPart(rightPart.Uri);
        foreach (var clone in noteClones)
            WmlComparer.MoveRelatedPartsToDestination(
                rightPkgPart, outPkgPart, clone, skipDanglingRelationships: true,
                skipHeaderFooterReferences: true);
    }

    // ----------------------------------------------------------------- header/footer story markup (2026-07-03)

    /// <summary>
    /// Apply header/footer story edit ops inside the output package's header/footer parts. A MATCHED
    /// story (both part URIs set) locates the LEFT part by URI — the output is a left-package clone, so
    /// left URIs resolve directly — renders its ops through the shared <see cref="RenderBlockOp"/>
    /// dispatch (story anchors resolve in the shared AnchorIndex; revision ids draw from the same
    /// <see cref="RenderState"/> counter as the body, staying document-unique), and replaces the part
    /// root's block-level children. Right-sourced clones inside a story import their media/hyperlink
    /// relationships into THAT part (header/footer parts own their relationships — the main-part import
    /// cannot resolve them).
    /// </summary>
    private static void RenderHeaderFooterScopes(
        IReadOnlyList<IrHeaderFooterDiff> hfOps, RenderState state, MainDocumentPart main,
        MainDocumentPart? rightMain, IrDiffSettings settings,
        OpenXmlMemoryStreamDocument leftStreamDoc, OpenXmlMemoryStreamDocument rightStreamDoc)
    {
        // A left story can feed several independently revised output stories. Materialize every clone BEFORE
        // rendering any primary story: otherwise a primary A→B rewrite would become the accidental source for a
        // later A→C clone and rejecting that clone would restore B rather than the real left A.
        var outputParts = new Dictionary<int, OpenXmlPart>();
        for (int index = 0; index < hfOps.Count; index++)
        {
            var diff = hfOps[index];
            if (!diff.CloneLeftPart || diff.LeftPartUri is not { } leftUri)
                continue;
            var sourcePart = FindHeaderFooterPart(main, diff.IsHeader, leftUri);
            var clone = sourcePart is null
                ? null
                : CloneHeaderFooterPart(sourcePart, diff.IsHeader, main, leftStreamDoc);
            if (clone is not null)
                outputParts[index] = clone;
        }

        for (int index = 0; index < hfOps.Count; index++)
        {
            var diff = hfOps[index];
            if (diff.LeftPartUri is { } leftUri)
            {
                // Matched (rebuild with token-level markup) or deleted-only (all content marked w:del —
                // the part and its reference stay; accept leaves an empty story, Word's own behavior).
                OpenXmlPart? part = diff.CloneLeftPart
                    ? outputParts.TryGetValue(index, out var clonedPart) ? clonedPart : null
                    : FindHeaderFooterPart(main, diff.IsHeader, leftUri);
                if (part is null)
                    continue; // left part vanished (malformed input) — keep the carry-over
                outputParts[index] = part;
                if (diff.RightPartUri is { } rightUri)
                    state.StoryOutputParts[rightUri] = part;
                // A topology-only record rebinds an inherited reference without changing the story's content.
                if (diff.Ops.Count > 0)
                    ApplyHeaderFooterDiffToPart(diff, part, state, settings, rightMain, leftStreamDoc, rightStreamDoc);
                if (diff.ReferenceBindings is not { Count: > 0 })
                    EnsureStoryReference(diff, part, main);
            }
            else
            {
                var part = InsertHeaderFooterStory(
                    diff, state, main, rightMain, settings, leftStreamDoc, rightStreamDoc);
                if (part is not null)
                    outputParts[index] = part;
            }
        }

        ApplyHeaderFooterReferenceBindings(hfOps, outputParts, main, rightMain);
    }

    /// <summary>
    /// Create a fresh header/footer part from a pristine LEFT story root before any redline rendering. The
    /// rebuilt story may contain deleted or moved LEFT content, so its original rIds must remain valid in the
    /// clone; connect the clone to the existing left-package targets under the SAME local relationship ids.
    /// RIGHT-sourced content receives fresh, isolated relationship ids during the normal story import below.
    /// </summary>
    private static OpenXmlPart? CloneHeaderFooterPart(
        OpenXmlPart sourcePart, bool isHeader, MainDocumentPart main,
        OpenXmlMemoryStreamDocument leftStreamDoc)
    {
        var sourceRoot = sourcePart.GetXDocument().Root;
        if (sourceRoot is null)
            return null;

        OpenXmlPart clonePart = isHeader
            ? main.AddNewPart<HeaderPart>()
            : main.AddNewPart<FooterPart>();
        var cloneRoot = new XElement(sourceRoot);
        var cloneXDoc = clonePart.GetXDocument();
        if (cloneXDoc.Root is null)
            cloneXDoc.Add(cloneRoot);
        else
            cloneXDoc.Root.ReplaceWith(cloneRoot);

        var sourcePackagePart = leftStreamDoc.GetPackage().GetPart(sourcePart.Uri);
        var clonePackagePart = leftStreamDoc.GetPackage().GetPart(clonePart.Uri);
        CopyStoryRelationshipsWithOriginalIds(sourcePackagePart, clonePackagePart);
        clonePart.PutXDocument();
        return clonePart;
    }

    /// <summary>
    /// Give a cloned story the LEFT source's relationship graph without changing its rIds. Header/footer parts
    /// live in the same package as their source, so their existing targets can be shared safely; the renderer
    /// never mutates those targets. This deliberately differs from cross-document imports, which must create
    /// isolated parts and rewrite only RIGHT-sourced XML.
    /// </summary>
    private static void CopyStoryRelationshipsWithOriginalIds(PackagePart sourcePart, PackagePart destinationPart)
    {
        foreach (var relationship in sourcePart.GetRelationships())
        {
            if (destinationPart.RelationshipExists(relationship.Id))
                continue;

            Uri targetUri = relationship.TargetUri;
            if (relationship.TargetMode == TargetMode.Internal)
            {
                try
                {
                    targetUri = PackUriHelper.ResolvePartUri(sourcePart.Uri, relationship.TargetUri);
                }
                catch (ArgumentException)
                {
                    // Preserve a malformed/dangling target verbatim. It was already present on the left source,
                    // and retaining its local rId is safer than silently rebinding it to an unrelated target.
                }
            }

            destinationPart.CreateRelationship(
                targetUri, relationship.TargetMode, relationship.RelationshipType, relationship.Id);
        }
    }

    /// <summary>
    /// Render a RIGHT-only (inserted) story: create a fresh header/footer part seeded from the right
    /// part's root shell, render the all-insert ops into it, attach a <c>w:headerReference</c>/
    /// <c>w:footerReference</c> (typed by the story's kind) to the output body's sectPr at the story's
    /// section ordinal, and ensure the visibility flag the story needs — <c>w:titlePg</c> on that sectPr
    /// for a First story, <c>w:evenAndOddHeaders</c> in the settings part for an Even story (ensured,
    /// not revision-tracked — sectPr/settings changes are outside w:sectPrChange scope in v1). Accept
    /// keeps the inserted content; reject strips it, leaving an EMPTY story — text-level ≡ the left's
    /// absent story, matching Word's own reject behavior for an inserted header.
    /// </summary>
    private static OpenXmlPart? InsertHeaderFooterStory(
        IrHeaderFooterDiff diff, RenderState state, MainDocumentPart main, MainDocumentPart? rightMain,
        IrDiffSettings settings, OpenXmlMemoryStreamDocument leftStreamDoc,
        OpenXmlMemoryStreamDocument rightStreamDoc)
    {
        if (rightMain is null || diff.RightPartUri is null)
            return null;
        var rightPart = FindHeaderFooterPart(rightMain, diff.IsHeader, diff.RightPartUri);
        var rightRoot = rightPart?.GetXDocument().Root;
        if (rightRoot is null)
            return null;

        // Locate the target sectPr FIRST — a story that cannot attach must not create an orphan part.
        // Document-order sectPr enumeration matches the reader's section ordinals. A section ordinal the
        // output body cannot represent (the section structure itself changed) is skipped: the revisions
        // surface still reports the story, the markup omission is a documented v1 ceiling.
        var mainXDoc = main.GetXDocument();
        var body = mainXDoc.Root?.Element(W.body);
        if (body is null)
            return null;
        // A w:sectPrChange's inner sectPr is change history, not a section — counting it mis-indexes
        // multi-section outputs.
        var sectPrs = body.Descendants(W.sectPr)
            .Where(s => s.Parent?.Name != W.sectPrChange)
            .ToList();
        if (diff.SectionIndex >= sectPrs.Count)
            return null;

        OpenXmlPart newPart = diff.IsHeader
            ? main.AddNewPart<HeaderPart>()
            : main.AddNewPart<FooterPart>();
        state.StoryOutputParts[diff.RightPartUri] = newPart;
        var newRoot = new XElement(rightRoot.Name, rightRoot.Attributes());
        var xDoc = newPart.GetXDocument();
        if (xDoc.Root == null)
            xDoc.Add(newRoot);
        else
            xDoc.Root.ReplaceWith(newRoot);

        int clonesBefore = state.RightSourcedClones.Count;
        var blocks = new List<XElement>();
        RenderBlockOpsWordShaped(diff.Ops, state, blocks);
        if (settings is { RenderMoves: true, SimplifyMoveMarkup: true })
            foreach (var b in blocks)
                SimplifyMoveMarkup(b);
        newRoot.Add(blocks);

        ImportStorySourcedRelationships(diff, state, clonesBefore, newPart, rightMain,
            leftStreamDoc, rightStreamDoc);

        foreach (var attr in newRoot.DescendantsAndSelf().Attributes()
                     .Where(a => a.Name.Namespace == PtOpenXml.pt).ToList())
            attr.Remove();
        newPart.PutXDocument();

        if (diff.ReferenceBindings is not { Count: > 0 })
        {
            // Legacy/single-cell shape: attach directly. Refined topology records attach all cells together
            // after every output part exists, so a clone can safely rejoin the original at a later section.
            BindHeaderFooterReference(diff.IsHeader, diff.Kind, diff.SectionIndex, newPart, main, rightMain);
        }
        return newPart;
    }

    /// <summary>Whether the right document's <paramref name="sectionIndex"/>-th section activates
    /// <c>w:titlePg</c> (sectPrChange inners excluded from the section count).</summary>
    private static bool RightSectionHasTitlePg(MainDocumentPart rightMain, int sectionIndex)
    {
        var body = rightMain.GetXDocument().Root?.Element(W.body);
        if (body is null)
            return false;
        var sectPrs = body.Descendants(W.sectPr)
            .Where(s => s.Parent?.Name != W.sectPrChange)
            .ToList();
        return sectionIndex < sectPrs.Count && sectPrs[sectionIndex].Element(W.titlePg) is not null;
    }

    /// <summary>
    /// Apply the static refined topology emitted by the builder. References themselves have no native revision
    /// representation, so each binding deliberately changes the output's section map while the bound story part
    /// contains the normal reversible content redline. A binding replaces an explicit same-kind ref when present
    /// (rather than appending a duplicate) and is otherwise inserted in schema order.
    /// </summary>
    private static void ApplyHeaderFooterReferenceBindings(
        IReadOnlyList<IrHeaderFooterDiff> hfOps, IReadOnlyDictionary<int, OpenXmlPart> outputParts,
        MainDocumentPart main, MainDocumentPart? rightMain)
    {
        for (int index = 0; index < hfOps.Count; index++)
        {
            var diff = hfOps[index];
            if (diff.ReferenceBindings is not { Count: > 0 } bindings ||
                !outputParts.TryGetValue(index, out var part))
                continue;
            foreach (var binding in bindings)
                BindHeaderFooterReference(diff.IsHeader, binding.Kind, binding.SectionIndex, part, main, rightMain);
        }
    }

    private static void BindHeaderFooterReference(
        bool isHeader, IrHeaderFooterKind kind, int sectionIndex, OpenXmlPart part,
        MainDocumentPart main, MainDocumentPart? rightMain)
    {
        var body = main.GetXDocument().Root?.Element(W.body);
        if (body is null)
            return;
        var sectPrs = body.Descendants(W.sectPr)
            .Where(s => s.Parent?.Name != W.sectPrChange)
            .ToList();
        if (sectionIndex < 0 || sectionIndex >= sectPrs.Count)
            return;

        var sectPr = sectPrs[sectionIndex];
        var refName = isHeader ? W.headerReference : W.footerReference;
        string typeValue = kind switch
        {
            IrHeaderFooterKind.First => "first",
            IrHeaderFooterKind.Even => "even",
            _ => "default",
        };
        string relationshipId = main.GetIdOfPart(part);
        var existing = sectPr.Elements(refName)
            .Where(reference => (string?)reference.Attribute(W.type) == typeValue)
            .ToList();
        if (existing.Count > 0)
        {
            existing[0].SetAttributeValue(R.id, relationshipId);
            foreach (var duplicate in existing.Skip(1))
                duplicate.Remove();
        }
        else
        {
            var newReference = new XElement(refName,
                new XAttribute(W.type, typeValue), new XAttribute(R.id, relationshipId));
            // CT_SectPr orders all headers before all footers. Insert a header before the first footer;
            // insert a footer after the final header (or before an existing footer when no headers exist).
            if (isHeader)
            {
                var firstFooter = sectPr.Element(W.footerReference);
                if (firstFooter is null)
                    sectPr.AddFirst(newReference);
                else
                    firstFooter.AddBeforeSelf(newReference);
            }
            else
            {
                var lastHeader = sectPr.Elements(W.headerReference).LastOrDefault();
                if (lastHeader is not null)
                    lastHeader.AddAfterSelf(newReference);
                else if (sectPr.Element(W.footerReference) is { } firstFooter)
                    firstFooter.AddBeforeSelf(newReference);
                else
                    sectPr.AddFirst(newReference);
            }
        }

        // Activate only the visibility flags genuinely active on the right. A latent First/Even ref must stay
        // latent; forcing a flag routes pages through an otherwise invisible, empty story.
        if (rightMain is not null && kind == IrHeaderFooterKind.First && sectPr.Element(W.titlePg) is null &&
            RightSectionHasTitlePg(rightMain, sectionIndex))
            InsertIntoSectPr(sectPr, new XElement(W.titlePg));
        if (rightMain?.DocumentSettingsPart?.GetXDocument().Root?.Element(W.evenAndOddHeaders) is not null &&
            kind == IrHeaderFooterKind.Even)
            WordprocessingMLUtil.EnsureEvenAndOddHeaders(main);
        main.PutXDocument();
    }

    /// <summary>
    /// A matched or deleted-only story merges into the LEFT part assuming "the part and its
    /// reference stay". When the left reference lived on an inline <c>w:sectPr</c> the body render
    /// did not preserve (the section structure collapsed), the merged part is ORPHANED — nothing
    /// renders and reject cannot restore the left story. Re-attach a reference of the story's
    /// kind/type at its section ordinal (clamped to the surviving structure) iff no reference of
    /// that kind/type is reachable there — a reference on this or any EARLIER section keeps the
    /// story reachable via OOXML inheritance.
    /// </summary>
    private static void EnsureStoryReference(IrHeaderFooterDiff diff, OpenXmlPart part, MainDocumentPart main)
    {
        var body = main.GetXDocument().Root?.Element(W.body);
        if (body is null)
            return;
        var sectPrs = body.Descendants(W.sectPr)
            .Where(s => s.Parent?.Name != W.sectPrChange)
            .ToList();
        if (sectPrs.Count == 0)
            return;
        int idx = Math.Min(diff.SectionIndex, sectPrs.Count - 1);
        var refName = diff.IsHeader ? W.headerReference : W.footerReference;
        string typeValue = diff.Kind switch
        {
            IrHeaderFooterKind.First => "first",
            IrHeaderFooterKind.Even => "even",
            _ => "default",
        };
        if (sectPrs.Take(idx + 1).Any(s =>
                s.Elements(refName).Any(e => (string?)e.Attribute(W.type) == typeValue)))
            return;
        sectPrs[idx].AddFirst(new XElement(refName,
            new XAttribute(W.type, typeValue),
            new XAttribute(R.id, main.GetIdOfPart(part))));
        main.PutXDocument();
    }

    /// <summary>Elements that FOLLOW <c>w:titlePg</c> in the CT_SectPr sequence — an insertion lands
    /// before the first of these (or at the end), keeping the sectPr schema-ordered.</summary>
    private static readonly XName[] SectPrAfterTitlePg =
        { W.textDirection, W.bidi, W.rtlGutter, W.docGrid, W.printerSettings, W.sectPrChange };

    private static void InsertIntoSectPr(XElement sectPr, XElement element)
    {
        var firstTail = sectPr.Elements().FirstOrDefault(e => SectPrAfterTitlePg.Contains(e.Name));
        if (firstTail is null)
            sectPr.Add(element);
        else
            firstTail.AddBeforeSelf(element);
    }

    /// <summary>The output package's header (or footer) part with the given URI, or null.</summary>
    private static OpenXmlPart? FindHeaderFooterPart(MainDocumentPart main, bool isHeader, Uri partUri)
    {
        var parts = isHeader ? main.HeaderParts.Cast<OpenXmlPart>() : main.FooterParts.Cast<OpenXmlPart>();
        return parts.FirstOrDefault(p => p.Uri == partUri);
    }

    /// <summary>
    /// Render one story diff's ops into <paramref name="part"/>: same recipe as a note scope — render,
    /// simplify moves, strip pt bookkeeping, replace the root's block-level children (keeping any
    /// non-block prelude) — plus the PER-PART media import: clones registered while rendering THIS
    /// story's ops are right-HEADER-part-sourced, so their relationships import from the right story
    /// part into this output part (not the main part).
    /// </summary>
    private static void ApplyHeaderFooterDiffToPart(
        IrHeaderFooterDiff diff, OpenXmlPart part, RenderState state, IrDiffSettings settings,
        MainDocumentPart? rightMain, OpenXmlMemoryStreamDocument leftStreamDoc,
        OpenXmlMemoryStreamDocument rightStreamDoc)
    {
        var xDoc = part.GetXDocument();
        var root = xDoc.Root;
        if (root is null)
            return;

        // Slice the clone registry around this story's render so ONLY its clones import into this part.
        int clonesBefore = state.RightSourcedClones.Count;
        var blocks = new List<XElement>();
        RenderBlockOpsWordShaped(diff.Ops, state, blocks);
        if (settings is { RenderMoves: true, SimplifyMoveMarkup: true })
            foreach (var b in blocks)
                SimplifyMoveMarkup(b);

        root.Elements().Where(e => e.Name == W.p || e.Name == W.tbl || e.Name == W.sdt).Remove();
        root.Add(blocks);

        ImportStorySourcedRelationships(diff, state, clonesBefore, part, rightMain,
            leftStreamDoc, rightStreamDoc);

        foreach (var attr in root.DescendantsAndSelf().Attributes()
                     .Where(a => a.Name.Namespace == PtOpenXml.pt).ToList())
            attr.Remove();
        part.PutXDocument();
    }

    /// <summary>
    /// Import media parts + hyperlink/external relationships referenced by the clones registered during
    /// one story's render (<paramref name="clonesBefore"/> marks the registry watermark) from the RIGHT
    /// story part into the output story part. Relationship ids are part-scoped in OOXML, which is why the
    /// body's main-part import cannot serve header/footer content.
    /// </summary>
    private static void ImportStorySourcedRelationships(
        IrHeaderFooterDiff diff, RenderState state, int clonesBefore, OpenXmlPart outputPart,
        MainDocumentPart? rightMain, OpenXmlMemoryStreamDocument leftStreamDoc,
        OpenXmlMemoryStreamDocument rightStreamDoc)
    {
        if (rightMain is null || diff.RightPartUri is null)
            return;
        var storyClones = state.RightSourcedClones.Skip(clonesBefore).ToList();
        if (storyClones.Count == 0)
            return;
        var rightPart = FindHeaderFooterPart(rightMain, diff.IsHeader, diff.RightPartUri);
        if (rightPart is null)
            return;

        ImportHyperlinkAndExternalRelationships(storyClones, outputPart, rightPart);

        var outPkgPart = leftStreamDoc.GetPackage().GetPart(outputPart.Uri);
        var rightPkgPart = rightStreamDoc.GetPackage().GetPart(rightPart.Uri);
        foreach (var clone in storyClones)
            WmlComparer.MoveRelatedPartsToDestination(
                rightPkgPart, outPkgPart, clone, skipDanglingRelationships: true,
                skipHeaderFooterReferences: true);
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
    /// <para><b>Nested references.</b> The body walk cannot see a reference living INSIDE a note's definition
    /// body (a note that itself cites a note). Each renumbered definition's old→new id is therefore returned
    /// to the caller, which — after BOTH kinds' passes ran — sweeps BOTH note parts remapping nested
    /// references of BOTH kinds (<see cref="RemapNestedNoteReferences"/>). Same-kind nesting (a footnote
    /// citing a footnote) and CROSS-kind nesting (a footnote citing an endnote, or vice versa) are both
    /// covered: the cross-kind case needs the OTHER kind's remap applied to THIS kind's part, which is why the
    /// sweep runs once over both parts with both maps rather than per-pass.</para>
    /// <para><b>Known limitation (unexercised in the M2.6 corpus; documented per the T1 review).</b>
    /// <i>Deleted EMPTY-bodied note dequeue keys on <c>w:delText</c>.</i> <c>IsDeletedOnly</c> classifies a
    /// definition as deleted-only via "has <c>w:delText</c> and no live <c>w:t</c>"; a deleted note whose body
    /// carries NO text at all (no <c>w:delText</c>, no <c>w:t</c>) is therefore not enqueued in <c>delDefs</c>,
    /// so a <c>w:del</c> body reference could dequeue the wrong deleted def (or none). No corpus fixture has a
    /// textless deleted note; a robust fix would key deletedness on the reference/definition correspondence the
    /// builder already records rather than on body text presence.</para>
    /// </summary>
    /// <returns>Each renumbered definition's OLD id → NEW id (empty when nothing renumbered), for the caller's
    /// nested-reference sweep across both note parts.</returns>
    internal static Dictionary<string, string> RenumberNoteIds(MainDocumentPart main, XName refName, XName noteName, XName rootName,
        OpenXmlPart? notePart, OpenXmlPart? rightNotePart)
    {
        var empty = new Dictionary<string, string>(StringComparer.Ordinal);
        if (notePart == null)
            return empty;
        var noteXDoc = notePart.GetXDocument();
        var noteRoot = noteXDoc.Root;
        if (noteRoot == null)
            return empty;

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
            return empty;
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
        // Each renumbered definition's OLD id → NEW id, so a reference to it that the body walk did NOT visit
        // (one nested INSIDE another note's body — a note that cites a note) can be remapped afterwards. Without
        // this the nested ref keeps the old id while its definition moves, dangling on accept/reject. Returned
        // to the caller, which sweeps BOTH note parts once BOTH kinds' remaps exist (cross-kind nesting).
        var idRemap = new Dictionary<string, string>(StringComparer.Ordinal);
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
                var defOldId = (string?)def.Attribute(W.id);
                def.SetAttributeValue(W.id, newId);
                if (defOldId != null && defOldId != newId) idRemap[defOldId] = newId;
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

        // NESTED references (same-kind AND cross-kind) are remapped by the caller's
        // RemapNestedNoteReferences sweep, which runs after BOTH kinds' renumber passes so a footnote-nested
        // endnote reference sees the endnote remap and vice versa.
        return idRemap;
    }

    /// <summary>
    /// Remap NESTED note references — references living inside another note's definition body, which the
    /// body-order renumber walk never visits — across BOTH note parts, for BOTH kinds. A footnote body can
    /// nest a footnote reference (same-kind) or an endnote reference (CROSS-kind), and vice versa; each nested
    /// reference whose definition was renumbered must follow it or it dangles on accept/reject. Runs once,
    /// after both <see cref="RenumberNoteIds"/> passes, with both kinds' old→new maps — a per-pass sweep
    /// cannot fix cross-kind nesting because the other kind's remap does not exist yet.
    /// </summary>
    internal static void RemapNestedNoteReferences(MainDocumentPart main,
        Dictionary<string, string> footnoteRemap, Dictionary<string, string> endnoteRemap)
    {
        if (footnoteRemap.Count == 0 && endnoteRemap.Count == 0)
            return;
        foreach (var part in new OpenXmlPart?[] { main.FootnotesPart, main.EndnotesPart })
        {
            var root = part?.GetXDocument().Root;
            if (root == null)
                continue;
            bool changed = false;
            foreach (var nestedRef in root.Descendants()
                         .Where(e => e.Name == W.footnoteReference || e.Name == W.endnoteReference))
            {
                var remap = nestedRef.Name == W.footnoteReference ? footnoteRemap : endnoteRemap;
                var id = (string?)nestedRef.Attribute(W.id);
                if (id != null && remap.TryGetValue(id, out var mapped))
                {
                    nestedRef.SetAttributeValue(W.id, mapped);
                    changed = true;
                }
            }
            if (changed)
                part!.PutXDocument();
        }
    }

    // ----------------------------------------------------------------- comment normalization

    private static readonly XNamespace W14ns = "http://schemas.microsoft.com/office/word/2010/wordml";
    private static readonly XNamespace W15ns = "http://schemas.microsoft.com/office/word/2012/wordml";
    private static readonly XNamespace W16cidNs = "http://schemas.microsoft.com/office/word/2016/wordml/cid";

    /// <summary>The set of comment ids ANCHORED in a source IR document's body — scanned from each block's
    /// provenance source element for <c>w:commentRangeStart</c>/<c>w:commentReference</c> ids (the comment
    /// analogue of <see cref="BodyBookmarkNames"/>). Used to classify a rendered comment as common (in both
    /// sources) / right-added / left-deleted for round-trip-correct normalization.</summary>
    internal static HashSet<string> BodyCommentIds(IrDocument? source)
    {
        var ids = new HashSet<string>();
        if (source == null)
            return ids;
        foreach (var block in source.Body.Blocks)
        {
            var el = block.Source.Element;
            if (el == null)
                continue;
            foreach (var m in el.DescendantsAndSelf()
                         .Where(e => e.Name == W.commentRangeStart || e.Name == W.commentReference))
                if ((string?)m.Attribute(W.id) is { Length: > 0 } i)
                    ids.Add(i);
        }
        return ids;
    }

    /// <summary>
    /// Merge RIGHT-only comment definitions (+ their <c>commentsExtended</c> reply links and
    /// <c>commentsIds</c> durable ids) into the output's LEFT-based comments part. The fine path emits
    /// right-sourced content (equal/insert spans, EqualBlock) verbatim, so a comment ADDED in the right document
    /// has its markers in the body but its <c>w:comment</c> definition only in the RIGHT package — it would
    /// dangle. Copy each referenced-but-undefined right comment in (creating the comments part if the left had
    /// none), then carry its threading metadata so a right-added reply still links to its parent. No-op when the
    /// right has no comments part or every referenced comment already resolves.
    /// </summary>
    internal static void MergeRightCommentDefinitions(
        MainDocumentPart main,
        MainDocumentPart? rightMain,
        OpenXmlMemoryStreamDocument outputStream,
        OpenXmlMemoryStreamDocument rightStream)
    {
        var rightComments = rightMain?.WordprocessingCommentsPart;
        var rightRoot = rightComments?.GetXDocument().Root;
        if (rightRoot == null)
            return;
        var body = main.GetXDocument().Root?.Element(W.body);
        if (body == null)
            return;

        var referenced = body.Descendants()
            .Where(e => e.Name == W.commentRangeStart || e.Name == W.commentRangeEnd || e.Name == W.commentReference)
            .Select(e => (string?)e.Attribute(W.id)).Where(id => id != null).ToHashSet();

        var have = main.WordprocessingCommentsPart?.GetXDocument().Root?.Elements(W.comment)
            .Select(c => (string?)c.Attribute(W.id)).Where(id => id != null).ToHashSet() ?? new HashSet<string?>();
        var toAdd = referenced.Where(id => !have.Contains(id))
            .Select(id => rightRoot.Elements(W.comment).FirstOrDefault(c => (string?)c.Attribute(W.id) == id))
            .Where(def => def != null).Select(def => def!).ToList();
        if (toAdd.Count == 0)
            return;

        var outRoot = EnsureCommentsRoot(main);
        var outputComments = main.WordprocessingCommentsPart!;
        var addedDefinitions = new List<XElement>(toAdd.Count);
        var addedParaIds = new HashSet<string>();
        foreach (var def in toAdd)
        {
            // Keep the live clone: relationship import rewrites its r:embed/r:id attributes in place.
            // Comment definitions live in comments.xml, not document.xml, so the ordinary body media import
            // cannot repair these references after the XML has been copied into the left-based package.
            var clone = new XElement(def);
            outRoot.Add(clone);
            addedDefinitions.Add(clone);
            foreach (var pid in def.Descendants().Attributes((W14ns + "paraId")))
                if ((string?)pid is { Length: > 0 } v) addedParaIds.Add(v);
        }

        // Import relationships against their OWNERS: both r:embed media and w:hyperlink r:ids in a
        // comment resolve from comments.xml. The left comments part may already own the same rId for a
        // different image/target, so the shared import routines remap the live clone to a fresh output id.
        ImportHyperlinkAndExternalRelationships(addedDefinitions, outputComments, rightComments!);
        var outputPackagePart = outputStream.GetPackage().GetPart(outputComments.Uri);
        var rightPackagePart = rightStream.GetPackage().GetPart(rightComments!.Uri);
        foreach (var clone in addedDefinitions)
            WmlComparer.MoveRelatedPartsToDestination(
                rightPackagePart, outputPackagePart, clone, skipDanglingRelationships: true,
                skipHeaderFooterReferences: true);

        main.WordprocessingCommentsPart!.PutXDocument();

        // Carry threading metadata for the merged comments (paraId-keyed), so a right-added reply keeps its link.
        if (addedParaIds.Count > 0)
        {
            // Guard the commentsExtended merge on the right actually HAVING that part (mirrors the commentsIds
            // guard below). Otherwise the eager `?? AddNewPart<…>()` creates an empty /word/commentsExtended.xml on
            // the LEFT-derived output, and MergeRightThreadingEntries early-returns on the null rightRoot BEFORE
            // seeding it — leaving a rootless part (Sch_MissingPartRootElement) for the common case of a
            // non-threaded right-added comment whose definition paragraph carries a w14:paraId.
            if (rightMain!.WordprocessingCommentsExPart != null)
                MergeRightThreadingEntries(
                    main.WordprocessingCommentsExPart ?? main.AddNewPart<WordprocessingCommentsExPart>(),
                    rightMain.WordprocessingCommentsExPart,
                    W15ns + "commentsEx", W15ns + "commentEx", W15ns + "paraId", addedParaIds,
                    $"<w15:commentsEx xmlns:w=\"{W.w.NamespaceName}\" xmlns:w15=\"{W15ns.NamespaceName}\"/>");
            if (rightMain.WordprocessingCommentsIdsPart != null)
                MergeRightThreadingEntries(
                    main.WordprocessingCommentsIdsPart ?? main.AddNewPart<WordprocessingCommentsIdsPart>(),
                    rightMain.WordprocessingCommentsIdsPart,
                    W16cidNs + "commentsIds", W16cidNs + "commentId", W16cidNs + "paraId", addedParaIds,
                    $"<w16cid:commentsIds xmlns:w=\"{W.w.NamespaceName}\" xmlns:w16cid=\"{W16cidNs.NamespaceName}\"/>");
        }
    }

    /// <summary>Copy every entry of <paramref name="rightPart"/> whose paraId attribute is in
    /// <paramref name="paraIds"/> into <paramref name="outPart"/> (seeded from <paramref name="seedXml"/> when
    /// empty), skipping paraIds already present. Shared by the <c>commentsExtended</c> + <c>commentsIds</c>
    /// threading merges.</summary>
    private static void MergeRightThreadingEntries(OpenXmlPart outPart, OpenXmlPart? rightPart,
        XName rootName, XName entryName, XName paraIdAttr, HashSet<string> paraIds, string seedXml)
    {
        var rightRoot = rightPart?.GetXDocument().Root;
        if (rightRoot == null)
            return;
        var outRoot = EnsurePartRoot(outPart, seedXml);
        var present = outRoot.Elements(entryName).Select(e => (string?)e.Attribute(paraIdAttr)).ToHashSet();
        bool changed = false;
        foreach (var entry in rightRoot.Elements(entryName)
                     .Where(e => (string?)e.Attribute(paraIdAttr) is { } p && paraIds.Contains(p) && !present.Contains(p)))
        {
            outRoot.Add(new XElement(entry));
            changed = true;
        }
        if (changed)
            outPart.PutXDocument();
    }

    /// <summary>Get the output comments part's root, creating an empty <c>w:comments</c> part if the left had
    /// none (a right-added comment with no left comments at all).</summary>
    private static XElement EnsureCommentsRoot(MainDocumentPart main)
    {
        var part = main.WordprocessingCommentsPart ?? main.AddNewPart<WordprocessingCommentsPart>();
        return EnsurePartRoot(part, $"<w:comments xmlns:w=\"{W.w.NamespaceName}\" xmlns:w14=\"{W14ns.NamespaceName}\"/>");
    }

    /// <summary>Return a part's XDocument root, seeding it from <paramref name="seedXml"/> when the part is
    /// empty (a freshly <c>AddNewPart</c>'d part has an empty stream — <see cref="PtOpenXmlExtensions.GetXDocument"/>
    /// returns a rootless <see cref="XDocument"/>). The seed root is added to the SAME cached XDocument the caller
    /// then mutates + <c>PutXDocument</c>s, so there is no stream/annotation staleness.</summary>
    private static XElement EnsurePartRoot(OpenXmlPart part, string seedXml)
    {
        var xdoc = part.GetXDocument();
        if (xdoc.Root == null)
            xdoc.Add(XElement.Parse(seedXml));
        return xdoc.Root!;
    }

    /// <summary>
    /// Reconcile the rendered body's comment markers so every <c>w:commentReference</c> resolves to exactly one
    /// <c>w:comment</c>, every <c>w:commentRangeStart</c> id is unique and pairs 1:1 with a
    /// <c>w:commentRangeEnd</c> of the same id, and an unchanged comment survives BOTH accept and reject. The
    /// comment analogue of <see cref="NormalizeBookmarks"/>; three phases:
    /// <list type="bullet">
    /// <item><b>(A) collapse.</b> A comment present in BOTH sources (common) that still has a BARE marker is
    /// position-stable (it anchors equal content): collapse each of its three marker kinds to a single bare
    /// survivor (lifting one out of a <c>w:ins</c>/<c>w:del</c> wrapper if needed) so it survives accept AND
    /// reject. The fine path can emit the SAME marker both bare (from the right model's equal span) and
    /// revision-wrapped (from the left model's del span) at an edit boundary — this de-duplicates that.</item>
    /// <item><b>(B) dedup.</b> A comment whose markers are wholly inside BOTH a <c>w:del</c> and a <c>w:ins</c>
    /// subtree (a wholly rewritten anchor, or a comment in a whole-block-bailed table/opaque) has no bare
    /// survivor: give the DELETED copy a fresh id + a cloned definition (paraId-stripped so threading never
    /// collides). Accept ≡ right comment, reject ≡ left comment.</item>
    /// <item><b>(C) pair + resolve.</b> Drop any marker/reference whose id has no definition; re-close an
    /// orphaned start with a synthetic zero-width end in the start's own context; drop an orphaned end.</item>
    /// </list>
    /// The blessed <see cref="WmlComparer"/> oracle drops ALL comments on any edit; this preserves them with
    /// fine per-word markup.
    /// </summary>
    internal static void NormalizeComments(MainDocumentPart main, HashSet<string> leftIds, HashSet<string> rightIds,
        RenderState state)
    {
        var doc = main.GetXDocument();
        var body = doc.Root?.Element(W.body);
        if (body == null)
            return;
        var commentsPart = main.WordprocessingCommentsPart;

        static string? IdOf(XElement e) => (string?)e.Attribute(W.id);
        static bool Inside(XElement e, XName wrapper) => e.Ancestors().Any(a => a.Name == wrapper);
        static bool IsBare(XElement e) => !e.Ancestors().Any(a => a.Name == W.ins || a.Name == W.del);
        List<XElement> Markers() => body.Descendants()
            .Where(e => (e.Name == W.commentRangeStart || e.Name == W.commentRangeEnd || e.Name == W.commentReference)
                        && IsRunLevelCommentMarker(e)).ToList();

        bool changed = false;

        // (A) Identity-aware collapse — common comment with a bare survivor → single bare marker per kind.
        foreach (var id in Markers().Select(IdOf).Where(i => i != null).Distinct().ToList())
        {
            if (!(leftIds.Contains(id!) && rightIds.Contains(id!)))
                continue; // right-added / left-deleted: keep its revision context
            var all = Markers().Where(m => IdOf(m) == id).ToList();
            if (!all.Any(IsBare))
                continue; // all wrapped (rewritten anchor) → (B)
            foreach (var kind in new[] { W.commentRangeStart, W.commentRangeEnd, W.commentReference })
            {
                var kinds = all.Where(m => m.Name == kind).ToList();
                if (kinds.Count == 0)
                    continue;
                // Pick the survivor by POSITION so the collapsed range ENCLOSES the surviving content on BOTH
                // sides, then lift it bare: the commentRangeStart must precede both the w:del and the w:ins (the
                // leftmost copy in document order), the commentRangeEnd must follow both (the RIGHTMOST copy).
                // Keeping a del-side end — even a BARE one that happens to sit before the deleted run — would
                // leave the inserted (accept) / deleted (reject) content OUTSIDE the range, bracketing nothing for
                // a wholly-rewritten or wholly-deleted anchor. Position is what matters; the lift is a no-op when
                // the survivor is already bare. The reference is zero-width: either copy resolves the same.
                var keep = kind == W.commentRangeEnd ? kinds[kinds.Count - 1] : kinds[0];
                foreach (var m in kinds)
                    if (!ReferenceEquals(m, keep)) { RemoveCommentMarker(m); changed = true; }
                if (!IsBare(keep) && LiftCommentMarkerBare(keep))
                    changed = true;
            }
        }

        // (A2) A comment present in only ONE source is added/deleted by the edit. A BARE marker (it landed in
        //      equal content) would survive BOTH accept and reject — leaking a right-added comment into reject,
        //      or keeping a left-deleted comment on accept. Wrap each bare marker in the matching revision
        //      element so it toggles with its side: right-added → w:ins (reject drops it); left-deleted → w:del
        //      (accept drops it). Markers already in their revision context are left alone.
        foreach (var id in Markers().Select(IdOf).Where(i => i != null).Distinct().ToList())
        {
            bool rightAdded = rightIds.Contains(id!) && !leftIds.Contains(id!);
            bool leftDeleted = leftIds.Contains(id!) && !rightIds.Contains(id!);
            if (!rightAdded && !leftDeleted)
                continue;
            var kind = rightAdded ? RevKind.Ins : RevKind.Del;
            foreach (var m in Markers().Where(m => IdOf(m) == id && IsBare(m)).ToList())
            {
                WrapCommentMarkerInRevision(m, kind, state);
                changed = true;
            }
        }

        // (B) Renumber/dedup remaining del∩ins duplicates → fresh id + cloned definition for the deleted copy.
        var commentsRoot = commentsPart?.GetXDocument().Root;
        if (commentsRoot != null)
        {
            var markers = Markers();
            var delIds = markers.Where(m => Inside(m, W.del)).Select(IdOf).Where(i => i != null).ToHashSet();
            var insIds = markers.Where(m => Inside(m, W.ins)).Select(IdOf).Where(i => i != null).ToHashSet();
            var duplicated = delIds.Where(insIds.Contains).ToList();
            if (duplicated.Count > 0)
            {
                int nextId = commentsRoot.Elements(W.comment)
                    .Select(c => int.TryParse((string?)c.Attribute(W.id), out var v) ? v : -1)
                    .DefaultIfEmpty(-1).Max() + 1;
                int nextParaId = MaxParaId(commentsRoot) + 1;
                foreach (var oldId in duplicated)
                {
                    var newId = (nextId++).ToString();
                    foreach (var m in markers.Where(m => IdOf(m) == oldId && Inside(m, W.del)))
                        m.SetAttributeValue(W.id, newId);
                    var def = commentsRoot.Elements(W.comment).FirstOrDefault(c => (string?)c.Attribute(W.id) == oldId);
                    if (def != null)
                    {
                        var copy = new XElement(def);
                        copy.SetAttributeValue(W.id, newId);
                        // Give the clone its OWN w14:paraId(s) so it never duplicates a threading key — and carry
                        // its commentsExtended/commentsIds entry under the fresh paraId so the REJECT-side clone
                        // (a threaded reply, say) keeps its parent link, exactly as the accept-side original does.
                        foreach (var pAttr in copy.Descendants().Attributes(W14ns + "paraId").ToList())
                        {
                            var oldPara = (string)pAttr;
                            var freshPara = (nextParaId++).ToString("X8");
                            pAttr.Value = freshPara;
                            CloneThreadingEntryForParaId(main, oldPara, freshPara);
                        }
                        commentsRoot.Add(copy);
                    }
                }
                commentsPart!.PutXDocument();
                changed = true;
            }
        }

        // (C) Pair + resolve. Drop unresolvable markers, then guarantee 1:1 start↔end pairing.
        var defIds = commentsPart?.GetXDocument().Root?.Elements(W.comment)
            .Select(c => (string?)c.Attribute(W.id)).Where(i => i != null).ToHashSet() ?? new HashSet<string?>();
        foreach (var m in Markers().Where(m => !defIds.Contains(IdOf(m))).ToList())
        {
            RemoveCommentMarker(m);
            changed = true;
        }
        var liveMarkers = Markers();
        var startsById = liveMarkers.Where(m => m.Name == W.commentRangeStart)
            .GroupBy(m => IdOf(m) ?? "").ToDictionary(g => g.Key, g => g.ToList());
        var endsById = liveMarkers.Where(m => m.Name == W.commentRangeEnd)
            .GroupBy(m => IdOf(m) ?? "").ToDictionary(g => g.Key, g => g.ToList());
        foreach (var (id, sl) in startsById)
        {
            int have = endsById.TryGetValue(id, out var el) ? el.Count : 0;
            for (int k = have; k < sl.Count; k++)
            {
                sl[k].AddAfterSelf(new XElement(W.commentRangeEnd, new XAttribute(W.id, id)));
                changed = true;
            }
        }
        foreach (var (id, el) in endsById)
        {
            int have = startsById.TryGetValue(id, out var sl) ? sl.Count : 0;
            for (int k = have; k < el.Count; k++)
            {
                RemoveCommentMarker(el[k]);
                changed = true;
            }
        }

        if (changed)
            main.PutXDocument();
    }

    /// <summary>True iff the comment marker sits at the paragraph/body run level (its ancestors up to the
    /// enclosing <c>w:p</c>/<c>w:body</c>/<c>w:tc</c> are only run-level wrappers — including the <c>w:r</c> that
    /// hosts a <c>w:commentReference</c>). A marker reached through opaque content (math, drawing, textbox) is
    /// part of that element's content hash and must NOT be reconciled.</summary>
    private static bool IsRunLevelCommentMarker(XElement marker)
    {
        for (var a = marker.Parent; a != null; a = a.Parent)
        {
            if (a.Name == W.p || a.Name == W.body || a.Name == W.tc)
                return true;
            if (a.Name != W.r && a.Name != W.ins && a.Name != W.del && a.Name != W.hyperlink &&
                a.Name != W.sdt && a.Name != W.sdtContent && a.Name != W.smartTag && a.Name != W.fldSimple)
                return false;
        }
        return false;
    }

    /// <summary>Remove a comment marker, dropping an emptied <c>w:ins</c>/<c>w:del</c> wrapper. A
    /// <c>w:commentReference</c> lives in a <c>w:r</c>: drop the whole run when it carries no text, else just the
    /// reference element.</summary>
    private static void RemoveCommentMarker(XElement marker)
    {
        if (marker.Name == W.commentReference)
        {
            var run = marker.Parent;
            if (run is { Name: var n } && n == W.r &&
                !run.Elements().Any(e => e.Name == W.t || e.Name == W.delText))
            {
                var wrapper = run.Parent;
                run.Remove();
                if (wrapper != null && (wrapper.Name == W.ins || wrapper.Name == W.del) && !wrapper.Elements().Any())
                    wrapper.Remove();
                return;
            }
            marker.Remove();
            return;
        }
        var parent = marker.Parent;
        if (parent == null)
            return;
        marker.Remove();
        if ((parent.Name == W.ins || parent.Name == W.del) && parent.Parent != null && !parent.Elements().Any())
            parent.Remove();
    }

    /// <summary>Lift a comment marker out of its sole-child <c>w:ins</c>/<c>w:del</c> wrapper to bare (a start
    /// before the wrapper, an end after; a reference run via <see cref="LiftRunBare"/>) so it survives BOTH
    /// accept and reject. Returns true if it moved.</summary>
    private static bool LiftCommentMarkerBare(XElement marker)
    {
        if (marker.Name == W.commentReference)
        {
            var run = marker.Parent;
            return run is { Name: var rn } && rn == W.r && LiftRunBare(run);
        }
        var parent = marker.Parent;
        if (parent == null || (parent.Name != W.ins && parent.Name != W.del) || parent.Parent == null)
            return false;
        marker.Remove();
        if (marker.Name == W.commentRangeEnd) parent.AddAfterSelf(marker);
        else parent.AddBeforeSelf(marker);
        if (!parent.Elements().Any())
            parent.Remove();
        return true;
    }

    /// <summary>The largest 8-hex <c>w14:paraId</c> present on any comment-definition paragraph (so a freshly
    /// allocated clone paraId collides with none).</summary>
    private static int MaxParaId(XElement commentsRoot)
    {
        int max = 0;
        foreach (var v in commentsRoot.Descendants().Attributes(W14ns + "paraId").Select(a => (string)a))
            if (int.TryParse(v, System.Globalization.NumberStyles.HexNumber, null, out var n) && n > max)
                max = n;
        return max;
    }

    /// <summary>Carry a comment's threading metadata onto a freshly-paraId'd dedup clone: clone the
    /// <c>commentsExtended</c> (<c>w15:commentEx</c>) and <c>commentsIds</c> (<c>w16cid:commentId</c>) entries that
    /// keyed off <paramref name="oldParaId"/> under <paramref name="freshParaId"/>, preserving each entry's
    /// <c>paraIdParent</c>/<c>durableId</c>. So the reject-side clone of a threaded reply still links to its
    /// parent. No-op when the part/entry is absent.</summary>
    private static void CloneThreadingEntryForParaId(MainDocumentPart main, string oldParaId, string freshParaId)
    {
        void Clone(OpenXmlPart? part, XName entryName, XName paraIdAttr)
        {
            var root = part?.GetXDocument().Root;
            var src = root?.Elements(entryName).FirstOrDefault(e => (string?)e.Attribute(paraIdAttr) == oldParaId);
            if (root == null || src == null)
                return;
            var copy = new XElement(src);
            copy.SetAttributeValue(paraIdAttr, freshParaId);
            root.Add(copy);
            part!.PutXDocument();
        }
        Clone(main.WordprocessingCommentsExPart, W15ns + "commentEx", W15ns + "paraId");
        Clone(main.WordprocessingCommentsIdsPart, W16cidNs + "commentId", W16cidNs + "paraId");
    }

    /// <summary>Wrap a BARE comment marker in a <c>w:ins</c>/<c>w:del</c> so it toggles with its revision side (a
    /// right-added or left-deleted comment whose marker landed in equal content). The <c>commentReference</c>'s
    /// host <c>w:r</c> is wrapped (a marker element is wrapped directly); both are valid children of
    /// <c>w:ins</c>/<c>w:del</c>.</summary>
    private static void WrapCommentMarkerInRevision(XElement marker, RevKind kind, RenderState state)
    {
        var target = marker;
        if (marker.Name == W.commentReference && marker.Parent?.Name == W.r)
            target = marker.Parent;
        if (target.Parent == null)
            return;
        var rev = new XElement(RevElementName(kind), state.RevisionAttributes());
        target.ReplaceWith(rev);
        rev.Add(target);
    }

    // ----------------------------------------------------------------- bookmark normalization

    /// <summary>
    /// Reconcile bookmark markers in the rendered body so every bookmark has a UNIQUE id, a 1:1
    /// start↔end pairing, and its NAME intact. An edit straddling a bookmark range endpoint — or a
    /// whole-block <c>del(left)+ins(right)</c> bail of a bookmark-bearing paragraph — emits the SAME
    /// bookmark id (and name) on both the <c>w:del</c> (left) and <c>w:ins</c> (right) side, which is
    /// schema-invalid (<c>Sem_UniqueAttributeValue</c>) and ambiguous to any cross-reference resolver.
    /// <para>Two cases, distinguished by whether the bookmark survives as a BARE (untracked) marker:</para>
    /// <list type="bullet">
    /// <item><b>Touches an equal region</b> (a bare endpoint exists): the bookmark is unchanged across the
    /// edit, so collapse it to a SINGLE bare start+end (lift any tracked copy out of its <c>w:ins</c>/
    /// <c>w:del</c> wrapper). It then survives BOTH accept and reject, wrapping the surviving content.</item>
    /// <item><b>Wholly inside a rewritten / whole-block-bailed span</b> (no bare endpoint): keep BOTH tracked
    /// copies but renumber them to unique ids. Accept drops the <c>w:del</c> copy and keeps the <c>w:ins</c>
    /// one; reject the reverse — each resolution lands exactly one, paired, name-preserved bookmark.</item>
    /// </list>
    /// Either way the Compare output carries unique, 1:1-paired bookmarks with names intact, so every
    /// <c>REF</c>/<c>PAGEREF</c>/<c>NOTEREF</c>/<c>HYPERLINK \l</c> field and internal hyperlink anchor still
    /// resolves. The blessed <see cref="WmlComparer"/> oracle strips ALL bookmarks on any edit; this preserves
    /// them.
    /// </summary>
    private static void NormalizeBookmarks(MainDocumentPart main, HashSet<string> leftNames, HashSet<string> rightNames)
    {
        var doc = main.GetXDocument();
        var body = doc.Root?.Element(W.body);
        if (body == null)
            return;

        // Only RUN-LEVEL bookmarks are reconciled here. A bookmark NESTED inside opaque content (a math
        // m:oMath, a w:drawing, a textbox) is part of that element's canonical content hash — renumbering or
        // removing it would silently change the opaque blob and break reject ≡ left (the IR hashes the whole
        // oMath, bookmark included). Such bookmarks round-trip WITH their opaque host untouched.
        var starts = body.Descendants(W.bookmarkStart).Where(IsRunLevelBookmark).ToList();
        var ends = body.Descendants(W.bookmarkEnd).Where(IsRunLevelBookmark).ToList();
        if (starts.Count == 0 && ends.Count == 0)
            return;

        static string? IdOf(XElement e) => (string?)e.Attribute(W.id);

        bool changed = false;

        // (A) Identity-aware collapse. A bookmark present in BOTH sources is UNCHANGED by the edit (only the text
        // around it moved), so it should be a single BARE pair that survives accept AND reject — collapsing the
        // del/ins copies (or recovering a side the dense-layout renderer dropped, which would otherwise leave the
        // bookmark surviving only one resolution). EXCEPTION: a whole-block-bailed paragraph (its paragraph mark
        // is itself tracked) cannot host a bare marker that survives both — there the del copy + ins copy each
        // survive their own resolution, so keep both and let (B) renumber them unique. A genuinely inserted /
        // deleted bookmark (one source only) keeps its w:ins/w:del context untouched.
        foreach (var grp in starts.Where(s => (string?)s.Attribute(W.name) != null)
                     .GroupBy(s => (string)s.Attribute(W.name)!).ToList())
        {
            var name = grp.Key;
            var nameStarts = grp.ToList();
            var ids = new HashSet<string>(nameStarts.Select(IdOf).Where(i => i != null)!);
            var nameEnds = ends.Where(e => IdOf(e) is { } id && ids.Contains(id)).ToList();

            bool common = leftNames.Contains(name) && rightNames.Contains(name);
            if (!common)
                continue; // inserted/deleted: leave revision context; (B)/(C) keep ids unique + paired
            if (nameStarts.Concat(nameEnds).Any(IsInWholeBlockRevisedParagraph))
                continue; // whole-block: del + ins copies each carry the name into their own resolution

            // Collapse to a single bare start + bare end (both survive accept AND reject).
            var keepStart = nameStarts[0];
            var keepEnd = nameEnds.Count > 0 ? nameEnds[0] : null;
            foreach (var s in nameStarts)
                if (!ReferenceEquals(s, keepStart)) { RemoveBookmarkMarker(s); changed = true; }
            foreach (var e in nameEnds)
                if (!ReferenceEquals(e, keepEnd)) { RemoveBookmarkMarker(e); changed = true; }
            if (LiftBookmarkBare(keepStart)) changed = true;
            if (keepEnd != null && LiftBookmarkBare(keepEnd)) changed = true;
        }

        // (B) Renumber the remaining duplicate ids (bookmarks wholly inside a rewritten / whole-block span) so
        //     each tracked copy is unique. Re-pair by document order; copy 0 keeps the id, copies 1.. get fresh.
        var liveStarts = body.Descendants(W.bookmarkStart).Where(IsRunLevelBookmark)
            .GroupBy(s => IdOf(s) ?? "").ToDictionary(g => g.Key, g => g.ToList());
        var liveEnds = body.Descendants(W.bookmarkEnd).Where(IsRunLevelBookmark)
            .GroupBy(e => IdOf(e) ?? "").ToDictionary(g => g.Key, g => g.ToList());
        var dupIds = liveStarts.Where(kv => kv.Value.Count > 1).Select(kv => kv.Key)
            .Union(liveEnds.Where(kv => kv.Value.Count > 1).Select(kv => kv.Key)).ToList();
        if (dupIds.Count > 0)
        {
            int next = GlobalMaxBookmarkId(main) + 1;
            foreach (var id in dupIds)
            {
                var ss = liveStarts.TryGetValue(id, out var sl) ? sl : new List<XElement>();
                var es = liveEnds.TryGetValue(id, out var el) ? el : new List<XElement>();
                int copies = Math.Max(ss.Count, es.Count);
                for (int k = 1; k < copies; k++)
                {
                    string fresh = (next++).ToString();
                    if (k < ss.Count) ss[k].SetAttributeValue(W.id, fresh);
                    if (k < es.Count) es[k].SetAttributeValue(W.id, fresh);
                    changed = true;
                }
            }
        }

        // (C) Reconcile pairing — GUARANTEE every run-level bookmarkStart has a matching bookmarkEnd and vice
        //     versa. A cross-paragraph range whose far endpoint lands in a churned span can be dropped by the
        //     token-diff renderer in dense, overlapping-bookmark layouts (the NVCA _DV_/_cp_text_ content-region
        //     bookmarks); a post-render guarantee keeps the output schema-sound and every cross-reference
        //     resolvable regardless of that edge. An orphaned START (its END was dropped) is re-closed with a
        //     synthetic zero-width end in the start's own revision context (so accept/reject keep it together);
        //     an orphaned END (its START — the name carrier — was dropped) is removed (nothing can reference a
        //     nameless marker). Faithful to the structure round-trip: the bookmark NAME survives and resolves.
        var startById = body.Descendants(W.bookmarkStart).Where(IsRunLevelBookmark)
            .GroupBy(s => IdOf(s) ?? "").ToDictionary(g => g.Key, g => g.ToList());
        var endById = body.Descendants(W.bookmarkEnd).Where(IsRunLevelBookmark)
            .GroupBy(e => IdOf(e) ?? "").ToDictionary(g => g.Key, g => g.ToList());
        foreach (var (id, sl) in startById)
        {
            int have = endById.TryGetValue(id, out var el) ? el.Count : 0;
            for (int k = have; k < sl.Count; k++)
            {
                sl[k].AddAfterSelf(new XElement(W.bookmarkEnd, new XAttribute(W.id, id)));
                changed = true;
            }
        }
        foreach (var (id, el) in endById)
        {
            int have = startById.TryGetValue(id, out var sl) ? sl.Count : 0;
            for (int k = have; k < el.Count; k++)
            {
                RemoveBookmarkMarker(el[k]);
                changed = true;
            }
        }

        if (changed)
            main.PutXDocument();
    }

    /// <summary>The set of body bookmark NAMES in a source IR document — scanned from each block's provenance
    /// source element (bookmarks are dropped from the IR model by rule N3 but retained on the source XML). Used
    /// to classify a rendered bookmark as common (in both sources) / inserted / deleted for round-trip-correct
    /// normalization.</summary>
    private static HashSet<string> BodyBookmarkNames(IrDocument? source)
    {
        var names = new HashSet<string>();
        if (source == null)
            return names;
        foreach (var block in source.Body.Blocks)
        {
            var el = block.Source.Element;
            if (el == null)
                continue;
            foreach (var bk in el.DescendantsAndSelf(W.bookmarkStart))
                if ((string?)bk.Attribute(W.name) is { } n && n.Length > 0)
                    names.Add(n);
        }
        return names;
    }

    /// <summary>True iff the marker's nearest enclosing <c>w:p</c> has a TRACKED paragraph mark
    /// (<c>w:pPr/w:rPr/w:del</c> or <c>w:ins</c>) — i.e. it sits in a whole-block-bailed deleted/inserted
    /// paragraph, where a bare marker could not survive both accept and reject.</summary>
    private static bool IsInWholeBlockRevisedParagraph(XElement marker)
    {
        var p = marker.Ancestors(W.p).FirstOrDefault();
        return p != null && IsWholeBlockRevisedParagraph(p);
    }

    /// <summary>True for a paragraph emitted as one complete insertion/deletion (or whole move) rather than a
    /// fine token-level revision. Its field plumbing already has the same revision context as its carrier and
    /// must never be lifted bare by <see cref="NormalizeFields"/> — an instruction-only field has no result run
    /// from which that normalizer could infer its context.</summary>
    private static bool IsWholeBlockRevisedParagraph(XElement paragraph)
    {
        var mark = paragraph.Element(W.pPr)?.Element(W.rPr);
        return mark != null && (mark.Element(W.del) != null || mark.Element(W.ins) != null);
    }

    /// <summary>True iff the bookmark marker sits at the paragraph/body run level — its ancestors up to the
    /// enclosing <c>w:p</c>/<c>w:body</c> are only run-level wrappers (ins/del/hyperlink/sdt/smartTag/fldSimple).
    /// A bookmark reached through anything else (a math <c>m:oMath</c>, a <c>w:drawing</c>, a textbox) is part of
    /// that opaque element's content hash and must NOT be renumbered/removed by the normalizer.</summary>
    private static bool IsRunLevelBookmark(XElement marker)
    {
        for (var a = marker.Parent; a != null; a = a.Parent)
        {
            if (a.Name == W.p || a.Name == W.body || a.Name == W.tc)
                return true;
            if (a.Name != W.ins && a.Name != W.del && a.Name != W.hyperlink &&
                a.Name != W.sdt && a.Name != W.sdtContent && a.Name != W.smartTag && a.Name != W.fldSimple)
                return false;
        }
        return false;
    }

    /// <summary>Remove a bookmark marker, dropping the empty <c>w:ins</c>/<c>w:del</c> wrapper that held only it.
    /// No-op if the marker (or its wrapper) was already detached by an earlier reconciliation step.</summary>
    private static void RemoveBookmarkMarker(XElement marker)
    {
        var parent = marker.Parent;
        if (parent == null)
            return;
        marker.Remove();
        if ((parent.Name == W.ins || parent.Name == W.del) && parent.Parent != null && !parent.Elements().Any())
            parent.Remove();
    }

    /// <summary>If the marker is the sole child of a <c>w:ins</c>/<c>w:del</c> wrapper, lift it OUT to a bare
    /// sibling (a start before the wrapper, an end after) so it survives BOTH accept and reject while keeping the
    /// bookmark wrapping the run content. Returns true if it moved.</summary>
    private static bool LiftBookmarkBare(XElement marker)
    {
        var parent = marker.Parent;
        if (parent == null || (parent.Name != W.ins && parent.Name != W.del) || parent.Parent == null)
            return false;
        marker.Remove();
        if (marker.Name == W.bookmarkEnd) parent.AddAfterSelf(marker);
        else parent.AddBeforeSelf(marker);
        if (!parent.Elements().Any())
            parent.Remove();
        return true;
    }

    /// <summary>The largest integer bookmark id present in ANY part of the document (body + headers/footers +
    /// note parts), so a freshly allocated id collides with no existing bookmark anywhere.</summary>
    private static int GlobalMaxBookmarkId(MainDocumentPart main)
    {
        int max = 0;
        void Scan(XElement? root)
        {
            if (root == null) return;
            foreach (var m in root.Descendants().Where(e => e.Name == W.bookmarkStart || e.Name == W.bookmarkEnd))
                if (int.TryParse((string?)m.Attribute(W.id), out var v) && v > max)
                    max = v;
        }
        Scan(main.GetXDocument().Root);
        foreach (var h in main.HeaderParts) Scan(h.GetXDocument().Root);
        foreach (var f in main.FooterParts) Scan(f.GetXDocument().Root);
        if (main.FootnotesPart != null) Scan(main.FootnotesPart.GetXDocument().Root);
        if (main.EndnotesPart != null) Scan(main.EndnotesPart.GetXDocument().Root);
        return max;
    }

    // ----------------------------------------------------------------- field-context normalization

    /// <summary>
    /// Re-home each fldChar field's plumbing (the <c>w:r</c>s carrying <c>w:fldChar</c>/<c>w:instrText</c>) to the
    /// field's OWN revision context. Field plumbing is kept across edit boundaries (AlwaysKeep) so a field is
    /// never orphaned, but a boundary can leave a begin/separate/end or instruction run in the wrong wrapper —
    /// e.g. wrapped in <c>w:ins</c> when the text BEFORE the field was edited, which would make the field vanish
    /// on reject. The field's home context is read from its RESULT runs: an unchanged field (result bare) or a
    /// result-EDITED field (result split del/ins) → all plumbing BARE (survives accept AND reject); a wholly
    /// deleted/inserted field keeps its plumbing in <c>w:del</c>/<c>w:ins</c> (toggles with the field).
    /// </summary>
    private static void NormalizeFields(MainDocumentPart main)
    {
        // The body is not the only renderer that emits fine-grained field revisions: note and header/footer
        // scope diffs share the same token renderer. Normalize each live story root after all scopes have been
        // rebuilt, otherwise a field boundary in a header or footnote can still vanish on one revision view.
        var parts = new List<OpenXmlPart> { main };
        parts.AddRange(main.HeaderParts);
        parts.AddRange(main.FooterParts);
        if (main.FootnotesPart is not null)
            parts.Add(main.FootnotesPart);
        if (main.EndnotesPart is not null)
            parts.Add(main.EndnotesPart);

        foreach (var part in parts.Distinct())
        {
            var root = part.GetXDocument().Root;
            if (root is not null && NormalizeFieldsInRoot(root))
                part.PutXDocument();
        }
    }

    private static bool NormalizeFieldsInRoot(XElement root)
    {
        static bool IsBareRun(XElement r) => !r.Ancestors().Any(a => a.Name == W.ins || a.Name == W.del);

        bool changed = false;
        foreach (var p in root.Descendants(W.p))
        {
            // A whole paired paragraph is already a complete w:del/w:ins carrier. Do not infer field context
            // from its result runs: instruction-only and empty-result fields have none, and lifting their begin /
            // instruction / end plumbing bare would make BOTH codes survive Accept and Reject.
            if (IsWholeBlockRevisedParagraph(p))
                continue;

            // Walk this paragraph's runs in document order, grouping each top-level fldChar field (begin..end)
            // into its plumbing runs (fldChar / instrText) and result runs (the visible display between separate
            // and end). Nested fields fold into the enclosing one (their plumbing is still field plumbing).
            var fields = new List<(List<XElement> Plumbing, List<XElement> Result)>();
            (List<XElement> Plumbing, List<XElement> Result)? cur = null;
            int depth = 0;
            bool afterSeparate = false;
            foreach (var r in p.Descendants(W.r))
            {
                var fc = r.Element(W.fldChar);
                var ty = fc != null ? (string?)fc.Attribute(W.fldCharType) : null;
                bool hasInstr = r.Element(W.instrText) != null || r.Element(W.delInstrText) != null;
                if (ty == "begin")
                {
                    if (depth == 0) { cur = (new(), new()); fields.Add(cur.Value); afterSeparate = false; }
                    depth++;
                    cur?.Plumbing.Add(r);
                }
                else if (ty == "separate")
                {
                    if (depth >= 1) { cur?.Plumbing.Add(r); if (depth == 1) afterSeparate = true; }
                }
                else if (ty == "end")
                {
                    if (depth >= 1) cur?.Plumbing.Add(r);
                    depth--;
                    if (depth == 0) { cur = null; afterSeparate = false; }
                }
                else if (depth >= 1)
                {
                    if (hasInstr || !afterSeparate) cur?.Plumbing.Add(r);   // instruction / pre-separate plumbing
                    else cur?.Result.Add(r);                                // the visible field result
                }
            }

            foreach (var (plumbing, result) in fields)
            {
                // Home context from the result: del-only → del, ins-only → ins, else (bare or result-edited) → bare.
                bool anyResultDel = result.Any(r => r.Ancestors(W.del).Any());
                bool anyResultIns = result.Any(r => r.Ancestors(W.ins).Any());
                bool anyResultBare = result.Any(IsBareRun);
                bool toBare = anyResultBare || !(anyResultDel ^ anyResultIns); // bare, no result, or mixed → bare
                if (!toBare)
                    continue; // wholly deleted/inserted field: plumbing already toggles with it
                foreach (var pr in plumbing)
                    if (!IsBareRun(pr) && LiftRunBare(pr))
                        changed = true;
            }
        }

        return changed;
    }

    /// <summary>Lift a run out of its sole-child <c>w:ins</c>/<c>w:del</c> wrapper to bare. A run lifted out of a
    /// deletion is no longer deleted, so its <c>w:delText</c>/<c>w:delInstrText</c> revert to
    /// <c>w:t</c>/<c>w:instrText</c>. Returns true if it moved.</summary>
    private static bool LiftRunBare(XElement run)
    {
        var w = run.Parent;
        if (w == null || (w.Name != W.ins && w.Name != W.del) || w.Parent == null)
            return false;
        if (w.Name == W.del)
        {
            foreach (var dt in run.DescendantsAndSelf(W.delText).ToList()) dt.Name = W.t;
            foreach (var di in run.DescendantsAndSelf(W.delInstrText).ToList()) di.Name = W.instrText;
        }
        run.Remove();
        w.AddBeforeSelf(run);
        if (!w.Elements().Any())
            w.Remove();
        return true;
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
        string? leftAnchor = state.MoveSourceAnchor(gid);
        var leftMovedPara = SourceElement(leftAnchor, state.Left);

        if (!op.RequiresWholeParagraphReplace &&
            op.Kind == IrEditOpKind.MoveModifyBlock && op.TokenDiff is { } tokenDiff &&
            op.TextboxDiffs is null && leftMovedPara?.Name == W.p && leftAnchor is not null)
        {
            // Build the destination paragraph from the token diff, like RenderModifiedParagraph, but with the
            // moved-and-equal spans wrapped in w:moveTo (instead of left unwrapped) so the whole relocated
            // content vanishes on reject and appears on accept. Insert spans → w:ins, Delete spans → w:del.
            var para = BuildMoveModifyDestination(op, tokenDiff, leftAnchor, state);
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
        // A moved paragraph whose pPr also changed tracks the property change at the DESTINATION
        // (Word's own shape: moveTo content + pPrChange). Source/reject keeps the left paragraph verbatim.
        if (leftMovedPara?.Name == W.p)
            ApplyBlockFormatChanges(dest, leftMovedPara, src, state);
        MarkWholeParagraphAs(dest, RevKind.MoveTo, state);
        BracketParagraphWithMoveRange(dest, isFrom: false, moveName, state);
        sink.Add(dest);
    }

    /// <summary>Build a MoveModify destination paragraph from its token diff: Equal/FormatChanged spans →
    /// <c>w:moveTo</c> (moved-and-unchanged), Insert spans → <c>w:ins</c>, Delete spans → <c>w:del</c>. Returns
    /// null if the source elements are unexpectedly missing (caller falls back to a plain whole-paragraph
    /// moveTo).</summary>
    private static XElement? BuildMoveModifyDestination(
        IrEditOp op, IrTokenDiff tokenDiff, string leftAnchor, RenderState state)
    {
        var leftPara = SourceElement(leftAnchor, state.Left);
        var rightPara = SourceElement(op.RightAnchor, state.RightSource);
        if (leftPara == null || rightPara == null)
            return null;

        var leftRuns = new SourceRunModel(leftPara);
        var rightRuns = new SourceRunModel(rightPara);
        var leftTokens = ParagraphTokens(leftAnchor, state.Left, state.Settings);
        var rightTokens = ParagraphTokens(op.RightAnchor, state.RightSource, state.Settings);

        var newPara = new XElement(W.p);
        var rightPPr = rightPara.Element(W.pPr);
        if (rightPPr != null)
            newPara.Add(StripUnids(new XElement(rightPPr)));
        ApplyBlockFormatChanges(newPara, leftPara, rightPara, state);

        // The regular fine renderer already knows how to produce both FormatChanged rPrChange markers and
        // the field-safe insert/delete projection. Its optional stable-span wrapper gives this move destination
        // the one additional shape it needs: unchanged and format-changed right spans live in w:moveTo.
        newPara.Add(BuildTokenOpContent(tokenDiff, leftTokens, rightTokens, leftRuns, rightRuns, state,
            stableSpanKind: RevKind.MoveTo));
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
        var wrapped = runChildren.SelectMany(c => WrapFieldAware(c, kind, state)).ToList();
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
        // Word-parity preservation: an EQUAL block emits the ORIGINAL right element(s) — pre-existing
        // input revision markup intact — when PreserveInputRevisions mapped a group; otherwise the
        // accepted working element exactly as before. A multi-member group carries the mark-deleted
        // paragraphs the document-level accept merged away; they vanish again on accept, so the accept
        // round-trip is unchanged. (The map only ever holds right-body elements, so left-side and
        // composite lookups are no-ops.)
        if (state.PreservedGroup(src) is { } group)
        {
            foreach (var member in group)
            {
                var preserved = NormalizePreservedClone(new XElement(member), state);
                if (fromRight)
                    state.RegisterMediaReferences(preserved);
                sink.Add(StripUnids(preserved));
            }
            return;
        }
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
        string? anchor,
        IrDocument doc,
        RenderState state,
        List<XElement> sink,
        RevKind kind,
        bool fromRight,
        bool projectLeftDeletionAsInsertion = false)
    {
        var src = SourceElement(anchor, doc);
        if (src == null)
            return;
        // Word-parity preservation: an INSERTED right-only block is built from the ORIGINAL right
        // element(s) when PreserveInputRevisions mapped a group, so the input's own w:ins/w:del markup
        // rides through — MarkWholeParagraph leaves foreign wrappers as-is (never nesting a same-kind
        // wrapper) and wraps only the plain runs as this diff's insertion; MarkParagraphMark keeps a
        // foreign mark marker (a multi-member group's mark-deleted members vanish again on accept, so
        // the accept round-trip is unchanged).
        var group = fromRight
            ? state.PreservedGroup(src)
            : IsDeleteGrade(kind) ? state.ProjectableLeftDeletionGroup(src) : null;
        bool projectsLeftDeletionAsInsertion = false;
        if (group == null && fromRight && projectLeftDeletionAsInsertion && kind == RevKind.Ins &&
            state.ProjectableLeftInsertionOriginal(src) is { } leftDeletion)
        {
            group = new List<XElement> { leftDeletion };
            projectsLeftDeletionAsInsertion = true;
        }
        if (group != null)
        {
            foreach (var member in group)
            {
                var preserved = NormalizePreservedClone(new XElement(member), state);
                if (!fromRight)
                    ProjectLeftInsertionsAsDeletions(preserved);
                else if (projectsLeftDeletionAsInsertion)
                    ProjectLeftDeletionsAsInsertions(preserved);

                // The reverse projection clones a raw LEFT source. Its media relationships already belong to
                // the output package (which is a clone of LEFT), so it must not enter the RIGHT import path.
                EmitOneWholeBlock(preserved, state, sink, kind,
                    fromRight: fromRight && !projectsLeftDeletionAsInsertion);
            }
            return;
        }
        EmitOneWholeBlock(new XElement(src), state, sink, kind, fromRight);
    }

    /// <summary>The single-block tail of <see cref="EmitWholeBlock"/>: strip engine bookkeeping, register
    /// media, revision-mark per block kind, and emit.</summary>
    private static void EmitOneWholeBlock(XElement clone, RenderState state, List<XElement> sink, RevKind kind, bool fromRight)
    {
        StripUnids(clone);
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
    /// Emit one whole block-level content control.  OOXML does not permit a bare <c>w:ins</c>/<c>w:del</c>
    /// around the control envelope: its property/binding shell would then survive the opposite revision view.
    /// Word represents an inserted/deleted control with two paired custom-XML range boundaries: one crosses the
    /// opening <c>w:sdt</c> tag and one crosses its closing tag.  The range layer toggles the wrapper; the
    /// recursively marked payload toggles the content after a delete-range accept collapses that wrapper.
    /// </summary>
    private static void EmitWholeSdt(
        string? anchor, IrDocument doc, RenderState state, List<XElement> sink, RevKind kind, bool fromRight)
    {
        var src = SourceElement(anchor, doc);
        if (src == null || src.Name != W.sdt || src.Element(W.sdtContent) is null)
            return;

        var sdt = StripUnids(new XElement(src));
        if (fromRight)
            state.RegisterMediaReferences(sdt);

        var boundaries = MarkWholeSdtEnvelope(sdt, kind, state);
        sink.Add(boundaries.Before);
        sink.Add(sdt);
        sink.Add(boundaries.After);
    }

    /// <summary>
    /// Mark a control envelope and all of its block payload.  For a nested control the returned boundaries are
    /// inserted as siblings in its parent's <c>w:sdtContent</c>; for the outer control they are emitted by
    /// <see cref="EmitWholeSdt"/> around the cloned <c>w:sdt</c>.  The two range IDs are intentionally
    /// distinct: the first captures the SDT's start tag and the second its end tag, which is the pairing shape
    /// <see cref="RevisionProcessor"/> recognizes when accepting a deleted control.
    /// </summary>
    private static (XElement Before, XElement After) MarkWholeSdtEnvelope(XElement sdt, RevKind kind, RenderState state)
    {
        var content = sdt.Element(W.sdtContent)!;
        foreach (var child in content.Elements().ToList())
            MarkWholeSdtContentChild(child, kind, state);

        var startName = IsDeleteGrade(kind) ? W.customXmlDelRangeStart : W.customXmlInsRangeStart;
        var endName = IsDeleteGrade(kind) ? W.customXmlDelRangeEnd : W.customXmlInsRangeEnd;

        // Range A begins immediately before the control and ends as the first sdtContent child, so it contains
        // the opening tag. Range B begins as the last sdtContent child and ends immediately after the control,
        // so it contains the closing tag. AcceptDeletedAndMovedFromContentControls intersects those two sets.
        var before = new XElement(startName, state.RevisionAttributes());
        var beforeId = (string?)before.Attribute(W.id) ?? "";
        content.AddFirst(new XElement(endName, new XAttribute(W.id, beforeId)));

        var afterStart = new XElement(startName, state.RevisionAttributes());
        var afterId = (string?)afterStart.Attribute(W.id) ?? "";
        content.Add(afterStart);
        var after = new XElement(endName, new XAttribute(W.id, afterId));
        return (before, after);
    }

    /// <summary>
    /// Mark payload carried by a block-level SDT without losing legal intermediate containers.  Paragraphs and
    /// tables own the established whole-block revision shapes; nested SDTs receive their own envelope ranges.
    /// Inline SDT payloads can be direct runs, so they are wrapped at run granularity too. Any other transparent
    /// block container is descended so customXml/smartTag wrappers do not leave visible paragraphs unmarked after
    /// an enclosing delete-range acceptance.
    /// </summary>
    private static void MarkWholeSdtContentChild(XElement child, RevKind kind, RenderState state)
    {
        if (child.Name == W.p)
        {
            MarkWholeParagraph(child, kind, state);
            return;
        }
        if (child.Name == W.tbl)
        {
            MarkWholeTable(child, kind, state);
            return;
        }
        if (child.Name == W.sdt && child.Element(W.sdtContent) is not null)
        {
            var boundaries = MarkWholeSdtEnvelope(child, kind, state);
            child.AddBeforeSelf(boundaries.Before);
            child.AddAfterSelf(boundaries.After);
            return;
        }

        if (child.Name == W.r || child.Name == W.hyperlink || child.Name == W.smartTag)
        {
            var replacement = WrapFieldAware(child, kind, state).Cast<object>().ToArray();
            child.ReplaceWith(replacement);
            return;
        }

        // Content controls can legally carry transparent block containers such as w:customXml. Preserve their
        // shell verbatim but descend into their block-bearing children; raw leaf metadata has no visible payload.
        foreach (var nested in child.Elements().ToList())
            MarkWholeSdtContentChild(nested, kind, state);
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
            // PreserveInputRevisions: a preserved ORIGINAL row that already carries a foreign row marker
            // (Arthur's inserted/deleted row) keeps it — same rule as MarkParagraphMark. Unreachable when
            // the flag is off (accept-view sources carry none).
            bool foreignRowMark = state.Settings.PreserveInputRevisions &&
                trPr.Elements().Any(e => e.Name == W.ins || e.Name == W.del);
            if (!foreignRowMark)
            {
                // In w:trPr the row-revision markers w:ins/w:del come at the END of the property order (after
                // cnfStyle/trHeight/cantSplit/…, before only w:trPrChange) — so APPEND, never AddFirst, or a
                // following w:trHeight becomes schema-invalid.
                trPr.Elements().Where(e => e.Name == W.ins || e.Name == W.del).Remove();
                trPr.Add(new XElement(kind == RevKind.Ins ? W.ins : W.del, state.RevisionAttributes()));
            }

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
        {
            // A content-control wrapper has native custom-XML range revisions that make the envelope itself
            // reversible. Ordinary w:del/w:ins can only toggle its runs, leaving an empty w:sdt behind in the
            // opposite view. Mark the full envelope here, then retain its range boundaries as paragraph-level
            // siblings while the payload receives run revisions inside MarkWholeSdtEnvelope.
            if (child.Name == W.sdt && child.Element(W.sdtContent) is not null)
            {
                var boundaries = MarkWholeSdtEnvelope(child, kind, state);
                wrapped.Add(boundaries.Before);
                wrapped.Add(child);
                wrapped.Add(boundaries.After);
                continue;
            }
            // PreserveInputRevisions: a preserved ORIGINAL block may carry foreign revision wrappers as
            // paragraph children. Leave them exactly as-is — their content is already revision-marked
            // (foreign ins stays inserted, foreign del/moveFrom stays deleted-grade), and re-wrapping
            // would nest same-kind wrappers. Unreachable when the flag is off: sources are accept-view.
            if (state.Settings.PreserveInputRevisions && PreservedWrapperNames.Contains(child.Name))
                wrapped.Add(child);
            else
                wrapped.AddRange(WrapFieldAware(child, kind, state));
        }

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
        // (A w:fldSimple is handled upstream by ExpandFieldForRevision — it is EXPANDED to fldChar runs before
        // reaching here, since w:fldSimple is not in w:del's content model and keeping its wrapper bare would
        // leave a dangling empty field on accept/reject.)
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
    private static IEnumerable<XElement> WrapContainerChild(XElement child, RevKind kind, RenderState state)
    {
        // A nested fldSimple is just as illegal inside w:ins/w:del as a top-level one. Expand it before
        // wrapping, preserving its result position inside the enclosing hyperlink/SDT/smartTag container.
        if (child.Name == W.fldSimple)
            return WrapFieldAware(child, kind, state);
        if (child.Name == W.r || child.Name == W.hyperlink || child.Name == W.smartTag || child.Name == W.sdt)
            return new[] { WrapRunLevel(child, kind, state) };
        if (child.Name == W.sdtContent)
        {
            var content = new XElement(child.Name, child.Attributes());
            foreach (var inner in child.Elements())
                content.Add(WrapContainerChild(inner, kind, state));
            return new[] { content };
        }
        return new[] { new XElement(child) };
    }

    /// <summary>Mark a paragraph's end-of-paragraph mark inserted/deleted: an EMPTY <c>w:ins</c>/<c>w:del</c>
    /// inside <c>w:pPr/w:rPr</c> (the encoding <see cref="RevisionProcessor"/> recognizes — accept of a
    /// deleted mark merges the paragraph with the following one; reject restores it). The paragraph mark
    /// supports only <c>w:ins</c>/<c>w:del</c>, so a move FROM marks the mark deleted (del-grade) and a move
    /// TO marks it inserted (ins-grade).</summary>
    private static void MarkParagraphMark(XElement para, RevKind kind, RenderState state)
    {
        // PreserveInputRevisions: a preserved ORIGINAL paragraph whose mark already carries a revision
        // marker (a foreign mark-insertion or mark-deletion) keeps it — the input's own paragraph-structure
        // revision is the Word-parity fact, and stamping ours would re-attribute (or resurrect) it. A
        // foreign mark-DELETED paragraph still vanishes on accept, so the accept round-trip is unchanged.
        // Unreachable when the flag is off: accept-view sources carry no mark markers.
        if (state.Settings.PreserveInputRevisions &&
            para.Element(W.pPr)?.Element(W.rPr) is { } presentRPr &&
            presentRPr.Elements().Any(e =>
                e.Name == W.ins || e.Name == W.del || e.Name == W.moveFrom || e.Name == W.moveTo))
            return;

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
        {
            var stamped = StripUnids(new XElement(rightPPr));
            DropUnresolvableStyleRef(stamped, state);
            newPara.Add(stamped);
        }
        ApplyBlockFormatChanges(newPara, leftPara, rightPara, state);

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

    /// <summary>
    /// Render a text-equal, format-differing TABLE pair (block-format-change family): emit the RIGHT table
    /// verbatim (its content equals the left's — this is a FormatOnly pair) and stamp the native table-shell
    /// property-revision markers (<c>w:tblPrChange</c>/<c>w:tblGridChange</c>/<c>w:trPrChange</c>/<c>w:tcPrChange</c>)
    /// wherever a shell differs. Rows and cells pair positionally — safe because ContentHash-equality guarantees
    /// an identical row/cell structure. Accept ≡ right shells; reject ≡ left shells.
    /// </summary>
    private static void EmitFormatOnlyTable(IrEditOp op, RenderState state, List<XElement> sink)
    {
        var rightTbl = SourceElement(op.RightAnchor, state.RightSource);
        var leftTbl = SourceElement(op.LeftAnchor, state.Left);
        if (rightTbl == null || leftTbl == null || rightTbl.Name != W.tbl || leftTbl.Name != W.tbl)
        {
            EmitVerbatim(op.RightAnchor, state.RightSource, state, sink, fromRight: true);
            return;
        }

        var newTbl = StripUnids(new XElement(rightTbl));
        state.RegisterMediaReferences(newTbl);
        ApplyTableLevelShellChanges(newTbl, leftTbl, state);

        // FormatOnly ⇒ ContentHash-equal ⇒ identical row/cell structure, so rows pair positionally.
        var leftRows = leftTbl.Elements(W.tr).ToList();
        var newRows = newTbl.Elements(W.tr).ToList();
        for (int i = 0; i < newRows.Count && i < leftRows.Count; i++)
            ApplyRowAndCellShellChanges(newRows[i], leftRows[i], state);

        sink.Add(newTbl);
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
        // its run-level content from the spans. A right pStyle absent from the LEFT style universe is
        // dropped — Word expresses a paired paragraph's format change in direct props only.
        var newPara = new XElement(W.p);
        var rightPPr = rightPara.Element(W.pPr);
        if (rightPPr != null)
        {
            var stamped = StripUnids(new XElement(rightPPr));
            DropUnresolvableStyleRef(stamped, state);
            newPara.Add(stamped);
        }
        ApplyBlockFormatChanges(newPara, leftPara, rightPara, state);

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
    /// <summary>
    /// Re-anchor a narrowly provable ambiguous whitespace match, then rearrange a paragraph's token ops into
    /// Word's replace-region grammar (the token-level analogue
    /// of <see cref="RenderBlockOpsWordShaped"/>): a maximal run of Insert/Delete ops separated only by
    /// WHITESPACE Equal ops renders as ONE inserted region (all right-side tokens, interior whitespace
    /// included) followed by ONE deleted region (all left-side tokens likewise) — Word writes
    /// "Heading <ins>2 Center</ins><del>1 Style</del> Demo", never per-word del/ins alternation.
    /// Interior whitespace Equals are CONVERTED to an Insert op over their right span plus a Delete op
    /// over their left span, so accept still yields exactly the right bytes and reject the left bytes
    /// (each side's raw text is emitted from its own source; per-side span tiling is preserved — the
    /// converted ops carry an empty span on the opposite side). Pure-insert or pure-delete runs, and
    /// any Equal that is non-whitespace or zero-width passes through untouched. A format-changing separator
    /// normally acts as interior glue too. Before that grouping, however, a bounded normalization repairs the
    /// one ambiguous shape in which Myers matched the WRONG repeated whitespace token: an isolated space is
    /// paired before left- or right-only text even though an identical space immediately precedes the next
    /// retained word on both sides. The normalizer moves that match to the retained-word boundary, so it becomes
    /// a format change instead of a synthetic insertion/deletion. The caller additionally restricts this to
    /// paragraphs made solely of direct text runs, so a transparent field/container boundary cannot be moved.
    /// This is intentionally render-only: it never changes the edit script, token differ, or revision-list
    /// projection.
    /// </summary>
    private static List<IrTokenOp> CoalesceTokenOpsWordShaped(
        IReadOnlyList<IrTokenOp> ops,
        IReadOnlyList<IrDiffToken> leftTokens, IReadOnlyList<IrDiffToken> rightTokens,
        IrFormatComparison formatComparison,
        SourceRunModel leftRuns, SourceRunModel rightRuns)
    {
        // Re-pairing an otherwise transparent token match is safe only when both source paragraphs
        // are ordinary direct text runs. Fields, content controls, hyperlinks, revision containers,
        // and zero-width inline plumbing are deliberately transparent to the tokenizer, but their
        // source ownership cannot be reconstructed from token provenance after moving a boundary.
        // Keep those paragraphs on the normal projection path.
        IReadOnlyList<IrTokenOp> projectedOps =
            leftRuns.SupportsWhitespaceReanchoring && rightRuns.SupportsWhitespaceReanchoring
                ? ReanchorAmbiguousWhitespaceMatches(ops, leftTokens, rightTokens, formatComparison)
                : ops;

        // Interior separator: an Equal OR FormatChanged span whose raw text is pure whitespace on
        // both sides. FormatChanged qualifies because when the two sides' run formats differ (a
        // formatting-change rewrite), every separator space between replaced words is a
        // FormatChanged span, not Equal — excluding it re-fragments the replacement into the
        // word-by-word zip this pass exists to remove. The conversion emits each side's own bytes
        // AND formatting (right space as w:ins, left space as w:del), so it is exact by construction.
        bool IsInteriorWhitespaceSeparator(IrTokenOp op)
        {
            if (op.Kind != IrTokenOpKind.Equal && op.Kind != IrTokenOpKind.FormatChanged)
                return false;
            var l = RawSpanText(leftTokens, op.LeftStart, op.LeftEnd);
            var r = RawSpanText(rightTokens, op.RightStart, op.RightEnd);
            return l.Length > 0 && r.Length > 0 &&
                   string.IsNullOrWhiteSpace(l) && string.IsNullOrWhiteSpace(r);
        }

        var result = new List<IrTokenOp>(projectedOps.Count);
        int i = 0;
        while (i < projectedOps.Count)
        {
            if (projectedOps[i].Kind != IrTokenOpKind.Insert && projectedOps[i].Kind != IrTokenOpKind.Delete)
            {
                result.Add(projectedOps[i]);
                i++;
                continue;
            }
            // Grow the window over changed ops and interior whitespace Equals; it must END changed.
            int lastChanged = i;
            int j = i + 1;
            while (j < projectedOps.Count)
            {
                if (projectedOps[j].Kind is IrTokenOpKind.Insert or IrTokenOpKind.Delete)
                {
                    lastChanged = j;
                    j++;
                }
                else if (IsInteriorWhitespaceSeparator(projectedOps[j]))
                {
                    j++;   // tentatively interior; trimmed below if nothing changed follows
                }
                else
                {
                    break;
                }
            }
            var group = new List<IrTokenOp>();
            for (int k = i; k <= lastChanged; k++)
                group.Add(projectedOps[k]);
            i = lastChanged + 1;

            var hasIns = group.Any(g => g.Kind == IrTokenOpKind.Insert);
            var hasDel = group.Any(g => g.Kind == IrTokenOpKind.Delete);
            if (!hasIns || !hasDel)
            {
                result.AddRange(group);
                continue;
            }
            foreach (var g in group)
            {
                if (g.Kind == IrTokenOpKind.Insert)
                    result.Add(g);
                else if (g.Kind is IrTokenOpKind.Equal or IrTokenOpKind.FormatChanged)
                    result.Add(new IrTokenOp(IrTokenOpKind.Insert, g.LeftStart, g.LeftStart, g.RightStart, g.RightEnd));
            }
            foreach (var g in group)
            {
                if (g.Kind == IrTokenOpKind.Delete)
                    result.Add(g);
                else if (g.Kind is IrTokenOpKind.Equal or IrTokenOpKind.FormatChanged)
                    result.Add(new IrTokenOp(IrTokenOpKind.Delete, g.LeftStart, g.LeftEnd, g.RightStart, g.RightStart));
            }
        }
        return result;
    }

    /// <summary>
    /// Normalizes the one whitespace-alignment ambiguity which is visible in Word-compare output.  The tokenizer
    /// intentionally represents every separator as its own token; when a replacement has no shared prefix,
    /// Myers may therefore match its first repeated space instead of the space immediately before a retained
    /// suffix word.  For example, it can pair the first spaces in
    /// <c>"Subtitle Style Demo"</c> / <c>"Superscript Demo"</c>, leaving the second left space deleted.  Word
    /// instead retains <c>" Demo"</c>.  That difference matters when formatting changes: grouping the early
    /// <c>FormatChanged</c> space below otherwise creates a visible inserted+deleted space pair.
    ///
    /// The proof gate is deliberately strict: exactly one ASCII-space separator must be matched;
    /// only unilateral delete (or mirror insert) ops may lie before the next raw-identical WORD anchor; and an
    /// identical separator must sit immediately before that anchor on the unilateral side.  Thus the operation
    /// re-tiles the same left/right token ranges without changing their text, and it cannot cross a hyperlink,
    /// NBSP-vs-space normalization, punctuation, or zero-width atom boundary.  This is a markup-projection
    /// normalization only; <see cref="IrTokenDiffer"/> and the persisted edit script remain byte-for-byte as
    /// produced by the diff engine.
    /// </summary>
    private static List<IrTokenOp> ReanchorAmbiguousWhitespaceMatches(
        IReadOnlyList<IrTokenOp> ops,
        IReadOnlyList<IrDiffToken> leftTokens, IReadOnlyList<IrDiffToken> rightTokens,
        IrFormatComparison formatComparison)
    {
        var normalized = new List<IrTokenOp>(ops.Count);
        for (int i = 0; i < ops.Count; i++)
        {
            if (TryReanchorWhitespaceToFollowingWordOnLeft(
                    ops, i, leftTokens, rightTokens, formatComparison,
                    out var consumedThrough, out var replacement) ||
                TryReanchorWhitespaceToFollowingWordOnRight(
                    ops, i, leftTokens, rightTokens, formatComparison,
                    out consumedThrough, out replacement))
            {
                normalized.AddRange(replacement);
                i = consumedThrough;
                continue;
            }

            normalized.Add(ops[i]);
        }
        return normalized;
    }

    /// <summary>
    /// Left-side form of the ambiguity: a whitespace match is followed only by deleted left tokens before the
    /// next retained word.  Re-pair the right whitespace with the identical left whitespace directly before
    /// that word, and fold the former left whitespace into the deletion range.
    /// </summary>
    private static bool TryReanchorWhitespaceToFollowingWordOnLeft(
        IReadOnlyList<IrTokenOp> ops, int matchIndex,
        IReadOnlyList<IrDiffToken> leftTokens, IReadOnlyList<IrDiffToken> rightTokens,
        IrFormatComparison formatComparison,
        out int consumedThrough, out List<IrTokenOp> replacement)
    {
        consumedThrough = -1;
        replacement = new List<IrTokenOp>();
        var match = ops[matchIndex];
        if (!IsSingleRawWhitespaceMatch(match, leftTokens, rightTokens))
            return false;

        int leftCursor = match.LeftEnd;
        int rightCursor = match.RightEnd;
        int anchorIndex = matchIndex + 1;
        while (anchorIndex < ops.Count && ops[anchorIndex].Kind == IrTokenOpKind.Delete)
        {
            var deleted = ops[anchorIndex];
            if (deleted.LeftStart != leftCursor || deleted.RightStart != rightCursor ||
                deleted.RightEnd != rightCursor)
                return false;
            leftCursor = deleted.LeftEnd;
            anchorIndex++;
        }

        if (anchorIndex == matchIndex + 1 || anchorIndex >= ops.Count)
            return false;
        var anchor = ops[anchorIndex];
        if (anchor.LeftStart != leftCursor || anchor.RightStart != rightCursor ||
            !IsRawEqualWordAnchor(anchor, leftTokens, rightTokens))
            return false;

        int retainedLeftSpace = anchor.LeftStart - 1;
        if (retainedLeftSpace < match.LeftEnd ||
            !SameRawWhitespaceToken(leftTokens[retainedLeftSpace], rightTokens[match.RightStart]))
            return false;

        // The old left match plus its intervening deleted content become one deletion; the whitespace directly
        // before the retained suffix becomes a freshly classified Equal/FormatChanged span.
        replacement.Add(new IrTokenOp(IrTokenOpKind.Delete,
            match.LeftStart, retainedLeftSpace, match.RightStart, match.RightStart));
        AppendReclassifiedWhitespaceAndAnchor(replacement,
            retainedLeftSpace, match.RightStart, anchor,
            leftTokens, rightTokens, formatComparison);
        consumedThrough = anchorIndex;
        return true;
    }

    /// <summary>
    /// Mirror of <see cref="TryReanchorWhitespaceToFollowingWordOnLeft"/>: a whitespace match is followed only
    /// by inserted right tokens before the next retained word.  Re-pair the left whitespace with the identical
    /// right whitespace directly before that word and fold the former right whitespace into the insertion.
    /// </summary>
    private static bool TryReanchorWhitespaceToFollowingWordOnRight(
        IReadOnlyList<IrTokenOp> ops, int matchIndex,
        IReadOnlyList<IrDiffToken> leftTokens, IReadOnlyList<IrDiffToken> rightTokens,
        IrFormatComparison formatComparison,
        out int consumedThrough, out List<IrTokenOp> replacement)
    {
        consumedThrough = -1;
        replacement = new List<IrTokenOp>();
        var match = ops[matchIndex];
        if (!IsSingleRawWhitespaceMatch(match, leftTokens, rightTokens))
            return false;

        int leftCursor = match.LeftEnd;
        int rightCursor = match.RightEnd;
        int anchorIndex = matchIndex + 1;
        while (anchorIndex < ops.Count && ops[anchorIndex].Kind == IrTokenOpKind.Insert)
        {
            var inserted = ops[anchorIndex];
            if (inserted.LeftStart != leftCursor || inserted.LeftEnd != leftCursor ||
                inserted.RightStart != rightCursor)
                return false;
            rightCursor = inserted.RightEnd;
            anchorIndex++;
        }

        if (anchorIndex == matchIndex + 1 || anchorIndex >= ops.Count)
            return false;
        var anchor = ops[anchorIndex];
        if (anchor.LeftStart != leftCursor || anchor.RightStart != rightCursor ||
            !IsRawEqualWordAnchor(anchor, leftTokens, rightTokens))
            return false;

        int retainedRightSpace = anchor.RightStart - 1;
        if (retainedRightSpace < match.RightEnd ||
            !SameRawWhitespaceToken(leftTokens[match.LeftStart], rightTokens[retainedRightSpace]))
            return false;

        replacement.Add(new IrTokenOp(IrTokenOpKind.Insert,
            match.LeftStart, match.LeftStart, match.RightStart, retainedRightSpace));
        AppendReclassifiedWhitespaceAndAnchor(replacement,
            match.LeftStart, retainedRightSpace, anchor,
            leftTokens, rightTokens, formatComparison);
        consumedThrough = anchorIndex;
        return true;
    }

    /// <summary>Add the re-paired whitespace under the engine's normal format policy, then retain the original
    /// word anchor as a separate span.  Keeping the source boundary is essential: a re-paired whitespace token
    /// can have different old/new run properties from the following retained word, and a single rPrChange cannot
    /// represent both reject-side formats.</summary>
    private static void AppendReclassifiedWhitespaceAndAnchor(
        List<IrTokenOp> sink, int leftSpace, int rightSpace, IrTokenOp anchor,
        IReadOnlyList<IrDiffToken> leftTokens, IReadOnlyList<IrDiffToken> rightTokens,
        IrFormatComparison formatComparison)
    {
        var whitespaceKind = IrModeledFormat.RunFormatEqual(
            leftTokens[leftSpace].Format, rightTokens[rightSpace].Format, formatComparison)
            ? IrTokenOpKind.Equal
            : IrTokenOpKind.FormatChanged;
        var whitespace = new IrTokenOp(whitespaceKind,
            leftSpace, leftSpace + 1, rightSpace, rightSpace + 1);
        sink.Add(whitespace);

        sink.Add(anchor);
    }

    private static bool IsSingleRawWhitespaceMatch(
        IrTokenOp op, IReadOnlyList<IrDiffToken> leftTokens, IReadOnlyList<IrDiffToken> rightTokens) =>
        (op.Kind is IrTokenOpKind.Equal or IrTokenOpKind.FormatChanged) &&
        op.LeftEnd == op.LeftStart + 1 && op.RightEnd == op.RightStart + 1 &&
        SameRawWhitespaceToken(leftTokens[op.LeftStart], rightTokens[op.RightStart]);

    /// <summary>The following retained anchor must be a raw-identical lexical word, rather than a punctuation,
    /// normalized NBSP/case match, hyperlink-context mismatch, or zero-width atom.  Those cases stay on the
    /// original engine projection.</summary>
    private static bool IsRawEqualWordAnchor(
        IrTokenOp op, IReadOnlyList<IrDiffToken> leftTokens, IReadOnlyList<IrDiffToken> rightTokens)
    {
        if (op.Kind is not (IrTokenOpKind.Equal or IrTokenOpKind.FormatChanged) ||
            op.LeftStart >= op.LeftEnd || op.RightStart >= op.RightEnd)
            return false;
        var left = leftTokens[op.LeftStart];
        var right = rightTokens[op.RightStart];
        return left.Kind == IrDiffTokenKind.Word && right.Kind == IrDiffTokenKind.Word &&
               left.MatchKey == right.MatchKey &&
               string.Equals(left.Text, right.Text, StringComparison.Ordinal) &&
               string.Equals(
                   RawSpanText(leftTokens, op.LeftStart, op.LeftEnd),
                   RawSpanText(rightTokens, op.RightStart, op.RightEnd),
                   StringComparison.Ordinal);
    }

    private static bool SameRawWhitespaceToken(IrDiffToken left, IrDiffToken right) =>
        left.Kind == IrDiffTokenKind.Separator && right.Kind == IrDiffTokenKind.Separator &&
        left.Text == " " && right.Text == " " &&
        left.MatchKey == right.MatchKey &&
        string.Equals(left.Text, right.Text, StringComparison.Ordinal);

    private static List<XElement> BuildTokenOpContent(
        IrTokenDiff tokenDiff,
        IReadOnlyList<IrDiffToken> leftTokens, IReadOnlyList<IrDiffToken> rightTokens,
        SourceRunModel leftRuns, SourceRunModel rightRuns, RenderState state,
        RevKind? stableSpanKind = null)
    {
        var content = new List<XElement>();

        // Fine paragraph changes normally leave equal/format-changed spans bare. A MoveModify destination
        // supplies MoveTo here instead: the spans still receive their ordinary rPrChange history first, then
        // the resulting run-level element is wrapped as relocated content.
        void AddStableSpan(XElement runLevel)
        {
            state.RegisterMediaReferences(runLevel);
            if (stableSpanKind is { } kind)
                content.AddRange(WrapFieldAware(runLevel, kind, state));
            else
                content.Add(runLevel);
        }

        foreach (var tokenOp in CoalesceTokenOpsWordShaped(
                     tokenDiff.Ops, leftTokens, rightTokens, state.Settings.FormatComparison,
                     leftRuns, rightRuns))
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
                            content.AddRange(WrapFieldAware(r, RevKind.Del, state));
                        foreach (var r in rightRuns.Slice(rs, re, rzs, rze))
                            content.AddRange(WrapFieldAware(r, RevKind.Ins, state));   // registers media on its clone
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
                            AddStableSpan(r);
                        }
                    }
                    else
                    {
                        foreach (var r in rightRuns.Slice(rs, re, rzs, rze))
                            AddStableSpan(r);
                    }
                    break;
                }
                case IrTokenOpKind.Insert:
                {
                    var (s, e) = RightSpanChars(rightTokens, tokenOp);
                    var (zs, ze) = ZeroWidthBoundaries(rightTokens, tokenOp.RightStart, tokenOp.RightEnd);
                    foreach (var r in rightRuns.Slice(s, e, zs, ze))
                        content.AddRange(WrapFieldAware(r, RevKind.Ins, state));   // registers media on its clone
                    break;
                }
                case IrTokenOpKind.Delete:
                {
                    var (s, e) = LeftSpanChars(leftTokens, tokenOp);
                    var (zs, ze) = ZeroWidthBoundaries(leftTokens, tokenOp.LeftStart, tokenOp.LeftEnd);
                    foreach (var r in leftRuns.Slice(s, e, zs, ze))
                        content.AddRange(WrapFieldAware(r, RevKind.Del, state));
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
            // Coalesce the run into ONE w:hyperlink when EITHER signal says the link's target is the SAME on both
            // sides:
            //   (1) it carries at least one plain (Equal/unchanged) run — some anchor text matched on both sides,
            //       so the link is the same link with an intra-anchor edit; OR
            //   (2) every fragment resolves to the SAME hyperlink TARGET (#232) — a single link whose ENTIRE
            //       anchor was replaced (pure del+ins, no shared token) still has one, unchanged target, so it
            //       must render as ONE w:hyperlink to be byte-faithful to the source's link structure.
            // A run with NO plain piece AND DIFFERING targets is a pure del→ins RETARGET (WC019: text AND href both
            // change → del-link of the OLD target, ins-link of the NEW target, the new id remapped post-assembly) —
            // those MUST stay separate so the remap + empty-shell-drop restore the right/left link on each side.
            // The target gate (not r:id-string equality) is essential: at coalesce time WC019's del/ins fragments
            // carry the SAME r:id STRING (the collision the remap later resolves), so only the RESOLVED target tells
            // a same-target text-rewrite apart from a genuine retarget.
            bool anyPlainRun = false;
            for (int k = i; k < j; k++)
                if (!IsRevisionPureHyperlink(content[k]))
                {
                    anyPlainRun = true;
                    break;
                }

            if (runLen > 1 && (anyPlainRun || AllFragmentsShareResolvedTarget(content, i, j)))
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

        // Strip the transient source-link markers (ordinal + resolved target) from every emitted wrapper so they
        // never reach output. (The body-wide pt: sweep in Render would also catch them, but stripping here keeps
        // the markers an internal detail of the coalescer and protects callers that inspect the returned content
        // before that sweep.)
        foreach (var el in merged)
            if (el.Name == W.hyperlink)
            {
                el.Attribute(SourceLinkId)?.Remove();
                el.Attribute(SourceLinkTarget)?.Remove();
            }
        return merged;
    }

    /// <summary>True iff every wrapper in the half-open run <c>[start,end)</c> carries the SAME non-empty resolved
    /// hyperlink target (<see cref="SourceLinkTarget"/>). This is the #232 signal that a single source link whose
    /// entire anchor was replaced (pure <c>w:del</c>+<c>w:ins</c>, no shared token) is nonetheless ONE link with an
    /// unchanged target, so its fragments may be rejoined. A missing/empty target on any fragment (dangling or
    /// unresolvable id) fails conservatively — the run is not merged on the target basis. Because a del fragment
    /// carries the LEFT target and an ins/plain fragment the RIGHT target, agreement across the run means the
    /// target is identical on both sides — precisely what distinguishes a same-target text rewrite from a WC019
    /// retarget (whose del/ins targets differ).</summary>
    private static bool AllFragmentsShareResolvedTarget(List<XElement> content, int start, int end)
    {
        var target = content[start].Attribute(SourceLinkTarget)?.Value;
        if (string.IsNullOrEmpty(target))
            return false;
        for (int k = start + 1; k < end; k++)
            if (content[k].Attribute(SourceLinkTarget)?.Value != target)
                return false;
        return true;
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
        // A deleted field instruction is w:delInstrText (mirrors Word; w:instrText alone in a w:del confuses
        // a field re-parser). Only field runs carry instrText, so this is a no-op for ordinary text.
        foreach (var it in runLevel.DescendantsAndSelf(W.instrText).ToList())
            it.Name = W.delInstrText;
    }

    /// <summary>
    /// A <c>w:fldSimple</c> is NOT a valid child of <c>w:ins</c>/<c>w:del</c>, so wrapping a field-bearing
    /// run-level element directly is schema-invalid and keeping the <c>w:fldSimple</c> wrapper bare leaves a
    /// DANGLING empty field on accept/reject (e.g. an orphaned <c>REF</c> after the referencing paragraph is
    /// deleted — the field code survives with no content). Expand it to the equivalent <c>fldChar</c> run
    /// sequence (begin / instrText / separate / cached result / end) so the WHOLE field — code and result —
    /// rides in revision-wrappable runs and toggles cleanly: accept of a deletion drops the entire field,
    /// reject restores it live. The simple field's <c>w:dirty</c>/<c>w:fldLock</c> state and <c>w:fldData</c>
    /// move to the generated begin <c>w:fldChar</c>, preserving field semantics even though a direct simple-field
    /// serialization itself cannot live inside a revision. Any non-field run-level element is yielded unchanged.
    /// </summary>
    private static IEnumerable<XElement> ExpandFieldForRevision(XElement runLevel)
    {
        if (runLevel.Name != W.fldSimple)
        {
            yield return runLevel;
            yield break;
        }
        var instr = (string?)runLevel.Attribute(W.instr) ?? "";
        var begin = new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin"));
        foreach (var attribute in runLevel.Attributes().Where(attribute =>
                     attribute.Name == W.dirty || attribute.Name == W.fldLock))
            begin.Add(new XAttribute(attribute));
        foreach (var fieldData in runLevel.Elements(W.fldData))
            begin.Add(new XElement(fieldData));
        yield return new XElement(W.r, begin);
        yield return new XElement(W.r,
            new XElement(W.instrText, new XAttribute(XNamespace.Xml + "space", "preserve"), instr));
        yield return new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate")));
        foreach (var child in runLevel.Elements().Where(child => child.Name != W.fldData))
            // A simple field may legally contain another simple field. Flatten nested fields to their equivalent
            // fldChar run sequence before revisions are applied; otherwise the inner fldSimple would become an
            // invalid direct child of w:ins/w:del.
            foreach (var expanded in ExpandFieldForRevision(child))
                yield return expanded;
        yield return new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "end")));
    }

    /// <summary>Wrap a run-level element in a revision (<see cref="WrapRunLevel"/>), first expanding a
    /// <c>w:fldSimple</c> via <see cref="ExpandFieldForRevision"/> so a field is fully deletable/insertable.
    /// Yields one or more wrapped run-level elements.</summary>
    private static IEnumerable<XElement> WrapFieldAware(XElement runLevel, RevKind kind, RenderState state) =>
        ExpandFieldForRevision(runLevel).Select(part => WrapRunLevel(part, kind, state));

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

    /// <summary>
    /// Stamp native PARAGRAPH-property revision markup (block-format-change family, 2026-07-03) on an
    /// emitted right-sourced paragraph: a <c>w:pPrChange</c> as the LAST child of the paragraph's pPr when
    /// the pPr differs under the format-comparison policy (inner = the LEFT pPr minus <c>w:rPr</c>/
    /// <c>w:sectPr</c>/<c>w:pPrChange</c> — the CT_PPrBase constraint: the mark rPr and section props are
    /// outside pPrChange scope by schema), and a <c>w:pPr/w:rPr/w:rPrChange</c> when the paragraph-MARK
    /// rPr differs canonically (inner = the LEFT mark rPr). Accept keeps the right properties (markers
    /// drop); reject restores the left (RevisionProcessor swaps the inner back, carrying the current mark
    /// rPr and inline sectPr). No-op when <see cref="IrDiffSettings.TrackBlockFormatChanges"/> is off —
    /// the Consolidate-pipeline pin.
    /// </summary>
    private static void ApplyBlockFormatChanges(XElement newPara, XElement leftPara, XElement rightPara, RenderState state)
    {
        // Two independent slices: pPr (+ mark rPr) gated on the PARAGRAPH flag, the inline sectPr gated on
        // the block flag (so the composite can stamp pPrChange but not sectPrChange — B1 vs B2).
        bool trackPPr = state.Settings.TrackParagraphFormatChanges;
        bool trackSect = state.Settings.TrackSectionFormatChanges;
        if (!trackPPr && !trackSect)
            return;

        var leftPPr = leftPara.Element(W.pPr);
        var rightPPr = rightPara.Element(W.pPr);

        // Policy-gated pPr delta: ModeledOnly compares the modeled ParaKey (the delta a consumer-grade
        // report can describe; unmodeled-only deltas stay untracked — the documented rPr-parallel blind
        // spot); Full compares canonically (unid/rsid-stripped, rPr/sectPr/pPrChange excluded).
        bool pPrDiffers = trackPPr && (state.Settings.FormatComparison == IrFormatComparison.ModeledOnly
            ? IrModeledFormat.ParaKey(IrReader.MapParaFormat(leftPPr)) !=
              IrModeledFormat.ParaKey(IrReader.MapParaFormat(rightPPr))
            : !IrHasher.CanonicalHash(PPrForCompare(leftPPr)).Equals(IrHasher.CanonicalHash(PPrForCompare(rightPPr))));

        // The paragraph-MARK rPr is outside pPrChange by schema; Word tracks it as w:pPr/w:rPr/w:rPrChange.
        // Compared canonically under both policies (the mark has no token to carry a run-level rPrChange).
        bool markDiffers = trackPPr && !IrHasher.CanonicalHash(MarkRPrForCompare(leftPPr))
            .Equals(IrHasher.CanonicalHash(MarkRPrForCompare(rightPPr)));

        // Mid-document section-property change (A3): the emitted right pPr already carries the right inline
        // w:sectPr (cloned above); when the left pPr also had one and their PROPERTIES differ, it is tracked
        // as w:sectPrChange INSIDE that sectPr (NOT in the pPrChange inner — CT_PPrBase excludes sectPr). A
        // one-sided inline sectPr (added/removed) is a structural change, not a property change — left as-is.
        var leftInlineSect = leftPPr?.Element(W.sectPr);
        var newPPrForSect = newPara.Element(W.pPr);
        var rightInlineSect = newPPrForSect?.Element(W.sectPr);
        bool sectDiffers = trackSect && leftInlineSect != null && rightInlineSect != null
            && SectPrPropsDiffer(leftInlineSect, rightInlineSect);

        if (!pPrDiffers && !markDiffers && !sectDiffers)
            return;

        var pPr = newPara.Element(W.pPr);
        if (pPr == null)
        {
            pPr = new XElement(W.pPr);
            newPara.AddFirst(pPr);
        }

        if (markDiffers)
        {
            var markRPr = pPr.Element(W.rPr);
            if (markRPr == null)
            {
                markRPr = new XElement(W.rPr);
                // Schema position: the mark rPr follows every pPr property child, before sectPr/pPrChange.
                var before = pPr.Elements().FirstOrDefault(e => e.Name == W.sectPr || e.Name == W.pPrChange);
                if (before != null) before.AddBeforeSelf(markRPr);
                else pPr.Add(markRPr);
            }
            markRPr.Elements(W.rPrChange).Remove();
            var leftMark = leftPPr?.Element(W.rPr);
            var markInner = leftMark != null
                ? StripUnids(new XElement(W.rPr, leftMark.Attributes(),
                      leftMark.Elements().Where(e => e.Name != W.rPrChange)))
                : new XElement(W.rPr);
            markRPr.Add(new XElement(W.rPrChange, state.RevisionAttributes(), markInner));
        }

        if (pPrDiffers)
        {
            var inner = leftPPr != null
                ? StripUnids(new XElement(W.pPr, leftPPr.Attributes(),
                      leftPPr.Elements().Where(e => e.Name != W.rPr && e.Name != W.sectPr && e.Name != W.pPrChange)))
                : new XElement(W.pPr);
            pPr.Elements(W.pPrChange).Remove();
            pPr.Add(new XElement(W.pPrChange, state.RevisionAttributes(), inner));   // last child of pPr
        }

        if (sectDiffers)
            // output (rightInlineSect) currently holds the RIGHT props; old = leftInlineSect.
            ApplySectPrChange(rightInlineSect!, leftInlineSect!, rightInlineSect!, state);
    }

    /// <summary>The pPrChange-comparable projection of a pPr: its property children only (no mark rPr, no
    /// inline sectPr, no pPrChange marker); a null pPr compares as an EMPTY pPr (no direct properties).</summary>
    private static XElement PPrForCompare(XElement? pPr) =>
        pPr == null
            ? new XElement(W.pPr)
            : new XElement(W.pPr, pPr.Attributes(),
                  pPr.Elements().Where(e => e.Name != W.rPr && e.Name != W.sectPr && e.Name != W.pPrChange));

    /// <summary>The paragraph-mark rPr of a pPr, minus any rPrChange marker; null/absent compares as empty.</summary>
    private static XElement MarkRPrForCompare(XElement? pPr)
    {
        var rPr = pPr?.Element(W.rPr);
        return rPr == null
            ? new XElement(W.rPr)
            : new XElement(W.rPr, rPr.Attributes(), rPr.Elements().Where(e => e.Name != W.rPrChange));
    }

    // ----------------------------------------------- table-shell property revisions (block-format family)

    // Per-shell inner-exclusion sets (the CT_*Base vs CT_* delta): a *PrChange's inner carries the shell's
    // FORMAT children only — never its own change marker or the revision markers a redline layers on top.
    private static readonly XName[] TblPrInnerExclude = { W.tblPrChange };
    private static readonly XName[] TblGridInnerExclude = { W.tblGridChange };
    private static readonly XName[] TrPrInnerExclude = { W.ins, W.del, W.trPrChange };
    private static readonly XName[] TcPrInnerExclude = { W.cellIns, W.cellDel, W.cellMerge, W.tcPrChange };
    private static readonly XName[] TblPrExInnerExclude = { W.tblPrExChange };

    /// <summary>
    /// Stamp the two TABLE-LEVEL native shell markers where the emitted right shell differs from the left:
    /// <c>w:tblPrChange</c> and <c>w:tblGridChange</c>. Row/cell shells are handled per-row by
    /// <see cref="ApplyRowAndCellShellChanges"/> (the Modified-table path interleaves inserted/deleted rows,
    /// so only the caller — which pairs rows by anchor — knows the correct left row). No-op when block-format
    /// tracking is off.
    /// </summary>
    private static void ApplyTableLevelShellChanges(XElement newTbl, XElement leftTbl, RenderState state)
    {
        if (!state.Settings.TrackTableFormatChanges)
            return;

        ApplyShellChange(newTbl, W.tblPr, W.tblPrChange, leftTbl.Element(W.tblPr), state, idOnly: false, TblPrInnerExclude);
        ApplyShellChange(newTbl, W.tblGrid, W.tblGridChange, leftTbl.Element(W.tblGrid), state, idOnly: true, TblGridInnerExclude);
    }

    /// <summary>Stamp a row's <c>w:trPrChange</c> and each of its cells' <c>w:tcPrChange</c> (cells paired
    /// positionally against the left row) where the emitted right shells differ from the left.  This helper
    /// remains for content-equal/legacy positional rows; ordinary-grid ModifyRow rendering calls the row and
    /// cell halves separately so an interior cellIns cannot shift tcPr history pairing.</summary>
    private static void ApplyRowAndCellShellChanges(XElement newRow, XElement leftRow, RenderState state)
    {
        ApplyRowShellChanges(newRow, leftRow, state);
        if (!state.Settings.TrackTableFormatChanges)
            return;

        var leftCells = leftRow.Elements(W.tc).ToList();
        var newCells = newRow.Elements(W.tc).ToList();
        for (int c = 0; c < newCells.Count && c < leftCells.Count; c++)
            ApplyPairedCellShellChange(newCells[c], leftCells[c], state);
    }

    /// <summary>Apply only the paired row-level shell histories.</summary>
    private static void ApplyRowShellChanges(XElement newRow, XElement leftRow, RenderState state)
    {
        if (!state.Settings.TrackTableFormatChanges)
            return;

        ApplyShellChange(newRow, W.trPr, W.trPrChange, leftRow.Element(W.trPr), state, idOnly: false, TrPrInnerExclude);
        ApplyShellChange(newRow, W.tblPrEx, W.tblPrExChange, leftRow.Element(W.tblPrEx), state, idOnly: false, TblPrExInnerExclude);
    }

    /// <summary>
    /// Apply a tcPr history for one explicitly paired cell.  A right-only cell insertion deliberately never
    /// calls this: its <c>w:cellIns</c> is its complete revision history, and borrowing a shifted left tcPr
    /// would corrupt rejection.
    /// </summary>
    private static void ApplyPairedCellShellChange(XElement newCell, XElement leftCell, RenderState state)
    {
        if (!state.Settings.TrackTableFormatChanges)
            return;
        ApplyShellChange(newCell, W.tcPr, W.tcPrChange, leftCell.Element(W.tcPr), state,
            idOnly: false, TcPrInnerExclude);
    }

    /// <summary>Apply ONE composed block-format shell across reviewers (Consolidate B2): swap the winning
    /// reviewer's shell into <paramref name="host"/> (accept ≡ winner) and stamp the change marker with
    /// inner = BASE shell (reject ≡ base), attributed to the winner's author. No-op when
    /// <paramref name="shellRef"/> is null (base shell kept — already present in host from its base clone) or
    /// the table slice is off.</summary>
    private static void ApplyComposedShell(
        XElement host, XElement baseHost, IrComposedShellRef? shellRef,
        XName shellName, XName changeName, XName[] innerExclude, bool idOnly,
        IReadOnlyList<IrDocument> reviewerIrs, RenderState state,
        Func<IrDocument, string, XElement?> findWinnerHost)
    {
        if (!state.Settings.TrackTableFormatChanges || shellRef == null
            || shellRef.Reviewer < 0 || shellRef.Reviewer >= reviewerIrs.Count)
            return;
        var winnerHost = findWinnerHost(reviewerIrs[shellRef.Reviewer], shellRef.RightAnchor);
        if (winnerHost == null)
            return;

        // Swap in the winner's shell (accept ≡ winner): drop host's base-cloned shell, insert the winner's in
        // schema order; ApplyShellChange then captures the BASE shell into the marker inner (reject ≡ base).
        host.Elements(shellName).Remove();
        if (winnerHost.Element(shellName) is { } winnerShell)
            InsertShellInSchemaOrder(host, StripUnids(new XElement(winnerShell)), shellName);

        var saved = state.AuthorOverride;
        state.AuthorOverride = shellRef.Author;
        ApplyShellChange(host, shellName, changeName, baseHost.Element(shellName), state, idOnly, innerExclude);
        state.AuthorOverride = saved;
    }

    /// <summary>Apply the composed TABLE-level shells (tblPr/tblGrid) to a composed table's output element (B2).</summary>
    private static void ApplyComposedTableShell(
        XElement newTbl, XElement baseTbl, IrComposedTableShell? shell,
        IReadOnlyList<IrDocument> reviewerIrs, RenderState state)
    {
        if (shell == null)
            return;
        ApplyComposedShell(newTbl, baseTbl, shell.TblPr, W.tblPr, W.tblPrChange, TblPrInnerExclude, idOnly: false, reviewerIrs, state, FindTableSource);
        ApplyComposedShell(newTbl, baseTbl, shell.TblGrid, W.tblGrid, W.tblGridChange, TblGridInnerExclude, idOnly: true, reviewerIrs, state, FindTableSource);
    }

    /// <summary>Apply the composed ROW-level shells (trPr/tblPrEx) to a composed row's output element (B2).</summary>
    private static void ApplyComposedRowShell(
        XElement newRow, XElement baseRow, IrAuthoredRowOp rowOp,
        IReadOnlyList<IrDocument> reviewerIrs, RenderState state)
    {
        ApplyComposedShell(newRow, baseRow, rowOp.TrPr, W.trPr, W.trPrChange, TrPrInnerExclude, idOnly: false, reviewerIrs, state, FindRowSource);
        ApplyComposedShell(newRow, baseRow, rowOp.TblPrEx, W.tblPrEx, W.tblPrExChange, TblPrExInnerExclude, idOnly: false, reviewerIrs, state, FindRowSource);
    }

    /// <summary>The source <c>w:tbl</c> a table anchor resolves to in <paramref name="ir"/> (tables ARE in the
    /// AnchorIndex, unlike rows/cells).</summary>
    private static XElement? FindTableSource(IrDocument ir, string tableAnchor) =>
        ir.AnchorIndex.TryGetValue(tableAnchor, out var b) && b is IrTable t ? t.Source.Element : null;

    /// <summary>Stamp <c>w:sectPrChange</c> on the composed output's trailing <c>w:sectPr</c> for a composed
    /// section change (Consolidate B2): resolve the winning reviewer's trailing <c>w:sectPr</c> and apply its
    /// properties (accept ≡ winner) while capturing the base properties into the marker (reject ≡ base),
    /// attributed to the winner's author. Header/footer references are preserved (owned by the hdr/ftr
    /// machinery). No-op when the section slice is off or no winner was attributed.</summary>
    internal static void ApplyComposedTrailingSectPr(
        XElement trailingSectPr, IrComposedShellRef? sectRef,
        IReadOnlyList<IrDocument> reviewerIrs, RenderState state)
    {
        if (!state.Settings.TrackSectionFormatChanges || sectRef == null
            || sectRef.Reviewer < 0 || sectRef.Reviewer >= reviewerIrs.Count)
            return;
        var winnerSectPr = SourceElement(sectRef.RightAnchor, reviewerIrs[sectRef.Reviewer]);
        if (winnerSectPr == null || winnerSectPr.Name != W.sectPr || !SectPrPropsDiffer(trailingSectPr, winnerSectPr))
            return;

        var saved = state.AuthorOverride;
        state.AuthorOverride = sectRef.Author;
        ApplySectPrChange(trailingSectPr, trailingSectPr, winnerSectPr, state);
        state.AuthorOverride = saved;
    }

    /// <summary>
    /// Core table-shell stamper. <paramref name="host"/> already carries the RIGHT shell (from a verbatim
    /// clone) or none. When the left/right shells differ (canonical, excluding <paramref name="innerExclude"/>
    /// + the change marker), append <paramref name="changeName"/> as the LAST child of the (right) shell,
    /// creating an empty shell in schema position if the right had none. The change's inner is the LEFT shell's
    /// format children (unids stripped, exclusions dropped). <paramref name="idOnly"/> ⇒ CT_Markup attributes
    /// (a bare <c>w:id</c>, as <c>w:tblGridChange</c> requires); otherwise author/date/id.
    /// </summary>
    private static void ApplyShellChange(
        XElement host, XName shellName, XName changeName, XElement? leftShell, RenderState state,
        bool idOnly, XName[] innerExclude)
    {
        var rightShell = host.Element(shellName);
        if (!ShellDiffers(leftShell, rightShell, changeName, innerExclude))
            return;

        if (rightShell == null)
        {
            rightShell = new XElement(shellName);
            InsertShellInSchemaOrder(host, rightShell, shellName);
        }

        rightShell.Elements(changeName).Remove();   // idempotence

        var inner = leftShell != null
            ? StripUnids(new XElement(shellName, leftShell.Attributes(),
                  leftShell.Elements().Where(e => e.Name != changeName && !innerExclude.Contains(e.Name))))
            : new XElement(shellName);

        var change = idOnly
            ? new XElement(changeName, new XAttribute(W.id, state.NextId()), inner)
            : new XElement(changeName, state.RevisionAttributes(), inner);
        rightShell.Add(change);   // last child of the shell
    }

    /// <summary>True when two table shells differ ignoring their change markers, the listed revision markers,
    /// and canonical noise (unids/rsids). An absent shell compares equal to an empty/format-less shell.</summary>
    private static bool ShellDiffers(XElement? left, XElement? right, XName changeName, XName[] exclude)
        => !IrHasher.CanonicalHash(CleanShell(left, changeName, exclude))
            .Equals(IrHasher.CanonicalHash(CleanShell(right, changeName, exclude)));

    /// <summary>A shell projected to its comparable format content under a FIXED container name (so absent vs
    /// empty-shell compare equal): the shell's attributes + its non-excluded, non-change-marker children.</summary>
    private static XElement CleanShell(XElement? shell, XName changeName, XName[] exclude)
    {
        var c = new XElement("shell");
        if (shell != null)
        {
            c.Add(shell.Attributes());
            foreach (var e in shell.Elements().Where(e => e.Name != changeName && !exclude.Contains(e.Name)))
                c.Add(new XElement(e));
        }
        return c;
    }

    /// <summary>Insert a freshly-created shell element at its schema position: <c>w:tblPr</c> first in the
    /// table; <c>w:tblGrid</c> after <c>w:tblPr</c>; <c>w:trPr</c> after an existing <c>w:tblPrEx</c> (CT_Row
    /// orders <c>tblPrEx</c> before <c>trPr</c>), else first; <c>w:tcPr</c> first in the cell.</summary>
    private static void InsertShellInSchemaOrder(XElement host, XElement shell, XName shellName)
    {
        if (shellName == W.tblGrid)
        {
            var tblPr = host.Element(W.tblPr);
            if (tblPr != null) { tblPr.AddAfterSelf(shell); return; }
        }
        else if (shellName == W.trPr)
        {
            var tblPrEx = host.Element(W.tblPrEx);
            if (tblPrEx != null) { tblPrEx.AddAfterSelf(shell); return; }
        }
        host.AddFirst(shell);
    }

    // -------------------------------------------------- trailing section-property revision (Phase 3)

    /// <summary>A sectPr child that belongs to the tracked PROPERTY set: everything except the header/footer
    /// references (owned by the header/footer machinery) and the change marker itself (CT_SectPrBase).</summary>
    private static bool IsSectPrProp(XElement e)
        => e.Name != W.headerReference && e.Name != W.footerReference && e.Name != W.sectPrChange;

    /// <summary>True when two trailing sectPrs differ in their tracked PROPERTIES (page setup/columns/…),
    /// ignoring header/footer references, the change marker, and canonical noise (rsids).</summary>
    private static bool SectPrPropsDiffer(XElement left, XElement right)
        => !IrHasher.CanonicalHash(SectPrPropsContainer(left)).Equals(IrHasher.CanonicalHash(SectPrPropsContainer(right)));

    private static XElement SectPrPropsContainer(XElement sectPr)
    {
        var c = new XElement("sect");
        foreach (var e in sectPr.Elements().Where(IsSectPrProp))
            c.Add(new XElement(e));
        return c;
    }

    /// <summary>
    /// Stamp native <c>w:sectPrChange</c> on the output (left-based) trailing sectPr: capture the LEFT
    /// properties into the change marker's inner (CT_SectPrBase — no references), replace the output's
    /// properties with the RIGHT (accepted-state) properties, and keep the output's references (owned by the
    /// header/footer machinery). Accept drops the marker (right properties remain); reject restores the left
    /// properties while preserving the references (via the <see cref="RevisionProcessor"/> sectPrChange fix).
    /// </summary>
    private static void ApplySectPrChange(XElement outputSectPr, XElement oldSectPr, XElement rightSectPr, RenderState state)
    {
        // Snapshot BOTH the OLD (change-inner) props and the RIGHT (accepted) props BEFORE mutating output —
        // `outputSectPr` may alias `oldSectPr` (the trailing case: output starts as the left clone) OR
        // `rightSectPr` (the mid-doc inline case: output starts as the right clone), so reading either after
        // the strip would be wrong.
        var inner = new XElement(W.sectPr);
        foreach (var e in oldSectPr.Elements().Where(IsSectPrProp))
            inner.Add(StripUnids(new XElement(e)));
        var rightProps = rightSectPr.Elements().Where(IsSectPrProp)
            .Select(e => StripUnids(new XElement(e))).ToList();

        outputSectPr.Elements().Where(IsSectPrProp).Remove();   // strip current props (references stay)
        foreach (var e in rightProps)
            outputSectPr.Add(e);                                 // apply the right (accepted) props after refs
        outputSectPr.Add(new XElement(W.sectPrChange, state.RevisionAttributes(), inner));   // last child
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
        List<XElement> rightClones, OpenXmlPart leftMain, OpenXmlPart rightMain)
    {
        var leftHyper = leftMain.HyperlinkRelationships.ToDictionary(r => r.Id, StringComparer.Ordinal);
        var leftExternalIds = new HashSet<string>(leftMain.ExternalRelationships.Select(r => r.Id), StringComparer.Ordinal);
        // ALL ids in use on the left part, any relationship kind. An id "free" among left hyperlinks
        // may still be TAKEN by a part relationship (comments.xml, an image, ...) — recreating the right
        // relationship under it makes System.IO.Packaging throw XmlException ("ID conflicts with the ID
        // of an existing relationship"), so those must take the remap path, not the same-id path.
        var leftUsedIds = UsedRelationshipIds(leftMain);
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
                if (!leftUsedIds.Contains(id))
                {
                    // The id is FREE in the left part — recreate the relationship under the SAME id so the
                    // cloned w:hyperlink/@r:id keeps resolving (the common, no-collision case).
                    try { leftMain.AddHyperlinkRelationship(hr.Uri, hr.IsExternal, id); }
                    catch (Exception ex) when (ex is ArgumentException or InvalidOperationException) { }
                }
                else if (!leftHyper.ContainsKey(id))
                {
                    // The id is taken by a left relationship of a DIFFERENT KIND (a part relationship —
                    // comments.xml, an image, ...): recreating under the same id is impossible, and the
                    // cloned r:id would resolve to the wrong object. Remap to a fresh id.
                    if (hr.Uri is { } takenUri)
                    {
                        string fresh = FreshRelationshipId(leftMain);
                        leftMain.AddHyperlinkRelationship(takenUri, hr.IsExternal, fresh);
                        RewriteReferenceId(rightClones, id, fresh);
                    }
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
                if (!leftUsedIds.Contains(id))
                {
                    try { leftMain.AddExternalRelationship(er.RelationshipType, er.Uri, id); }
                    catch (Exception ex) when (ex is ArgumentException or InvalidOperationException) { }
                }
                else if (!leftExternalIds.Contains(id))
                {
                    // Taken by a non-external left relationship (part/hyperlink/data) — remap, as above.
                    if (er.Uri is { } takenUri)
                    {
                        string fresh = FreshRelationshipId(leftMain);
                        leftMain.AddExternalRelationship(er.RelationshipType, takenUri, fresh);
                        RewriteReferenceId(rightClones, id, fresh);
                    }
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

    /// <summary>One deliberately narrow style-normalization candidate. It is created only for a
    /// total main-body replacement whose LEFT stories cannot observe a newly-defaulted paragraph
    /// style after rejection. The style ids are the complete RIGHT paragraph-style closure actually
    /// reachable from inserted blocks, including default and basedOn ancestors.</summary>
    private sealed class InsertedStyleNormalization
    {
        public InsertedStyleNormalization(HashSet<string> usedStyleIds)
        {
            UsedStyleIds = usedStyleIds;
        }

        public HashSet<string> UsedStyleIds { get; }

        public bool Uses(string styleId) => UsedStyleIds.Contains(styleId);
    }

    /// <summary>
    /// Build the first reversible document-presentation plan. It intentionally admits only literal docDefaults
    /// changes in a simple story universe: package parts remain left-owned, and no theme/numbering/settings,
    /// glossary, pre-existing style history, or complex consumer can be mistaken for a style-revision problem.
    /// Unsupported presentation changes stay visible as the left package rather than being silently copied into
    /// both accept and reject views.
    /// </summary>
    private static DocDefaultsStyleProjection? TryCreateDocDefaultsStyleProjection(
        WordprocessingDocument leftDocument, WordprocessingDocument rightDocument, RenderState state)
    {
        if (state.Settings.PreserveInputRevisions)
            return null;

        var leftMain = leftDocument.MainDocumentPart;
        var rightMain = rightDocument.MainDocumentPart;
        var leftStyles = leftMain?.StyleDefinitionsPart?.GetXDocument().Root;
        var rightStyles = rightMain?.StyleDefinitionsPart?.GetXDocument().Root;
        if (leftMain is null || rightMain is null || leftStyles is null || rightStyles is null ||
            leftMain.GlossaryDocumentPart is not null ||
            DocDefaultsPayloadsEqual(leftStyles, rightStyles) ||
            !PartsEqual(leftMain.ThemePart, rightMain.ThemePart) ||
            !PartsEqual(leftMain.NumberingDefinitionsPart, rightMain.NumberingDefinitionsPart) ||
            !PartsEqual(leftMain.DocumentSettingsPart, rightMain.DocumentSettingsPart) ||
            HasThemeReference(leftStyles) || HasThemeReference(rightStyles) ||
            HasUnsafePresentationConsumer(state.Left) || HasUnsafePresentationConsumer(state.Right))
            return null;

        var candidates = CollectUsedStyleIdentities(state.Left);
        candidates.UnionWith(CollectUsedStyleIdentities(state.Right));
        candidates.RemoveWhere(identity => !StyleChainIsSharedAndResolvable(leftStyles, rightStyles, identity));
        return candidates.Count == 0 ? null : new DocDefaultsStyleProjection(candidates);
    }

    private static bool DocDefaultsPayloadsEqual(XElement leftStyles, XElement rightStyles)
    {
        XElement Payload(XElement styles) => new XElement(W.docDefaults,
            new XElement(W.pPrDefault,
                new XElement(W.pPr, styles.Element(W.docDefaults)?.Element(W.pPrDefault)?.Element(W.pPr)?.Elements())),
            new XElement(W.rPrDefault,
                new XElement(W.rPr, styles.Element(W.docDefaults)?.Element(W.rPrDefault)?.Element(W.rPr)?.Elements())));
        var left = Payload(leftStyles);
        var right = Payload(rightStyles);
        StripStyleNoise(left);
        StripStyleNoise(right);
        return XNode.DeepEquals(left, right);
    }

    private static bool PartsEqual(OpenXmlPart? left, OpenXmlPart? right)
    {
        if (left is null || right is null)
            return left is null && right is null;
        using var leftStream = left.GetStream(FileMode.Open, FileAccess.Read);
        using var rightStream = right.GetStream(FileMode.Open, FileAccess.Read);
        while (true)
        {
            int a = leftStream.ReadByte();
            int b = rightStream.ReadByte();
            if (a != b)
                return false;
            if (a < 0)
                return true;
        }
    }

    private static bool HasThemeReference(XElement root) => root.DescendantsAndSelf().Attributes()
        .Any(attribute => attribute.Name.LocalName.IndexOf("theme", StringComparison.OrdinalIgnoreCase) >= 0);

    private static bool HasUnsafePresentationConsumer(IrDocument document)
    {
        // This first style-level slice deliberately avoids shapes whose effective appearance includes a higher
        // precedence layer (table conditional styles and list labels), an independent package graph (drawing),
        // or an envelope that needs its own structural revision. Later presentation-plan phases can support each
        // explicitly.
        var unsafeNames = new HashSet<string>(StringComparer.Ordinal)
        {
            "tbl", "numPr", "drawing", "pict", "txbxContent", "fldSimple", "fldChar", "instrText", "delInstrText",
            "sdt", "smartTag",
        };
        return document.Sources.Values.Any(source => source.Root?.DescendantsAndSelf().Any(element =>
            element.Name.Namespace == W.w && unsafeNames.Contains(element.Name.LocalName)) == true) ||
            document.Sources.Values.Any(source => source.Root is not null && HasThemeReference(source.Root));
    }

    private static HashSet<StyleIdentity> CollectUsedStyleIdentities(IrDocument document)
    {
        var identities = new HashSet<StyleIdentity>();
        foreach (var paragraph in document.AnchorIndex.Values.OfType<IrParagraph>())
        {
            var paragraphStyle = paragraph.Format.StyleId ?? document.Styles.DefaultParagraphStyleId;
            if (!string.IsNullOrEmpty(paragraphStyle))
                identities.Add(new StyleIdentity("paragraph", paragraphStyle));
            CollectInlineCharacterStyles(paragraph.Inlines, identities);
        }

        // Most run-like IR nodes preserve their character style through IrFormat, but a break or note-reference
        // run can legally own w:rStyle without carrying an IrFormat. Source XML is retained by the renderer and
        // represents the accepted view, so supplement the IR scan with its direct live style references rather
        // than overlooking a visible consumer of the projected defaults.
        foreach (var source in document.Sources.Values)
        foreach (var styleRef in source.Descendants().Where(element =>
                     element.Name == W.pStyle || element.Name == W.rStyle))
        {
            if (styleRef.Ancestors().Any(ancestor =>
                    ancestor.Name == W.pPrChange || ancestor.Name == W.rPrChange))
                continue;
            var styleId = (string?)styleRef.Attribute(W.val);
            if (!string.IsNullOrEmpty(styleId))
                identities.Add(new StyleIdentity(styleRef.Name == W.pStyle ? "paragraph" : "character", styleId));
        }
        return identities;
    }

    private static void CollectInlineCharacterStyles(IEnumerable<IrInline> inlines, HashSet<StyleIdentity> identities)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case IrTextRun text when !string.IsNullOrEmpty(text.Format.StyleId):
                    identities.Add(new StyleIdentity("character", text.Format.StyleId));
                    break;
                case IrTab tab when !string.IsNullOrEmpty(tab.Format.StyleId):
                    identities.Add(new StyleIdentity("character", tab.Format.StyleId));
                    break;
                case IrHyperlink hyperlink:
                    CollectInlineCharacterStyles(hyperlink.Inlines, identities);
                    break;
                case IrFieldRun field:
                    CollectInlineCharacterStyles(field.CachedResult, identities);
                    break;
            }
        }
    }

    /// <summary>Prove that a used style and every same-type ancestor can be safely materialized. Existing
    /// property history must remain verbatim, and numbered styles need their own label-level revision model,
    /// so either condition deliberately excludes the chain from this first projection slice.</summary>
    private static bool StyleChainIsSharedAndResolvable(
        XElement leftStyles, XElement rightStyles, StyleIdentity identity)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        string? current = identity.Id;
        for (int depth = 0; current is not null && depth < 16; depth++)
        {
            if (!seen.Add(current))
                return false;
            var left = leftStyles.Elements(W.style).Where(style =>
                string.Equals((string?)style.Attribute(W.type), identity.Type, StringComparison.Ordinal) &&
                string.Equals((string?)style.Attribute(W.styleId), current, StringComparison.Ordinal)).ToList();
            var right = rightStyles.Elements(W.style).Where(style =>
                string.Equals((string?)style.Attribute(W.type), identity.Type, StringComparison.Ordinal) &&
                string.Equals((string?)style.Attribute(W.styleId), current, StringComparison.Ordinal)).ToList();
            if (left.Count != 1 || right.Count != 1 ||
                !StyleCascadeMetadataEqual(left[0], right[0]) ||
                HasStylePropertyRevisions(left[0]) || HasStylePropertyRevisions(right[0]) ||
                left[0].Descendants(W.numPr).Any() || right[0].Descendants(W.numPr).Any())
                return false;
            current = (string?)left[0].Element(W.basedOn)?.Attribute(W.val);
        }
        return current is null;
    }

    /// <summary>
    /// Return the special Word-style provenance mode only when the comparison is a literal full
    /// replacement of the main body, the LEFT has neither docDefaults nor a default paragraph style,
    /// and every paragraph in every LEFT story names a fully resolvable LEFT paragraph-style chain.
    /// Those conditions make a right-only default style observationally unreachable after reject, so
    /// its tracked style projection cannot alter LEFT body content or direct formatting.
    /// </summary>
    private static InsertedStyleNormalization? TryCreateInsertedStyleNormalization(
        IrEditScript script, RenderState state, MainDocumentPart main, MainDocumentPart? rightMain)
    {
        // Notes and headers/footers are global-style consumers too. Keep this compatibility path
        // main-body-only until their right-only style provenance has a separately proven projection.
        // The renderer preserves the glossary part untouched. Until its style reachability is
        // explicitly modeled, a right default must not be allowed to change its rejected view.
        if (main.GlossaryDocumentPart is not null ||
            script.NoteOps is not null || script.HeaderFooterOps is not null ||
            !IsPureFullBodyReplacement(script, state) ||
            !LeftLacksDefaultsAndDefaultParagraphStyle(main) ||
            !AllLeftStoryParagraphStylesResolveWithinLeftStyles(main))
            return null;

        var directlyUsedStyleIds = new HashSet<string>(StringComparer.Ordinal);
        bool usesImplicitDefault = false;
        foreach (var op in script.Operations)
        {
            if (op.Kind != IrEditOpKind.InsertBlock || IsSectionBreakOp(op, state))
                continue;
            var source = SourceElement(op.RightAnchor, state.RightSource);
            if (source is null)
                return null; // Provenance is load-bearing for the guard as well as the renderer.
            foreach (var paragraph in source.DescendantsAndSelf(W.p))
            {
                var styleId = (string?)paragraph.Element(W.pPr)?.Element(W.pStyle)?.Attribute(W.val);
                if (string.IsNullOrEmpty(styleId))
                    usesImplicitDefault = true;
                else
                    directlyUsedStyleIds.Add(styleId);
            }
        }

        if (directlyUsedStyleIds.Count == 0 && !usesImplicitDefault)
            return null;

        // A direct pStyle can inherit all visible formatting from right-only ancestors. Resolve the
        // whole same-type chain now, while the untouched RIGHT styles part is available; ambiguity,
        // a missing node, a cycle, or a cross-type hop makes this special projection unsafe.
        var leftStylesRoot = main.StyleDefinitionsPart?.GetXDocument().Root;
        var rightStylesRoot = rightMain?.StyleDefinitionsPart?.GetXDocument().Root;
        return leftStylesRoot is not null && rightStylesRoot is not null &&
            TryResolveUsedRightParagraphStyleClosure(
                leftStylesRoot, rightStylesRoot, directlyUsedStyleIds, usesImplicitDefault, out var usedStyleIds)
            ? new InsertedStyleNormalization(usedStyleIds)
            : null;
    }

    private static bool TryResolveUsedRightParagraphStyleClosure(
        XElement leftStylesRoot, XElement stylesRoot, IReadOnlyCollection<string> directlyUsedStyleIds,
        bool usesImplicitDefault,
        out HashSet<string> usedStyleIds)
    {
        usedStyleIds = new HashSet<string>(StringComparer.Ordinal);

        // Style ids are package-global. Refuse an ambiguous id even if its duplicate is a different
        // type: a malformed package can otherwise bind a basedOn edge differently after style copy.
        var stylesById = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var style in stylesRoot.Elements(W.style))
        {
            var styleId = (string?)style.Attribute(W.styleId);
            if (string.IsNullOrEmpty(styleId))
                continue;
            if (!stylesById.TryAdd(styleId, style))
                return false;
        }

        var roots = new HashSet<string>(directlyUsedStyleIds, StringComparer.Ordinal);
        if (usesImplicitDefault)
        {
            var defaults = stylesById
                .Where(pair => string.Equals((string?)pair.Value.Attribute(W.type), "paragraph", StringComparison.Ordinal) &&
                    IsOn((string?)pair.Value.Attribute(W._default)))
                .Select(pair => pair.Key)
                .ToList();
            if (defaults.Count != 1)
                return false;
            roots.Add(defaults[0]);
        }

        foreach (var styleId in roots)
            if (!TryAddRightParagraphStyleChain(styleId, stylesById, usedStyleIds))
                return false;

        // The general style merger matches by (type, styleId), but styleId itself is package-global.
        // A right paragraph style that collides with any LEFT type would be appended alongside the
        // left definition. Decline the exceptional projection rather than making that ambiguity
        // carry a retained default or synthesized property revisions.
        var leftStyleIds = leftStylesRoot.Elements(W.style)
            .Select(style => (string?)style.Attribute(W.styleId))
            .Where(styleId => !string.IsNullOrEmpty(styleId))
            .Cast<string>()
            .ToHashSet(StringComparer.Ordinal);
        if (usedStyleIds.Overlaps(leftStyleIds))
            return false;
        return true;
    }

    private static bool TryAddRightParagraphStyleChain(
        string styleId, IReadOnlyDictionary<string, XElement> stylesById, HashSet<string> usedStyleIds)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var currentId = styleId;
        while (true)
        {
            if (!seen.Add(currentId) || !stylesById.TryGetValue(currentId, out var style) ||
                !string.Equals((string?)style.Attribute(W.type), "paragraph", StringComparison.Ordinal))
                return false;

            usedStyleIds.Add(currentId);
            var basedOn = style.Element(W.basedOn);
            if (basedOn is null)
                return true;

            var basedOnStyleId = (string?)basedOn.Attribute(W.val);
            if (string.IsNullOrEmpty(basedOnStyleId))
                return false;
            currentId = basedOnStyleId;
        }
    }

    /// <summary>Strictly prove that every non-section main-body block is either deleted from LEFT or
    /// inserted from RIGHT, exactly once. A paired/modified/moved block is intentionally disqualifying:
    /// it could still observe a style on both accept and reject.</summary>
    private static bool IsPureFullBodyReplacement(IrEditScript script, RenderState state)
    {
        var leftAnchors = state.Left.Body.Blocks
            .Where(b => b is not IrSectionBreak)
            .Select(b => b.Anchor.ToString())
            .ToHashSet(StringComparer.Ordinal);
        var rightAnchors = state.RightSource.Body.Blocks
            .Where(b => b is not IrSectionBreak)
            .Select(b => b.Anchor.ToString())
            .ToHashSet(StringComparer.Ordinal);
        if (leftAnchors.Count == 0 || rightAnchors.Count == 0)
            return false;

        var deleted = new HashSet<string>(StringComparer.Ordinal);
        var inserted = new HashSet<string>(StringComparer.Ordinal);
        foreach (var op in script.Operations)
        {
            if (IsSectionBreakOp(op, state))
                continue;
            switch (op.Kind)
            {
                case IrEditOpKind.DeleteBlock when op.LeftAnchor is { } left && op.RightAnchor is null:
                    if (!leftAnchors.Contains(left) || !deleted.Add(left))
                        return false;
                    break;
                case IrEditOpKind.InsertBlock when op.RightAnchor is { } right && op.LeftAnchor is null:
                    if (!rightAnchors.Contains(right) || !inserted.Add(right))
                        return false;
                    break;
                default:
                    return false;
            }
        }
        return deleted.SetEquals(leftAnchors) && inserted.SetEquals(rightAnchors);
    }

    /// <summary>The compatibility projection is valid only when the left styles part is present but
    /// has neither docDefaults nor a default paragraph style. Missing a styles part takes the existing
    /// stock-default backfill path instead; it has no right-style copy surface to normalize.</summary>
    private static bool LeftLacksDefaultsAndDefaultParagraphStyle(MainDocumentPart main)
    {
        var root = main.StyleDefinitionsPart?.GetXDocument().Root;
        return root is not null && root.Element(W.docDefaults) is null &&
            !root.Elements(W.style).Any(style =>
                ((string?)style.Attribute(W.type) is null or "paragraph") &&
                IsOn((string?)style.Attribute(W._default)));
    }

    /// <summary>Scan the same text stories the reader/revision processor understands. Every LEFT
    /// paragraph style must resolve entirely within the original LEFT styles part: a missing node,
    /// cycle, type mismatch, malformed basedOn, or default-style dependency could otherwise become
    /// observable when the right style universe is copied into the rejected package.</summary>
    private static bool AllLeftStoryParagraphStylesResolveWithinLeftStyles(MainDocumentPart main)
    {
        var stylesRoot = main.StyleDefinitionsPart?.GetXDocument().Root;
        if (stylesRoot is null)
            return false;

        // Style ids are package-global across types. Ambiguous duplicate ids could resolve to a
        // different definition after right-only styles are copied, so this deliberately declines.
        var stylesById = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var style in stylesRoot.Elements(W.style))
        {
            var styleId = (string?)style.Attribute(W.styleId);
            if (string.IsNullOrEmpty(styleId))
                continue;
            if (!stylesById.TryAdd(styleId, style))
                return false;
        }

        foreach (var part in StyleSafetyStoryParts(main))
        {
            var root = part.GetXDocument().Root;
            if (root is null)
                return false;
            foreach (var paragraph in root.DescendantsAndSelf(W.p))
            {
                var styleId = (string?)paragraph.Element(W.pPr)?.Element(W.pStyle)?.Attribute(W.val);
                if (string.IsNullOrEmpty(styleId) ||
                    !ResolvesLeftParagraphStyleChain(styleId, stylesById))
                    return false;
            }
        }
        return true;
    }

    private static bool ResolvesLeftParagraphStyleChain(
        string styleId, IReadOnlyDictionary<string, XElement> stylesById)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var currentId = styleId;
        while (true)
        {
            if (!seen.Add(currentId) || !stylesById.TryGetValue(currentId, out var style) ||
                !string.Equals((string?)style.Attribute(W.type), "paragraph", StringComparison.Ordinal) ||
                IsOn((string?)style.Attribute(W._default)))
                return false;

            var basedOn = style.Element(W.basedOn);
            if (basedOn is null)
                return true;

            var basedOnStyleId = (string?)basedOn.Attribute(W.val);
            if (string.IsNullOrEmpty(basedOnStyleId))
                return false;
            currentId = basedOnStyleId;
        }
    }

    private static IEnumerable<OpenXmlPart> StyleSafetyStoryParts(MainDocumentPart main)
    {
        yield return main;
        foreach (var header in main.HeaderParts)
            yield return header;
        foreach (var footer in main.FooterParts)
            yield return footer;
        if (main.FootnotesPart is not null)
            yield return main.FootnotesPart;
        if (main.EndnotesPart is not null)
            yield return main.EndnotesPart;
        if (main.WordprocessingCommentsPart is not null)
            yield return main.WordprocessingCommentsPart;
    }

    private static bool IsOn(string? value) =>
        string.Equals(value, "1", StringComparison.Ordinal) ||
        string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "on", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Mirror Word compare's style-definition treatment into the LEFT-based output package (decoded
    /// from the Word-compare oracle corpus — see DocxDiffStyleProvenanceTests): keep the left styles
    /// part (docDefaults et al.), copy right-only styles, and rewrite every shared style whose RAW
    /// definition formatting differs (rsid noise ignored) to the right's EFFECTIVE formatting
    /// (docDefaults + basedOn chain + own definition), with the left effective payload archived in a
    /// tracked <c>w:rPrChange</c>/<c>w:pPrChange</c>. A narrowly proven docDefaults-only plan applies
    /// the same reversible projection to eligible raw-equal styles. This lets the redline render with
    /// the revised look while Reject restores the original, without copying an untrackable package
    /// part. No-ops gracefully when either side lacks a styles part.
    /// </summary>
    private static HashSet<StyleIdentity> TrackStyleDefinitionChanges(
        WordprocessingDocument wDoc, WordprocessingDocument wDocRight, RenderState state,
        InsertedStyleNormalization? insertedStyleNormalization, bool leftHadTheme,
        DocDefaultsStyleProjection? docDefaultsStyleProjection)
    {
        var rightImportedStyles = new HashSet<StyleIdentity>();
        if (wDoc.MainDocumentPart?.StyleDefinitionsPart is not { } leftStyles ||
            wDocRight.MainDocumentPart?.StyleDefinitionsPart is not { } rightStyles)
            return rightImportedStyles;

        var outXDoc = leftStyles.GetXDocument();
        if (outXDoc.Root is not { } root || rightStyles.GetXDocument().Root is not { } rightRoot)
            return rightImportedStyles;
        var leftOriginalRoot = new XElement(root);   // frozen snapshot for left-effective resolution
        var stockDocDefaults = insertedStyleNormalization is null
            ? null
            : XElement.Parse(leftHadTheme ? WordStockDocDefaults.ClassicXml : WordStockDocDefaults.ModernXml);

        foreach (var rightStyle in rightRoot.Elements(W.style))
        {
            var type = (string?)rightStyle.Attribute(W.type);
            var styleId = (string?)rightStyle.Attribute(W.styleId);
            var leftStyle = root.Elements(W.style).FirstOrDefault(st =>
                (string?)st.Attribute(W.type) == type &&
                (string?)st.Attribute(W.styleId) == styleId);
            if (leftStyle is null)
            {
                var cloned = new XElement(rightStyle);
                bool usedInsertedParagraphStyle = styleId is not null && type == "paragraph" &&
                    insertedStyleNormalization is not null &&
                    insertedStyleNormalization.Uses(styleId);
                // RawStylePayload deliberately strips these markers; never send a pre-existing input
                // style revision through that path. Keep this individual used style verbatim while
                // still normalizing other independently-safe members of the same basedOn closure.
                bool preservesInputPropertyRevisions = usedInsertedParagraphStyle &&
                    HasStylePropertyRevisions(rightStyle);
                if (usedInsertedParagraphStyle && !preservesInputPropertyRevisions)
                {
                    NormalizeInsertedParagraphStyle(cloned, rightStyle, stockDocDefaults!, state);
                }
                else if (!preservesInputPropertyRevisions)
                {
                    // In the general path the LEFT's default paragraph style remains authoritative.
                    // The exceptional normalized path above retains a right default only after proving
                    // every LEFT story paragraph names a style explicitly. A used default with an
                    // input property revision is also retained: stripping its default flag would
                    // change the source revision's reachability before it can be preserved.
                    cloned.Attribute(W._default)?.Remove();
                }
                root.Add(cloned);
                if (type is not null && styleId is not null)
                    rightImportedStyles.Add(new StyleIdentity(type, styleId));
                continue;
            }
            if (styleId is null)
                continue;

            if (StyleDefinitionPayloadsEqual(leftStyle, rightStyle))
            {
                if (docDefaultsStyleProjection?.Includes(type, styleId) == true &&
                    StyleCascadeMetadataEqual(leftStyle, rightStyle))
                    ApplyDocDefaultsStyleProjection(leftStyle, leftOriginalRoot, rightRoot, type, styleId, state);
                continue;
            }

            // Payload provenance decoded from the oracle: the PARAGRAPH side stays at RAW definition
            // payloads — Word's updated Normal carries an EMPTY current pPr (docDefaults spacing is
            // NOT materialized into the style; doing so outranks table-style conditional spacing and
            // inflates every styled table's rows) — while the RUN side materializes the resolved
            // (docDefaults + basedOn chain) fonts/size, exactly as Word writes them.
            var rightRawPPr = RawStylePayload(W.pPr, rightStyle);
            var leftRawPPr = RawStylePayload(W.pPr, leftStyle);
            var (_, rightRPr) = ResolveEffectiveStyleFormatting(rightRoot, type, styleId);
            var (_, leftRPr) = ResolveEffectiveStyleFormatting(leftOriginalRoot, type, styleId);

            leftStyle.Elements(W.pPr).Remove();
            leftStyle.Elements(W.rPr).Remove();
            var newPPr = new XElement(W.pPr, rightRawPPr.Elements(),
                new XElement(W.pPrChange, state.RevisionAttributes(),
                    new XElement(W.pPr, leftRawPPr.Elements())));
            var newRPr = new XElement(W.rPr, rightRPr.Elements(),
                new XElement(W.rPrChange, state.RevisionAttributes(),
                    new XElement(W.rPr, leftRPr.Elements())));
            // Schema: pPr precedes rPr, both after the leading metadata children and before any
            // table-style children (w:tblPr/w:trPr/w:tcPr/w:tblStylePr).
            var anchor = leftStyle.Elements().FirstOrDefault(e =>
                e.Name == W.tblPr || e.Name == W.trPr || e.Name == W.tcPr || e.Name == W.tblStylePr);
            if (anchor is null)
            {
                leftStyle.Add(newPPr);
                leftStyle.Add(newRPr);
            }
            else
            {
                anchor.AddBeforeSelf(newPPr);
                anchor.AddBeforeSelf(newRPr);
            }
        }
        leftStyles.PutXDocument();
        return rightImportedStyles;
    }

    /// <summary>
    /// Materialize a docDefaults-only presentation delta into one shared style definition. The package still
    /// carries LEFT docDefaults, so the current effective properties must be made direct on the style for the
    /// accepted view; the LEFT effective properties inside the native property-change marker restore Reject.
    /// Character styles can own run properties only, while paragraph styles safely own both slices.
    /// </summary>
    private static void ApplyDocDefaultsStyleProjection(
        XElement outputStyle, XElement leftStylesRoot, XElement rightStylesRoot,
        string? type, string styleId, RenderState state)
    {
        var (leftPPr, leftRPr) = ResolveEffectiveStyleFormatting(leftStylesRoot, type, styleId);
        var (rightPPr, rightRPr) = ResolveEffectiveStyleFormatting(rightStylesRoot, type, styleId);
        NormalizeStylePropertyOrder(leftPPr, StylePPrChildOrder);
        NormalizeStylePropertyOrder(rightPPr, StylePPrChildOrder);
        NormalizeStylePropertyOrder(leftRPr, StyleRPrChildOrder);
        NormalizeStylePropertyOrder(rightRPr, StyleRPrChildOrder);

        bool projectPPr = string.Equals(type, "paragraph", StringComparison.Ordinal) &&
            !XNode.DeepEquals(leftPPr, rightPPr);
        bool projectRPr = (string.Equals(type, "paragraph", StringComparison.Ordinal) ||
                           string.Equals(type, "character", StringComparison.Ordinal)) &&
            !XNode.DeepEquals(leftRPr, rightRPr);
        if (!projectPPr && !projectRPr)
            return;

        XElement? currentPPr = null;
        if (projectPPr)
        {
            currentPPr = new XElement(rightPPr);
            currentPPr.Add(new XElement(W.pPrChange, state.RevisionAttributes(), new XElement(leftPPr)));
        }

        XElement? currentRPr = null;
        if (projectRPr)
        {
            currentRPr = new XElement(rightRPr);
            currentRPr.Add(new XElement(W.rPrChange, state.RevisionAttributes(), new XElement(leftRPr)));
        }

        ReplaceStyleProperties(outputStyle, currentPPr, currentRPr);
    }

    /// <summary>Replace only the style property slices supplied by a presentation projection, preserving the
    /// style metadata and the schema-required pPr → rPr → table-property ordering.</summary>
    private static void ReplaceStyleProperties(XElement style, XElement? pPr, XElement? rPr)
    {
        if (pPr is not null)
            style.Elements(W.pPr).Remove();
        if (rPr is not null)
            style.Elements(W.rPr).Remove();

        if (pPr is not null)
        {
            var pPrAnchor = style.Elements().FirstOrDefault(e =>
                e.Name == W.rPr || e.Name == W.tblPr || e.Name == W.trPr || e.Name == W.tcPr ||
                e.Name == W.tblStylePr);
            if (pPrAnchor is null)
                style.Add(pPr);
            else
                pPrAnchor.AddBeforeSelf(pPr);
        }

        if (rPr is not null)
        {
            var rPrAnchor = style.Elements().FirstOrDefault(e =>
                e.Name == W.tblPr || e.Name == W.trPr || e.Name == W.tcPr || e.Name == W.tblStylePr);
            if (rPrAnchor is null)
                style.Add(rPr);
            else
                rPrAnchor.AddBeforeSelf(rPr);
        }
    }

    private static void NormalizeStylePropertyOrder(XElement props, string[] order)
    {
        var children = props.Elements().Select((element, index) => new
        {
            Element = element,
            Index = index,
            Rank = element.Name.Namespace == W.w
                ? System.Array.IndexOf(order, element.Name.LocalName)
                : -1,
        }).OrderBy(item => item.Rank >= 0 ? item.Rank : 10_000 + item.Index).ToList();
        foreach (var child in children)
            child.Element.Remove();
        props.Add(children.Select(child => child.Element));
    }

    /// <summary>Raw-equivalent pPr/rPr payloads are not enough to prove a docDefaults-only change: a changed
    /// basedOn edge would also change effective formatting but needs its own semantic style diff. The first
    /// presentation slice requires the cascade edge to be identical on both sides.</summary>
    private static bool StyleCascadeMetadataEqual(XElement leftStyle, XElement rightStyle) =>
        string.Equals((string?)leftStyle.Element(W.basedOn)?.Attribute(W.val),
            (string?)rightStyle.Element(W.basedOn)?.Attribute(W.val), StringComparison.Ordinal);

    private static bool HasStylePropertyRevisions(XElement style) =>
        style.Descendants(W.pPrChange).Any() || style.Descendants(W.rPrChange).Any();

    /// <summary>
    /// Project a copied, right-only paragraph style as Word does for the guarded total-replacement
    /// shape. The body keeps its source pPr/rPr verbatim; only the otherwise unreachable style
    /// definition is made explicit against the LEFT stock defaults, with the raw source payload kept
    /// in tracked property history. That keeps accept/reject content and direct formatting unchanged.
    /// </summary>
    private static void NormalizeInsertedParagraphStyle(
        XElement outputStyle, XElement rightStyle, XElement stockDocDefaults, RenderState state)
    {
        var sourcePPr = RawStylePayload(W.pPr, rightStyle);
        var sourceRPr = RawStylePayload(W.rPr, rightStyle);
        var currentPPr = new XElement(sourcePPr);
        var currentRPr = new XElement(sourceRPr);
        var stockPPr = StockStyleProperties(stockDocDefaults, W.pPr);
        var stockRPr = StockStyleProperties(stockDocDefaults, W.rPr);

        // Word's imported style payloads pin an auto 12pt line even when the right source delegated
        // spacing to its own empty docDefaults. Preserve real before/after values from the source;
        // only supply a zero after-value when it was inherited. This is the layout-critical piece of
        // the malformed pgsz → uiPriority projection.
        EnsureCompactStyleSpacing(currentPPr);
        NormalizeInsertedStyleRunProperties(currentRPr, stockRPr);

        bool isDefaultParagraphStyle = IsOn((string?)rightStyle.Attribute(W._default));
        // For the right default, the old payload is the left package's stock defaults — precisely the
        // values this projection supersedes. Other right-only styles archive their direct source
        // payload; after reject every one is unreachable by the guard, but retaining it makes the
        // change history honest and inspectable.
        var previousPPr = isDefaultParagraphStyle ? stockPPr : sourcePPr;
        var previousRPr = isDefaultParagraphStyle ? stockRPr : sourceRPr;
        if (!XNode.DeepEquals(currentPPr, previousPPr))
            currentPPr.Add(new XElement(W.pPrChange, state.RevisionAttributes(), new XElement(previousPPr)));
        if (!XNode.DeepEquals(currentRPr, previousRPr))
            currentRPr.Add(new XElement(W.rPrChange, state.RevisionAttributes(), new XElement(previousRPr)));

        outputStyle.Elements(W.pPr).Remove();
        outputStyle.Elements(W.rPr).Remove();
        var anchor = outputStyle.Elements().FirstOrDefault(e =>
            e.Name == W.tblPr || e.Name == W.trPr || e.Name == W.tcPr || e.Name == W.tblStylePr);
        if (anchor is null)
        {
            outputStyle.Add(currentPPr);
            outputStyle.Add(currentRPr);
        }
        else
        {
            anchor.AddBeforeSelf(currentPPr);
            anchor.AddBeforeSelf(currentRPr);
        }
    }

    private static XElement StockStyleProperties(XElement docDefaults, XName propertyName)
    {
        var properties = propertyName == W.pPr
            ? docDefaults.Element(W.pPrDefault)?.Element(W.pPr)
            : docDefaults.Element(W.rPrDefault)?.Element(W.rPr);
        return new XElement(propertyName, properties?.Elements());
    }

    /// <summary>Ensure the compact Word-style line spacing without replacing source direct pPr
    /// facts such as heading before/after spacing, outline level, shading, or indentation.</summary>
    private static void EnsureCompactStyleSpacing(XElement pPr)
    {
        var spacing = pPr.Element(W.spacing);
        if (spacing is null)
        {
            spacing = new XElement(W.spacing,
                new XAttribute(W.after, 0),
                new XAttribute(W.line, 240),
                new XAttribute(W.lineRule, "auto"));
            // Insert before the first present CT_PPrBase child ordered after w:spacing (or any
            // foreign-namespace extension, which likewise belongs after the standard sequence). In
            // particular w:ind/w:jc commonly occur without explicit spacing; appending in that
            // shape violates schema order and makes Word repair the styles part on open.
            var tail = pPr.Elements().FirstOrDefault(e =>
                PPrChildrenAfterSpacing.Contains(e.Name) || e.Name.Namespace != W.w);
            if (tail is null)
                pPr.Add(spacing);
            else
                tail.AddBeforeSelf(spacing);
        }
        else
        {
            if (spacing.Attribute(W.after) is null)
                spacing.SetAttributeValue(W.after, 0);
            spacing.SetAttributeValue(W.line, 240);
            spacing.SetAttributeValue(W.lineRule, "auto");
        }

        // Word writes the schema-default color explicitly when it materializes a shaded style.
        // This is rendering-neutral but makes the copied style deterministic across consumers.
        var shading = pPr.Element(W.shd);
        if (shading is not null && shading.Attribute(W.color) is null)
            shading.SetAttributeValue(W.color, "auto");
    }

    /// <summary>
    /// Materialize only the run-format facts that differ from the stock left defaults. A right
    /// 12-point style inherits size from stock, while an explicit larger heading/title keeps its
    /// size. Modern stock has kern=2, whereas Word projects these imported source styles with
    /// kern=0.
    /// </summary>
    private static void NormalizeInsertedStyleRunProperties(XElement rPr, XElement stockRPr)
    {
        var stockSize = (string?)stockRPr.Element(W.sz)?.Attribute(W.val);
        var stockSizeCs = (string?)stockRPr.Element(W.szCs)?.Attribute(W.val);
        var size = rPr.Element(W.sz);
        var sizeCs = rPr.Element(W.szCs);
        if (size is not null && sizeCs is not null &&
            string.Equals((string?)size.Attribute(W.val), stockSize, StringComparison.Ordinal) &&
            string.Equals((string?)sizeCs.Attribute(W.val), stockSizeCs, StringComparison.Ordinal))
        {
            size.Remove();
            sizeCs.Remove();
        }

        // Only override kerning when the stock defaults supplied a nonzero baseline. An explicit
        // source kern setting is direct style formatting and must remain authoritative.
        if (stockRPr.Element(W.kern) is not null && rPr.Element(W.kern) is null)
        {
            var kern = new XElement(W.kern, new XAttribute(W.val, 0));
            // Extensions such as w14:ligatures belong after the standard CT_RPr sequence, so they
            // are an insertion tail too. Appending kern after one is just as schema-invalid as
            // appending it after w:lang.
            var tail = rPr.Elements().FirstOrDefault(e =>
                RPrChildrenAfterKern.Contains(e.Name) || e.Name.Namespace != W.w);
            if (tail is null)
                rPr.Add(kern);
            else
                tail.AddBeforeSelf(kern);
        }

    }

    /// <summary>
    /// Synthesize numbering definitions for body <c>w:numId</c> references that resolve to nothing —
    /// Microsoft Word's own dangling-numId repair (verified against its compare oracle output): a
    /// decimal multilevel abstract (lvlText "%N.", 720-twip-per-level hanging indents) plus a
    /// <c>w:num</c> mapping the dangling id onto it. No-op when every referenced id is defined.
    /// </summary>
    private static void RepairDanglingNumberingReferences(MainDocumentPart main)
    {
        var body = main.GetXDocument().Root?.Element(W.body);
        if (body is null)
            return;
        var referenced = body.Descendants(W.numId)
            .Select(e => (string?)e.Attribute(W.val))
            .Where(v => !string.IsNullOrEmpty(v) && v != "0")
            .Select(v => v!)
            .ToHashSet(StringComparer.Ordinal);
        if (referenced.Count == 0)
            return;

        var numberingPart = main.NumberingDefinitionsPart;
        var defined = numberingPart?.GetXDocument().Root?.Elements(W.num)
            .Select(n => (string?)n.Attribute(W.numId))
            .Where(v => v is not null)
            .Select(v => v!)
            .ToHashSet(StringComparer.Ordinal) ?? new HashSet<string>(StringComparer.Ordinal);
        var dangling = referenced.Except(defined).OrderBy(v => v, StringComparer.Ordinal).ToList();
        if (dangling.Count == 0)
            return;

        numberingPart ??= main.AddNewPart<NumberingDefinitionsPart>();
        var numXDoc = numberingPart.GetXDocument();
        if (numXDoc.Root is null)
            numXDoc.Add(new XElement(W.numbering,
                new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName)));
        var root = numXDoc.Root!;

        int nextAbstract = root.Elements(W.abstractNum)
            .Select(a => int.TryParse((string?)a.Attribute(W.abstractNumId), out var id) ? id : -1)
            .DefaultIfEmpty(-1)
            .Max() + 1;
        // Schema order inside w:numbering: numPicBullet*, abstractNum*, num* — new abstracts go
        // before the first existing w:num; the num mappings append at the end.
        var firstNum = root.Elements(W.num).FirstOrDefault();
        foreach (var id in dangling)
        {
            var abstractNum = new XElement(W.abstractNum,
                new XAttribute(W.abstractNumId, nextAbstract),
                new XElement(W.multiLevelType, new XAttribute(W.val, "multilevel")),
                Enumerable.Range(0, 9).Select(i => new XElement(W.lvl,
                    new XAttribute(W.ilvl, i),
                    new XElement(W.start, new XAttribute(W.val, 1)),
                    new XElement(W.numFmt, new XAttribute(W.val, "decimal")),
                    new XElement(W.lvlText, new XAttribute(W.val, $"%{i + 1}.")),
                    new XElement(W.lvlJc, new XAttribute(W.val, "left")),
                    new XElement(W.pPr,
                        new XElement(W.ind,
                            new XAttribute(W.left, 720 * (i + 1)),
                            new XAttribute(W.hanging, 720))))));
            if (firstNum is null)
                root.Add(abstractNum);
            else
                firstNum.AddBeforeSelf(abstractNum);
            root.Add(new XElement(W.num,
                new XAttribute(W.numId, id),
                new XElement(W.abstractNumId, new XAttribute(W.val, nextAbstract))));
            nextAbstract++;
        }
        numberingPart.PutXDocument();
    }

    /// <summary>True when any story part that <see cref="RevisionProcessor"/> considers carries tracked
    /// revision markup. The scan is namespace-aware so an alternate prefix for WordprocessingML cannot evade
    /// the one-sided input-revision-preservation gate. It deliberately covers the main document plus headers,
    /// footers, endnotes, and footnotes — a dirty header is just as asymmetric as a dirty body when only the
    /// RIGHT input's revisions would otherwise be preserved.</summary>
    private static bool HasTrackedRevisionMarkup(WmlDocument doc)
    {
        try
        {
            using var ms = new MemoryStream(doc.DocumentByteArray, writable: false);
            using var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read);
            foreach (var entry in zip.Entries)
            {
                // Mirrors RevisionProcessor.HasTrackedRevisions' story-part scope without opening the
                // document through the SDK before the renderer's package pass. Header/footer part names are
                // generated as word/header*.xml and word/footer*.xml by Word/Open XML SDK; scanning every
                // such part is conservative for an orphaned relationship and never risks asymmetric output.
                bool trackedStoryPart = entry.FullName is "word/document.xml" or "word/endnotes.xml" or "word/footnotes.xml" ||
                    (entry.FullName.StartsWith("word/header", StringComparison.Ordinal) &&
                     entry.FullName.EndsWith(".xml", StringComparison.Ordinal)) ||
                    (entry.FullName.StartsWith("word/footer", StringComparison.Ordinal) &&
                     entry.FullName.EndsWith(".xml", StringComparison.Ordinal));
                if (!trackedStoryPart)
                    continue;

                using var stream = entry.Open();
                using var reader = System.Xml.XmlReader.Create(stream);
                while (reader.Read())
                {
                    if (reader.NodeType == System.Xml.XmlNodeType.Element &&
                        TrackedRevisionNames.Contains(XName.Get(reader.LocalName, reader.NamespaceURI)))
                        return true;
                }
            }
            return false;
        }
        catch (Exception e) when (e is InvalidDataException or System.Xml.XmlException)
        {
            return false;
        }
    }

    /// <summary>Style ids defined in a document's styles part, read without opening the package
    /// through the SDK (called before the render's package pass).</summary>
    private static HashSet<string> ReadStyleIds(WmlDocument doc)
    {
        var ids = new HashSet<string>(StringComparer.Ordinal);
        try
        {
            using var ms = new MemoryStream(doc.DocumentByteArray, writable: false);
            using var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read);
            var entry = zip.GetEntry("word/styles.xml");
            if (entry is null)
                return ids;
            using var stream = entry.Open();
            var root = XDocument.Load(stream).Root;
            if (root is null)
                return ids;
            foreach (var style in root.Elements(W.style))
                if ((string?)style.Attribute(W.styleId) is { } id)
                    ids.Add(id);
        }
        catch (Exception e) when (e is InvalidDataException or System.Xml.XmlException)
        {
            // Malformed package/part: leave the set as-is; the drop check degrades to a no-op set.
        }
        return ids;
    }

    /// <summary>Drop a stamped-current pPr's <c>w:pStyle</c> when the referenced style is not
    /// defined in the LEFT styles part. Word expresses a PAIRED paragraph's format change within
    /// the left style universe — an unresolvable style reference is dropped and the delta lives in
    /// direct properties (oracle-verified: the output styles part never gains the right-only style
    /// for a paired paragraph; only wholly-inserted paragraphs import their styles).</summary>
    private static void DropUnresolvableStyleRef(XElement pPr, RenderState state)
    {
        if (state.LeftStyleIds is not { } known)
            return;
        var pStyle = pPr.Element(W.pStyle);
        if (pStyle is not null && (string?)pStyle.Attribute(W.val) is { } id && !known.Contains(id))
            pStyle.Remove();
    }

    /// <summary>
    /// Remove paragraph-style references in the rendered main body that cannot resolve against the
    /// FINAL paragraph-style registry. This closes the verbatim EqualBlock and archived
    /// <c>w:pPrChange</c> paths, which have no paired-paragraph context in which to invoke
    /// <see cref="DropUnresolvableStyleRef"/>. Restricting the lookup to paragraph styles is
    /// intentional: a matching character/table style does not make a <c>w:pStyle</c> valid.
    /// </summary>
    private static void DropDanglingParagraphStyleRefs(MainDocumentPart main)
    {
        var body = main.GetXDocument().Root?.Element(W.body);
        if (body is null)
            return;

        var knownParagraphStyles = main.StyleDefinitionsPart?.GetXDocument().Root?
            .Elements(W.style)
            .Where(style => (string?)style.Attribute(W.type) == "paragraph")
            .Select(style => (string?)style.Attribute(W.styleId))
            .Where(id => !string.IsNullOrEmpty(id))
            .Select(id => id!)
            .ToHashSet(StringComparer.Ordinal) ?? new HashSet<string>(StringComparer.Ordinal);

        var dangling = body.Descendants(W.pStyle)
            .Where(pStyle => (string?)pStyle.Attribute(W.val) is not { } id ||
                !knownParagraphStyles.Contains(id))
            .ToList();
        if (dangling.Count == 0)
            return;

        foreach (var pStyle in dangling)
            pStyle.Remove();
        main.PutXDocument();
    }

    /// <summary>
    /// Rebind live RIGHT-sourced paragraph numbering to definitions that collision handling imported under a
    /// fresh id.  An equal paragraph can still be semantically changed when its shared <c>w:numId</c> resolves
    /// through a different numbering definition, so its current properties take the imported id and its left
    /// properties are preserved in <c>w:pPrChange</c>.  Accept therefore resolves the right definition and
    /// reject restores the left definition.  Deleted/move-from paragraphs and archived <c>*Change</c> payloads
    /// remain left-sourced.  Covers the main document, headers, footers, footnotes, and endnotes.
    /// </summary>
    private static void RebindRightNumberingReferences(
        MainDocumentPart main, Dictionary<int, int> numIdMap, RenderState state)
    {
        if (numIdMap.Count == 0)
            return;
        var parts = new List<OpenXmlPart> { main };
        parts.AddRange(main.HeaderParts);
        parts.AddRange(main.FooterParts);
        if (main.FootnotesPart is not null)
            parts.Add(main.FootnotesPart);
        if (main.EndnotesPart is not null)
            parts.Add(main.EndnotesPart);
        foreach (var part in parts)
        {
            var xDoc = part.GetXDocument();
            var changed = false;
            foreach (var paragraph in xDoc.Descendants(W.p).ToList())
            {
                if (IsDeletedParagraph(paragraph))
                    continue;
                var pPr = paragraph.Element(W.pPr);
                var numIdEl = pPr?.Element(W.numPr)?.Element(W.numId);
                if (numIdEl is null ||
                    !int.TryParse((string?)numIdEl.Attribute(W.val), out var id) ||
                    !numIdMap.TryGetValue(id, out var mapped))
                    continue;

                // An inserted/move-to paragraph disappears on reject.  Every other live paragraph needs a
                // standard pPr history unless another formatting pass already supplied one; in that case its
                // archived pPr is the left payload and must retain the original numId.
                XElement? oldPPr = null;
                if (state.Settings.TrackParagraphFormatChanges &&
                    !IsInsertedParagraph(paragraph) &&
                    pPr!.Element(W.pPrChange) is null)
                {
                    oldPPr = StripUnids(new XElement(W.pPr, pPr.Attributes(),
                        pPr.Elements().Where(e => e.Name != W.rPr && e.Name != W.sectPr && e.Name != W.pPrChange)));
                }

                numIdEl.SetAttributeValue(W.val, mapped);
                if (oldPPr is not null)
                    pPr!.Add(new XElement(W.pPrChange, state.RevisionAttributes(), oldPPr));
                changed = true;
            }
            if (changed)
                part.PutXDocument();
        }
    }

    /// <summary>
    /// Rebind numbering references owned by styles copied from the RIGHT package.  This deliberately
    /// receives the import-provenance set from <see cref="TrackStyleDefinitionChanges"/> instead of
    /// walking every output style: pre-existing LEFT styles, including archived property histories,
    /// must keep resolving through the LEFT numbering definitions.  Property-change archives within
    /// a copied style are likewise left untouched.
    /// </summary>
    private static void RebindRightImportedStyleNumberingReferences(
        StyleDefinitionsPart? stylesPart, IReadOnlySet<StyleIdentity> rightImportedStyles,
        Dictionary<int, int> numIdMap)
    {
        if (stylesPart is null || rightImportedStyles.Count == 0 || numIdMap.Count == 0)
            return;

        var root = stylesPart.GetXDocument().Root;
        if (root is null)
            return;

        var changed = false;
        foreach (var style in root.Elements(W.style))
        {
            var type = (string?)style.Attribute(W.type);
            var styleId = (string?)style.Attribute(W.styleId);
            if (type is null || styleId is null ||
                !rightImportedStyles.Contains(new StyleIdentity(type, styleId)))
                continue;

            // A style can carry paragraph properties directly or in a table-style conditional
            // payload.  Rebind either current payload, but never an archived *Change payload.
            foreach (var numIdEl in style.Descendants(W.numPr).Elements(W.numId).ToList())
            {
                if (numIdEl.Ancestors().Any(a =>
                        a.Name.LocalName.EndsWith("Change", StringComparison.Ordinal)))
                    continue;
                if (int.TryParse((string?)numIdEl.Attribute(W.val), out var id) &&
                    numIdMap.TryGetValue(id, out var mapped))
                {
                    numIdEl.SetAttributeValue(W.val, mapped);
                    changed = true;
                }
            }
        }

        if (changed)
            stylesPart.PutXDocument();
    }

    /// <summary>A paragraph that exists only in the right document: inserted pilcrow
    /// (<c>pPr/rPr/w:ins</c>), or all of its content inside <c>w:ins</c>/<c>w:moveTo</c> wrappers
    /// with no live or deleted runs.</summary>
    private static bool IsInsertedParagraph(XElement paragraph)
    {
        if (paragraph.Element(W.pPr)?.Element(W.rPr)?.Element(W.ins) is not null)
            return true;
        var hasIns = false;
        foreach (var child in paragraph.Elements())
        {
            if (child.Name == W.ins || child.Name == W.moveTo)
                hasIns = true;
            else if (child.Name == W.r || child.Name == W.hyperlink || child.Name == W.del || child.Name == W.moveFrom)
                return false;
        }
        return hasIns;
    }

    /// <summary>A paragraph whose mark is deleted (including a move-from encoded as a deleted mark) belongs to
    /// the left/reject state and must continue resolving through the left numbering definition.</summary>
    private static bool IsDeletedParagraph(XElement paragraph) =>
        paragraph.Element(W.pPr)?.Element(W.rPr)?.Elements().Any(e =>
            e.Name == W.del || e.Name == W.moveFrom) == true;

    /// <summary>When the output styles part lacks <c>w:docDefaults</c> (or the whole part is
    /// missing), insert Word's stock docDefaults (<see cref="WordStockDocDefaults"/>) — the era
    /// variant keyed on whether the left shipped a theme. See the call site for provenance.</summary>
    private static void BackfillStockDocDefaults(MainDocumentPart main, bool leftHadTheme)
    {
        var stylesPart = main.StyleDefinitionsPart;
        if (stylesPart is null)
        {
            stylesPart = main.AddNewPart<StyleDefinitionsPart>("rIdStylesBackfill");
            var stylesDoc = stylesPart.GetXDocument();
            stylesDoc.Add(new XElement(W.styles,
                new XAttribute(XNamespace.Xmlns + "w", W.w.NamespaceName)));
            stylesPart.PutXDocument();
        }
        var root = stylesPart.GetXDocument().Root;
        if (root is null || root.Element(W.docDefaults) is not null)
            return;
        var stock = XElement.Parse(leftHadTheme ? WordStockDocDefaults.ClassicXml : WordStockDocDefaults.ModernXml);
        root.AddFirst(stock);
        stylesPart.PutXDocument();
    }

    /// <summary>Write Microsoft Word's stock default theme (<see cref="WordStockTheme"/> — Aptos
    /// fonts, 2023+ Office palette, byte-for-byte as Word's compare backfills it) into a fresh
    /// <see cref="ThemePart"/>. See the call site for why the RIGHT's theme must not be adopted
    /// instead.</summary>
    private static void BackfillDefaultTheme(MainDocumentPart main)
    {
        // Explicit relationship id: AddNewPart's auto-generated ids are RANDOM, which breaks
        // byte-determinism between identical Compare invocations.
        var themePart = main.AddNewPart<ThemePart>("rIdThemeBackfill");
        using var writer = new StreamWriter(themePart.GetStream(FileMode.Create), new System.Text.UTF8Encoding(false));
        writer.Write(WordStockTheme.Xml);
    }

    /// <summary>The RAW formatting payload of a style definition: its direct <c>pPr</c>/<c>rPr</c>
    /// minus tracked-change markers and rsid noise.</summary>
    private static XElement RawStylePayload(XName name, XElement style)
    {
        var props = style.Element(name);
        var clone = props is null ? new XElement(name) : new XElement(props);
        clone.Descendants().Where(d => d.Name == W.rsid || d.Name == W.pPrChange || d.Name == W.rPrChange)
            .Remove();
        clone.DescendantsAndSelf().Attributes().Where(at => at.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal))
            .Remove();
        return clone;
    }

    /// <summary>Compare raw style formatting while deliberately ignoring property-history and rsid noise.
    /// A raw match usually needs no style revision; the separate, tightly guarded docDefaults projection
    /// may still materialize a reversible effective-format delta for a used style.</summary>
    private static bool StyleDefinitionPayloadsEqual(XElement a, XElement b)
        => XNode.DeepEquals(RawStylePayload(W.pPr, a), RawStylePayload(W.pPr, b)) &&
           XNode.DeepEquals(RawStylePayload(W.rPr, a), RawStylePayload(W.rPr, b));

    /// <summary>
    /// Resolve a style's EFFECTIVE direct formatting within one styles part: docDefaults underlaid,
    /// then the basedOn chain overlaid outermost-first, then the style's own definition. Same-named
    /// property elements replace; <c>w:rFonts</c> merges attribute-wise (matching how Word materializes
    /// the resolved fonts into a tracked style update). Tracked-change and rsid noise excluded.
    /// </summary>
    private static (XElement PPr, XElement RPr) ResolveEffectiveStyleFormatting(
        XElement stylesRoot, string? type, string styleId)
    {
        var accPPr = new XElement(W.pPr,
            stylesRoot.Element(W.docDefaults)?.Element(W.pPrDefault)?.Element(W.pPr)?.Elements());
        var accRPr = new XElement(W.rPr,
            stylesRoot.Element(W.docDefaults)?.Element(W.rPrDefault)?.Element(W.rPr)?.Elements());

        var chain = new List<XElement>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var currentId = styleId;
        while (currentId is not null && seen.Add(currentId) && chain.Count < 16)
        {
            var style = stylesRoot.Elements(W.style).FirstOrDefault(st =>
                (string?)st.Attribute(W.type) == type &&
                (string?)st.Attribute(W.styleId) == currentId);
            if (style is null)
                break;
            chain.Add(style);
            currentId = (string?)style.Element(W.basedOn)?.Attribute(W.val);
        }
        chain.Reverse();   // outermost ancestor first, the style itself last

        foreach (var style in chain)
        {
            OverlayProps(accPPr, style.Element(W.pPr));
            OverlayProps(accRPr, style.Element(W.rPr));
        }
        StripStyleNoise(accPPr);
        StripStyleNoise(accRPr);
        return (accPPr, accRPr);
    }

    private static void OverlayProps(XElement acc, XElement? layer)
    {
        if (layer is null)
            return;
        foreach (var prop in layer.Elements())
        {
            if (prop.Name == W.pPrChange || prop.Name == W.rPrChange || prop.Name == W.rsid)
                continue;
            var existing = acc.Element(prop.Name);
            if (prop.Name == W.rFonts && existing is not null)
            {
                foreach (var at in prop.Attributes())
                    existing.SetAttributeValue(at.Name, at.Value);
                continue;
            }
            existing?.Remove();
            acc.Add(new XElement(prop));
        }
    }

    private static void StripStyleNoise(XElement props)
    {
        props.Descendants().Where(d => d.Name == W.rsid).Remove();
        props.DescendantsAndSelf().Attributes()
            .Where(at => at.Name.LocalName.StartsWith("rsid", StringComparison.Ordinal)).Remove();
    }

    /// <summary>
    /// Reconcile every <c>w:headerReference</c>/<c>w:footerReference</c> in the output body (see the
    /// call site in <see cref="Render"/> for the phenomenon): a reference whose <c>r:id</c> does not
    /// resolve to a part of its own kind is REBOUND to the output part carrying that story — the part
    /// the story diff produced for the right part with that id (matched → merged left part; inserted →
    /// fresh part), or a wholesale import of the right story part as a last resort — and only
    /// references neither package can resolve are removed. Finally, duplicate same-kind/same-type
    /// references within one sectPr (a rebind can land next to a story-diff-attached reference to the
    /// SAME part) are collapsed to the first.
    /// </summary>
    private static void RebindOrStripStoryReferences(
        RenderState state, MainDocumentPart main, MainDocumentPart? rightMain)
    {
        var xDoc = main.GetXDocument();
        var body = xDoc.Root?.Element(W.body);
        if (body is null)
            return;

        var headerIds = new HashSet<string>(StringComparer.Ordinal);
        var footerIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (var rel in main.Parts)
        {
            if (rel.OpenXmlPart is HeaderPart)
                headerIds.Add(rel.RelationshipId);
            else if (rel.OpenXmlPart is FooterPart)
                footerIds.Add(rel.RelationshipId);
        }

        var remap = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var el in body.Descendants()
                     .Where(e => e.Name == W.headerReference || e.Name == W.footerReference)
                     .ToList())
        {
            var isHeader = el.Name == W.headerReference;
            var valid = isHeader ? headerIds : footerIds;
            if ((string?)el.Attribute(R.id) is not { Length: > 0 } id)
            {
                el.Remove();
                continue;
            }
            if (valid.Contains(id))
                continue;
            if (remap.TryGetValue(id, out var mappedId))
            {
                el.SetAttributeValue(R.id, mappedId);
                continue;
            }

            OpenXmlPart? rightPart = null;
            if (rightMain is not null)
            {
                try { rightPart = rightMain.GetPartById(id); }
                catch (ArgumentOutOfRangeException) { }
            }
            var kindMatches = isHeader ? rightPart is HeaderPart : rightPart is FooterPart;
            if (!kindMatches || rightPart is null)
            {
                el.Remove();   // unresolvable in either package — inheritance takes over
                continue;
            }

            if (!state.StoryOutputParts.TryGetValue(rightPart.Uri, out var target))
            {
                // Last resort: import the right story part wholesale. Descendants carrying their own
                // relationship references (drawings, embedded images) are pruned — their relationship
                // graph did not come along, and a dangling id INSIDE the story part would make the
                // package unloadable again. Text, fields and page numbers survive.
                target = isHeader
                    ? main.AddNewPart<HeaderPart>()
                    : (OpenXmlPart)main.AddNewPart<FooterPart>();
                var clone = new XElement(rightPart.GetXDocument().Root!);
                clone.Descendants()
                    .Where(d => d.Attributes().Any(a => a.Name.Namespace == R.r))
                    .ToList()
                    .ForEach(d => d.Remove());
                foreach (var attr in clone.DescendantsAndSelf().Attributes()
                             .Where(a => a.Name.Namespace == PtOpenXml.pt).ToList())
                    attr.Remove();
                var targetXDoc = target.GetXDocument();
                if (targetXDoc.Root is null)
                    targetXDoc.Add(clone);
                else
                    targetXDoc.Root.ReplaceWith(clone);
                target.PutXDocument();
                state.StoryOutputParts[rightPart.Uri] = target;
            }

            var newId = main.GetIdOfPart(target);
            remap[id] = newId;
            el.SetAttributeValue(R.id, newId);
            (isHeader ? headerIds : footerIds).Add(newId);
        }

        // Collapse duplicate same-kind/same-type references within each sectPr (keep the first).
        foreach (var sectPr in body.Descendants(W.sectPr))
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (var el in sectPr.Elements()
                         .Where(e => e.Name == W.headerReference || e.Name == W.footerReference)
                         .ToList())
            {
                var key = $"{el.Name.LocalName}:{(string?)el.Attribute(W.type)}";
                if (!seen.Add(key))
                    el.Remove();
            }
        }
        main.PutXDocument();
    }

    /// <summary>Every relationship id currently in use on <paramref name="part"/>, all kinds
    /// (part, hyperlink, external, data-part reference).</summary>
    private static HashSet<string> UsedRelationshipIds(OpenXmlPart part)
    {
        var used = new HashSet<string>(StringComparer.Ordinal);
        foreach (var rel in part.Parts) used.Add(rel.RelationshipId);
        foreach (var rel in part.HyperlinkRelationships) used.Add(rel.Id);
        foreach (var rel in part.ExternalRelationships) used.Add(rel.Id);
        foreach (var rel in part.DataPartReferenceRelationships) used.Add(rel.Id);
        return used;
    }

    /// <summary>A relationship id not currently in use by any of the left main part's relationships (parts,
    /// hyperlinks, external links, and data-part references alike). Deterministic: the first free
    /// <c>rIdRemap{n}</c> (n ascending from 1). The dedicated <c>rIdRemap</c> prefix avoids colliding with the
    /// document's own <c>rId{n}</c> numbering — the very collision this remap exists to resolve.</summary>
    private static string FreshRelationshipId(OpenXmlPart leftMain)
    {
        var used = UsedRelationshipIds(leftMain);
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

    // ------------------------------------------- input-revision preservation (PreserveInputRevisions)

    /// <summary>Every element name <see cref="RevisionProcessor"/> recognizes as tracked-revision markup —
    /// the "does this block carry pre-existing revisions worth preserving?" gate.</summary>
    private static readonly HashSet<XName> TrackedRevisionNames = new(RevisionProcessor.TrackedRevisionsElements);

    /// <summary>Range-start → matching range-end names among the tracked revision elements. The two
    /// endpoints of one range deliberately share an id; every other tracked-revision element needs a
    /// distinct annotation id.</summary>
    private static readonly IReadOnlyDictionary<XName, XName> PreservedRangeEnds =
        new Dictionary<XName, XName>
        {
            [W.moveFromRangeStart] = W.moveFromRangeEnd,
            [W.moveToRangeStart] = W.moveToRangeEnd,
            [W.customXmlDelRangeStart] = W.customXmlDelRangeEnd,
            [W.customXmlInsRangeStart] = W.customXmlInsRangeEnd,
        };

    /// <summary>Range-end → matching range-start names, materialized once for normalizing preserved clones.</summary>
    private static readonly IReadOnlyDictionary<XName, XName> PreservedRangeStarts =
        PreservedRangeEnds.ToDictionary(p => p.Value, p => p.Key);

    /// <summary>The run-level revision WRAPPERS a preserved (original) block may legitimately contain as
    /// paragraph children. When <c>PreserveInputRevisions</c> is on, <see cref="MarkWholeParagraph"/> leaves
    /// these as-is instead of re-wrapping them — a foreign <c>w:ins</c> stays a single <c>w:ins</c> (its
    /// content is already marked inserted), a foreign <c>w:del</c>/<c>w:moveFrom</c> stays deleted-grade, and
    /// no same-kind wrapper ever nests. Range markers ride through unchanged (they wrap nothing).</summary>
    private static readonly HashSet<XName> PreservedWrapperNames = new()
    {
        W.ins, W.del, W.moveFrom, W.moveTo,
        W.moveFromRangeStart, W.moveFromRangeEnd, W.moveToRangeStart, W.moveToRangeEnd,
    };

    /// <summary>
    /// Build the accepted-working-element → ORIGINAL right body element(s) map that powers
    /// <c>PreserveInputRevisions</c>. The renderer's IR read normalizes revisions by ACCEPTING the whole
    /// working copy first (see <see cref="IrReaderOptions.RevisionView"/>), so every retained
    /// <c>Source.Element</c> is revision-free; preserving Word-style requires reaching back to the ORIGINAL
    /// elements. Pairing is an in-order two-pointer walk over the two bodies' child sequences that mirrors
    /// the document-level accept's modeled body restructurings — a paragraph whose MARK is deleted merges into
    /// the NEXT paragraph (a fully-deleted paragraph vanishes the same way), while adjacent tables with the
    /// same bidi setting coalesce. A GROUP of original elements grows until its last member no longer
    /// merge-continues, then its accepted visible
    /// text to equal the working block's text. A verified group maps working → [originals] (one working
    /// block may correspond to SEVERAL originals: the mark-deleted members ride along and vanish again on
    /// accept, exactly Word's shape). Any unexplained divergence (removed content controls, trailing
    /// unmatched originals) stops the walk and returns the PARTIAL map — every entry
    /// already emitted was verified in order, and unmapped blocks degrade to accepted-view emission rather
    /// than risking a wrong pairing. Only markup-bearing groups are mapped (a clean block's accepted
    /// emission already equals its original content, so mapping it would change nothing but bytes churn).
    /// <para><b>Note scopes ride the same map.</b> Footnote/endnote definitions pair by <c>w:id</c> (stable
    /// across accept) and each paired definition's child blocks are aligned with the same walk — the note
    /// renderer dispatches through the shared <see cref="RenderBlockOp"/>, so Equal/Insert note blocks then
    /// preserve through the same two emit hooks with zero extra plumbing.</para>
    /// </summary>
    private static Dictionary<XElement, List<XElement>>? BuildPreservedOriginalIndex(IrDocument irRight, WmlDocument right)
    {
        // The working (accepted) part trees: the IR read pins each part's parsed XDocument in Sources;
        // every body block's Source.Element lives in the w:document tree, note blocks in w:footnotes/w:endnotes.
        XElement? WorkingRoot(XName rootName) =>
            irRight.Sources.Values.Select(xd => xd.Root).FirstOrDefault(r => r?.Name == rootName);

        var workingBody = WorkingRoot(W.document)?.Element(W.body);
        if (workingBody == null)
            return null;

        var map = new Dictionary<XElement, List<XElement>>();
        using (var streamDoc = new OpenXmlMemoryStreamDocument(right))
        using (var wDoc = streamDoc.GetWordprocessingDocument())
        {
            var main = wDoc.MainDocumentPart;
            var originalBody = main?.GetXDocument().Root?.Element(W.body);
            if (originalBody == null)
                return null;
            AlignPreservedChildren(workingBody, originalBody, map);

            AlignPreservedNoteScope(WorkingRoot(W.footnotes),
                main?.FootnotesPart?.GetXDocument().Root, W.footnote, map);
            AlignPreservedNoteScope(WorkingRoot(W.endnotes),
                main?.EndnotesPart?.GetXDocument().Root, W.endnote, map);
        }
        return map.Count == 0 ? null : map;
    }

    /// <summary>
    /// Find the deliberately tiny inverse of the dirty-left insertion projection: a raw LEFT whole-block
    /// deletion that Word carries forward as its ORIGINAL insertion when the accepted RIGHT view contains that
    /// block. This is intentionally not an alignment algorithm. A target must be a direct child of the RIGHT
    /// working body's exact ordinal, and the raw LEFT child at that SAME ordinal must be a fully deleted
    /// <c>w:p</c>/<c>w:tbl</c> which has no accepted LEFT block at that ordinal. Matching the block name and
    /// projected text is the final guard. That shape occurs in Word's own reintroduced-deletion redlines; any
    /// shifted, nested, mixed, field-bearing, move/property, or partially bare shape remains on the ordinary
    /// comparer-authored insertion fallback.
    /// </summary>
    private static Dictionary<XElement, XElement>? BuildLeftDeletedInsertionIndex(
        IrDocument irLeft,
        IrDocument irRight,
        WmlDocument left)
    {
        XElement? WorkingBody(IrDocument ir) =>
            ir.Sources.Values.Select(xd => xd.Root).FirstOrDefault(r => r?.Name == W.document)?.Element(W.body);

        var acceptedLeftBody = WorkingBody(irLeft);
        var workingRightBody = WorkingBody(irRight);
        if (acceptedLeftBody == null || workingRightBody == null)
            return null;

        using var streamDoc = new OpenXmlMemoryStreamDocument(left);
        using var wDoc = streamDoc.GetWordprocessingDocument();
        var rawLeftBody = wDoc.MainDocumentPart?.GetXDocument().Root?.Element(W.body);
        if (rawLeftBody == null)
            return null;

        // Ordinals intentionally count EVERY direct body element, including comment/bookmark/range leaves.
        // The raw and working bodies may have different lengths after accepting revisions; that is precisely
        // why this is a per-ordinal eligibility check rather than a two-pointer resynchronization walk.
        var rawLeft = rawLeftBody.Elements().ToList();
        var acceptedLeft = acceptedLeftBody.Elements().ToList();
        var workingRight = workingRightBody.Elements().ToList();
        var map = new Dictionary<XElement, XElement>();

        for (int ordinal = 0; ordinal < workingRight.Count && ordinal < rawLeft.Count; ordinal++)
        {
            var target = workingRight[ordinal];
            var candidate = rawLeft[ordinal];
            if ((target.Name != W.p && target.Name != W.tbl) || candidate.Name != target.Name)
                continue;

            // The raw deleted block must not still be represented by an accepted LEFT p/tbl at this position.
            // We deliberately do not search elsewhere: a shifted match could reattribute an unrelated change.
            if (ordinal < acceptedLeft.Count &&
                (acceptedLeft[ordinal].Name == W.p || acceptedLeft[ordinal].Name == W.tbl))
                continue;

            if (!IsProjectableLeftDeletionAsInsertion(candidate))
                continue;
            if (!string.Equals(DeletedProjectionVisibleText(candidate), VisibleText(target), StringComparison.Ordinal))
                continue;

            map[target] = candidate;
        }

        return map.Count == 0 ? null : map;
    }

    /// <summary>Whether one raw LEFT direct body block is safe to turn from native deletion to native insertion.
    /// The source has to be all-and-only deletion markup: paragraph candidates need a deleted paragraph mark;
    /// table candidates need every direct row marked deleted. Fields, moves, property revisions, a foreign
    /// insertion, malformed <c>w:t</c> in deleted content, and every bare paragraph child are rejected.
    /// </summary>
    private static bool IsProjectableLeftDeletionAsInsertion(XElement candidate)
    {
        if (candidate.Name == W.p)
        {
            if (candidate.Element(W.pPr)?.Element(W.rPr)?.Element(W.del) == null)
                return false;
        }
        else if (candidate.Name == W.tbl)
        {
            var rows = candidate.Elements(W.tr).ToList();
            if (rows.Count == 0 || rows.Any(row => row.Element(W.trPr)?.Element(W.del) == null))
                return false;
        }
        else
        {
            return false;
        }

        // A simple/complex field has distinct containment and text-conversion rules. This projection does not
        // expand or rebuild it, so it is safer to leave it to the proven accepted-view whole-block renderer.
        if (candidate.DescendantsAndSelf().Any(e =>
                e.Name == W.fldSimple || e.Name == W.fldChar ||
                e.Name == W.instrText || e.Name == W.delInstrText))
            return false;

        // RevisionProcessor's public tracked-element list does not include the custom-XML move range
        // endpoints even though it transforms them. They are move provenance, never simple deletion
        // provenance, so keep all four on the accepted-view fallback explicitly.
        if (candidate.DescendantsAndSelf().Any(e =>
                e.Name == W.customXmlMoveFromRangeStart || e.Name == W.customXmlMoveFromRangeEnd ||
                e.Name == W.customXmlMoveToRangeStart || e.Name == W.customXmlMoveToRangeEnd))
            return false;

        var tracked = candidate.DescendantsAndSelf()
            .Where(e => TrackedRevisionNames.Contains(e.Name))
            .ToList();
        if (!tracked.Any(e => e.Name == W.del) ||
            tracked.Any(e => e.Name != W.del && e.Name != W.delText))
            return false;

        // `w:delText` is valid only under a deletion wrapper. A raw `w:t` would need a different semantic
        // conversion and proves the candidate is not wholly deleted.
        if (candidate.DescendantsAndSelf(W.delText).Any(t => !t.Ancestors(W.del).Any()) ||
            candidate.DescendantsAndSelf(W.t).Any())
            return false;

        // At the paragraph level no run-level child may be bare: preserving an arbitrary marker/run alongside
        // the converted wrappers would make part of an ostensibly deleted block survive the projection.
        return candidate.DescendantsAndSelf(W.p)
            .All(p => p.Elements().All(child => child.Name == W.pPr || child.Name == W.del));
    }

    /// <summary>The visible text a raw deleted block will expose after its <c>w:delText</c> nodes become
    /// <c>w:t</c>. Called only after <see cref="IsProjectableLeftDeletionAsInsertion"/> verified the strict
    /// all-deleted shape.</summary>
    private static string DeletedProjectionVisibleText(XElement candidate) =>
        string.Concat(candidate.Descendants(W.delText).Select(t => t.Value));

    /// <summary>Pair each working note definition with the original of the SAME <c>w:id</c> (note ids are
    /// untouched by the accept normalization) and align their child blocks — the note-scope leg of
    /// <see cref="BuildPreservedOriginalIndex"/>. Null roots (scope absent on either side) are a no-op.</summary>
    private static void AlignPreservedNoteScope(
        XElement? workingRoot, XElement? originalRoot, XName noteName, Dictionary<XElement, List<XElement>> map)
    {
        if (workingRoot == null || originalRoot == null)
            return;
        var originalById = new Dictionary<string, XElement>(StringComparer.Ordinal);
        foreach (var note in originalRoot.Elements(noteName))
            if ((string?)note.Attribute(W.id) is { } id && !originalById.ContainsKey(id))
                originalById[id] = note;
        foreach (var workingNote in workingRoot.Elements(noteName))
            if ((string?)workingNote.Attribute(W.id) is { } id && originalById.TryGetValue(id, out var originalNote))
                AlignPreservedChildren(workingNote, originalNote, map);
    }

    /// <summary>The container-level two-pointer alignment walk of <see cref="BuildPreservedOriginalIndex"/>
    /// (see there for the model and the conservative-bail rules). Adds verified markup-bearing groups to
    /// <paramref name="map"/>; a divergence stops THIS container's walk only (entries already added stand).</summary>
    private static void AlignPreservedChildren(
        XElement workingContainer, XElement originalContainer, Dictionary<XElement, List<XElement>> map)
    {
        var working = workingContainer.Elements().ToList();
        var original = originalContainer.Elements().ToList();

        int i = 0;
        foreach (var w in working)
        {
            // AcceptDeletedAndMoveFromParagraphMarks rebuilds body/note containers from p/tbl blocks and
            // retains only body sectPr. Direct leaf annotations (comment/bookmark/range markers, etc.) are
            // therefore discarded. Skip only such leaves, and only when the working side does not carry the
            // same name, so a non-normalized 1:1 marker still has to match exactly.
            while (i < original.Count && IsAcceptDiscardedLeaf(original[i], w))
                i++;
            if (i >= original.Count)
                return;   // originals exhausted early — alignment lost; keep what was verified.

            // AcceptRevisions merges a maximal direct run of clean, same-bidi tables into one table. This is
            // resynchronization only: the accepted working table is already the correct renderer source, and
            // mapping several raw tables back to it could later re-emit a structurally different table run.
            if (TryConsumeCleanTableMerge(w, original, ref i))
                continue;

            // Any remaining non-block child (not an accept-discarded leaf) requires an exact 1:1 name match
            // to stay aligned; it is never preserved itself.
            if (w.Name != W.p && w.Name != W.tbl)
            {
                if (original[i].Name != w.Name)
                    return;
                i++;
                continue;
            }

            string target = VisibleText(w);
            var group = new List<XElement>();
            var acc = new System.Text.StringBuilder();
            while (true)
            {
                // Cover a direct leaf that sat between a mark-deleted paragraph and the paragraph it merges
                // into. The accept transform drops it just like a leaf before an ordinary working block.
                while (i < original.Count && IsAcceptDiscardedLeaf(original[i], w))
                    i++;
                if (i >= original.Count)
                    return;   // ran out of originals mid-group — alignment lost.
                var o = original[i];
                group.Add(o);
                acc.Append(AcceptedVisibleText(o));
                i++;
                // A paragraph whose MARK is deleted merge-continues into the next original — keep growing.
                if (HasDeletedParagraphMark(o))
                    continue;
                // Group boundary: the last member contributes the block identity. Verify.
                if (o.Name != w.Name || !string.Equals(acc.ToString(), target, StringComparison.Ordinal))
                    return;   // accept semantics we do not model (removed sdt, etc.) — stop.
                break;
            }

            if (group.Any(g => g.Descendants().Any(d => TrackedRevisionNames.Contains(d.Name))))
                map[w] = group;
        }
    }

    /// <summary>True when the original child is a direct leaf discarded by the accept transform that rebuilds
    /// body/note block containers. A wrapper with any nested paragraph/table is deliberately NOT skipped: its
    /// topology needs explicit modeling. The name guard preserves exact matching when the working container
    /// still carries the same leaf.</summary>
    private static bool IsAcceptDiscardedLeaf(XElement original, XElement working) =>
        original.Name != working.Name &&
        original.Name != W.p && original.Name != W.tbl && original.Name != W.sectPr &&
        !original.Descendants().Any(e => e.Name == W.p || e.Name == W.tbl);

    /// <summary>Consume exactly the clean table run that
    /// <see cref="RevisionProcessor.AcceptRevisions"/> coalesces into <paramref name="working"/>. The run is
    /// never entered into the preservation map: cloning several raw tables for one accepted source would change
    /// table structure. Failure leaves <paramref name="originalIndex"/> unchanged, so the ordinary strict
    /// 1:1 walk either verifies a non-merged table or bails conservatively.</summary>
    private static bool TryConsumeCleanTableMerge(
        XElement working, IReadOnlyList<XElement> original, ref int originalIndex)
    {
        if (working.Name != W.tbl || originalIndex >= original.Count || original[originalIndex].Name != W.tbl)
            return false;

        bool BidiVisual(XElement table) => table.Element(W.tblPr)?.Element(W.bidiVisual) != null;
        bool bidiVisual = BidiVisual(original[originalIndex]);
        int end = originalIndex + 1;
        while (end < original.Count && original[end].Name == W.tbl && BidiVisual(original[end]) == bidiVisual)
            end++;
        if (end - originalIndex < 2)
            return false;

        var run = original.Skip(originalIndex).Take(end - originalIndex).ToList();
        if (run.Any(table => table.DescendantsAndSelf().Any(e => TrackedRevisionNames.Contains(e.Name))))
            return false;
        if (!string.Equals(VisibleText(working), string.Concat(run.Select(AcceptedVisibleText)), StringComparison.Ordinal))
            return false;
        if (working.Elements(W.tr).Count() != run.Sum(table => table.Elements(W.tr).Count()))
            return false;

        originalIndex = end;
        return true;
    }

    /// <summary>
    /// Normalize a PRESERVED original clone before emission: every tracked-revision element gets a FRESH
    /// <c>w:id</c> from the render's single ascending counter (the input's own ids would collide with this
    /// diff's — validator-flagged duplicates). The two endpoints of an actual range keep one shared fresh
    /// id, while a wrapper/property-change that happens to reuse its input id gets its own id. This repairs
    /// malformed input documents that reuse one id for unrelated revisions. Non-W-namespace extension
    /// attributes on revision elements (e.g. Word 2023's
    /// <c>w16du:dateUtc</c>) are dropped — the preserved facts are author/date/content, and the extension
    /// attrs are undeclared noise to the SDK schema the output is validated against.
    /// </summary>
    private static XElement NormalizePreservedClone(XElement clone, RenderState state)
    {
        foreach (var rev in clone.DescendantsAndSelf().Where(e => TrackedRevisionNames.Contains(e.Name)))
        {
            if (rev.Attribute(W.id) is { } id)
            {
                id.Value = PreservedRangeEnds.ContainsKey(rev.Name)
                    ? state.OpenPreservedRange(rev.Name, id.Value)
                    : PreservedRangeStarts.TryGetValue(rev.Name, out var startName)
                        ? state.ClosePreservedRange(startName, id.Value)
                        : state.NextId().ToString(System.Globalization.CultureInfo.InvariantCulture);
            }
            rev.Attributes()
                .Where(a => !a.IsNamespaceDeclaration && a.Name.Namespace != W.w)
                .Remove();
        }
        return clone;
    }

    /// <summary>Project a raw LEFT input insertion onto the comparison's DELETE side. Word does not nest the
    /// old <c>w:ins</c> inside a new deletion; it converts that insertion — including a paragraph/row mark —
    /// to deletion-grade markup in place. This helper is called only for the conservative eligibility slice in
    /// <see cref="RenderState.ProjectableLeftDeletionGroup"/>: groups whose tracked state is exclusively
    /// <c>w:ins</c>. That keeps moves/property revisions on the accepted-view fallback until their distinct
    /// source-history semantics are modeled.</summary>
    private static void ProjectLeftInsertionsAsDeletions(XElement clone)
    {
        foreach (var ins in clone.DescendantsAndSelf(W.ins).ToList())
        {
            ins.Name = W.del;
            ConvertTextToDelText(ins);
        }
    }

    /// <summary>Project a raw LEFT whole-block deletion onto the comparison's INSERT side. This is the reverse
    /// of <see cref="ProjectLeftInsertionsAsDeletions"/> and is called only after the direct-body ordinal and
    /// strict pure-deletion checks in <see cref="BuildLeftDeletedInsertionIndex"/>. The original author/date
    /// attributes stay on the native revision wrapper; only its sense and text element names change.</summary>
    private static void ProjectLeftDeletionsAsInsertions(XElement clone)
    {
        foreach (var del in clone.DescendantsAndSelf(W.del).ToList())
            del.Name = W.ins;
        foreach (var delText in clone.DescendantsAndSelf(W.delText).ToList())
            delText.Name = W.t;
    }

    /// <summary>True when a paragraph's MARK is revision-deleted (<c>w:pPr/w:rPr/w:del</c> or
    /// <c>w:moveFrom</c>) — on accept the paragraph merges into the NEXT one (vanishing entirely when its
    /// content is all delete-grade), the one body restructuring the preservation walk models.</summary>
    private static bool HasDeletedParagraphMark(XElement block) =>
        block.Name == W.p &&
        (block.Element(W.pPr)?.Element(W.rPr)?.Elements()
            .Any(e => e.Name == W.del || e.Name == W.moveFrom) ?? false);

    /// <summary>Concatenated <c>w:t</c> text of an (already accepted) working block.</summary>
    private static string VisibleText(XElement block) =>
        string.Concat(block.Descendants(W.t).Select(t => t.Value));

    /// <summary>The accepted-view visible text of an ORIGINAL (markup-bearing) block: every <c>w:t</c> not
    /// inside delete-grade markup (<c>w:del</c> holds <c>w:delText</c>, excluded implicitly; <c>w:moveFrom</c>
    /// holds <c>w:t</c> that accept REMOVES, excluded explicitly).</summary>
    private static string AcceptedVisibleText(XElement block) =>
        string.Concat(block.Descendants(W.t)
            .Where(t => !t.Ancestors(W.moveFrom).Any() && !t.Ancestors(W.del).Any())
            .Select(t => t.Value));

    private static IReadOnlyList<IrDiffToken> ParagraphTokens(string? anchor, IrDocument doc, IrDiffSettings settings)
    {
        if (anchor != null && doc.AnchorIndex.TryGetValue(anchor, out var block) && block is IrParagraph p)
            return IrDiffTokenizer.Tokenize(p, settings);
        return Array.Empty<IrDiffToken>();
    }

    /// <summary>Strip the reader-assigned <c>pt:Unid</c> bookkeeping attributes/elements from a cloned element so
    /// the output carries no engine-internal markup.</summary>
    internal static XElement StripUnids(XElement el)
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

        /// <summary>LEFT source anchor by move-group id for the operation list currently being rendered.
        /// Move destinations intentionally carry only their RIGHT anchor; nested render scopes replace this
        /// map temporarily because cell, note, and header projections each use local move-group ids.</summary>
        public IReadOnlyDictionary<int, string> ActiveMoveSourceAnchors { get; set; } =
            new Dictionary<int, string>();

        /// <summary>Resolve the LEFT source anchor for the active move-group scope, if present.</summary>
        public string? MoveSourceAnchor(int moveGroupId) =>
            ActiveMoveSourceAnchors.TryGetValue(moveGroupId, out var anchor) ? anchor : null;

        /// <summary>The document the CURRENTLY-emitting op draws inserted/modified ("right-side") block elements
        /// and token text from. In a two-way render this is always <see cref="Right"/> (set once in the ctor and
        /// never reassigned), so behavior is byte-identical to before this field existed. The composite renderer
        /// switches it per op to the contributing reviewer's IR (or <see cref="Left"/>/base for a base-sourced
        /// equal/delete) so the existing emit helpers can be reused per-reviewer.</summary>
        public IrDocument RightSource { get; set; }

        /// <summary>When non-null, overrides Settings.AuthorForRevisions for emitted revision attributes
        /// (composite multi-author rendering). Null for normal two-way render → behavior unchanged.</summary>
        public string? AuthorOverride { get; set; }

        /// <summary>Style ids defined in the LEFT document's styles part. A PAIRED paragraph's
        /// right-side <c>w:pStyle</c> referencing a style outside this set is dropped when stamped
        /// as current (<see cref="DropUnresolvableStyleRef"/>) — Word expresses a paired
        /// paragraph's format change within the left style universe. Null disables the check.</summary>
        public HashSet<string>? LeftStyleIds { get; set; }

        /// <summary>Accepted-working-element → ORIGINAL right body element(s) map for
        /// <c>PreserveInputRevisions</c> (see <see cref="IrMarkupRenderer.BuildPreservedOriginalIndex"/>).
        /// A working block maps to MULTIPLE originals when the document-level accept merged mark-deleted
        /// paragraphs into it (they ride through and vanish again on accept). Null when preservation is
        /// off, when the original body could not be paired, or in a composite render (Consolidate does
        /// not preserve input revisions in v1) — a null map means zero behavior change.</summary>
        public Dictionary<XElement, List<XElement>>? PreservedOriginals { get; set; }

        /// <summary>Accepted-working-element → ORIGINAL LEFT body element(s) for the narrowly supported
        /// dirty-left delete projection. Unlike <see cref="PreservedOriginals"/>, this map is never used to
        /// carry a left block verbatim: its raw <c>w:ins</c> wrappers are converted to delete-grade markup when
        /// the comparison deletes that block.</summary>
        public Dictionary<XElement, List<XElement>>? LeftPreservedOriginals { get; set; }

        /// <summary>Accepted RIGHT main-body block → raw LEFT whole-block deletion for the narrow inverse
        /// projection. Entries are populated only for direct-body, same-ordinal, fully deleted p/tbl matches;
        /// unlike <see cref="PreservedOriginals"/> this map is consulted only by a literal
        /// <see cref="IrEditOpKind.InsertBlock"/> emission.</summary>
        public Dictionary<XElement, XElement>? LeftDeletedInsertionOriginals { get; set; }

        /// <summary>The ORIGINAL right element group mapped for <paramref name="src"/> under
        /// <c>PreserveInputRevisions</c>, or null (the common case: flag off, a left/composite-sourced
        /// element, or a block with no pre-existing markup).</summary>
        public List<XElement>? PreservedGroup(XElement src) =>
            PreservedOriginals != null && PreservedOriginals.TryGetValue(src, out var group) ? group : null;

        /// <summary>Return a raw LEFT group only when it is safe to project it as a whole-block deletion. The
        /// group may contain ordinary structure (comments/bookmarks/etc.), but every tracked-revision element
        /// must be a <c>w:ins</c> and it must contain no simple fields (which cannot safely sit in a converted
        /// <c>w:del</c>). A source move/property-change needs richer provenance than this whole-block slice can
        /// provide and falls back to the accepted-view renderer.</summary>
        internal List<XElement>? ProjectableLeftDeletionGroup(XElement src)
        {
            if (LeftPreservedOriginals == null || !LeftPreservedOriginals.TryGetValue(src, out var group))
                return null;
            return group.All(member => member.DescendantsAndSelf()
                    .Where(e => TrackedRevisionNames.Contains(e.Name))
                    .All(e => e.Name == W.ins)) &&
                !group.Any(member => member.DescendantsAndSelf(W.fldSimple).Any())
                ? group
                : null;
        }

        /// <summary>The verified raw LEFT deletion to emit as its native insertion at this accepted RIGHT
        /// source block, or null when this is not one of the strict direct-body matches.</summary>
        internal XElement? ProjectableLeftInsertionOriginal(XElement src) =>
            LeftDeletedInsertionOriginals != null && LeftDeletedInsertionOriginals.TryGetValue(src, out var original)
                ? original
                : null;

        // Active preserved range ids, shared across emitted clones because one source range can span
        // multiple preserved blocks. The start element name is part of the key: the same malformed input
        // id may legitimately occur in both a move-from and a move-to range. A stack repairs duplicate/nested
        // same-id ranges deterministically while preserving well-formed start/end pairs.
        private readonly Dictionary<(XName StartName, string OriginalId), Stack<string>> _preservedRangeIds = new();

        /// <summary>Allocate and remember a fresh id for one preserved range start.</summary>
        public string OpenPreservedRange(XName startName, string originalId)
        {
            string fresh = NextId().ToString(System.Globalization.CultureInfo.InvariantCulture);
            var key = (startName, originalId);
            if (!_preservedRangeIds.TryGetValue(key, out var ids))
                _preservedRangeIds[key] = ids = new Stack<string>();
            ids.Push(fresh);
            return fresh;
        }

        /// <summary>Close a preserved range, or give an unmatched malformed end marker its own fresh id.</summary>
        public string ClosePreservedRange(XName startName, string originalId)
        {
            var key = (startName, originalId);
            if (_preservedRangeIds.TryGetValue(key, out var ids) && ids.Count > 0)
            {
                string fresh = ids.Pop();
                if (ids.Count == 0)
                    _preservedRangeIds.Remove(key);
                return fresh;
            }
            return NextId().ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

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

        /// <summary>RIGHT story-part URI → the OUTPUT part carrying that story: the merged left part
        /// for a matched pair, the freshly-created part for an inserted story, or a wholesale import.
        /// Populated by the header/footer scope renderer and the story-reference rebind pass — see
        /// <see cref="RebindOrStripStoryReferences"/>.</summary>
        public Dictionary<Uri, OpenXmlPart> StoryOutputParts { get; } = new();

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

        /// <summary>RIGHT-sourced clone roots that carry a footnote/endnote REFERENCE, bucketed by
        /// <see cref="RightSourceId"/> — the composite renderer's note-id rewrite pass walks these to remap
        /// each reviewer-sourced reference from that reviewer's id space to the base-anchored output space.
        /// Unused (empty) in a two-way render (the rewrite pass only runs in the composite path).</summary>
        public Dictionary<int, List<XElement>> NoteRefClonesBySource { get; } = new();

        /// <summary>Record a RIGHT-sourced clone for media import iff it references any relationship id (an
        /// image embed/link), into the bucket for the currently-active <see cref="RightSourceId"/>. The recorded
        /// element is the live tree node; importing happens post-assembly. Two-way always records into bucket 0.
        /// Also records the clone into <see cref="NoteRefClonesBySource"/> when it carries a footnote/endnote
        /// reference — the composite note-id rewrite's per-reviewer attribution rides the same choke point
        /// every right-sourced clone already passes through.</summary>
        public void RegisterMediaReferences(XElement clone)
        {
            if (clone.DescendantsAndSelf().Attributes().Any(a => a.Name.Namespace == R.r))
            {
                if (!RightSourcedClonesBySource.TryGetValue(RightSourceId, out var list))
                    RightSourcedClonesBySource[RightSourceId] = list = new List<XElement>();
                list.Add(clone);
            }
            if (clone.DescendantsAndSelf()
                    .Any(e => e.Name == W.footnoteReference || e.Name == W.endnoteReference))
            {
                if (!NoteRefClonesBySource.TryGetValue(RightSourceId, out var refs))
                    NoteRefClonesBySource[RightSourceId] = refs = new List<XElement>();
                refs.Add(clone);
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

        /// <summary>Resolved target (external URI, or <c>"#" + anchor</c> for an internal link; null when a
        /// dangling/unresolvable <c>r:id</c>) of each top-level source <c>w:hyperlink</c>, keyed by reference
        /// identity. Stamped onto emitted fragments as <see cref="SourceLinkTarget"/> so the coalescer can merge a
        /// fully-replaced single link (same target both sides) while keeping a genuine retarget (WC019) split.</summary>
        private readonly Dictionary<XElement, string?> _hyperlinkTarget = new(ReferenceEqualityComparer.Instance);

        /// <summary>
        /// Whether this paragraph has no source structure which is transparent to diff tokens. The whitespace
        /// re-anchoring projection may move a token boundary, so it is allowed only for direct <c>w:r</c> children
        /// containing ordinary text (and optional run properties). This excludes fields, hyperlinks, SDTs, smart
        /// tags, revision wrappers, bookmarks, tabs/breaks/drawings, and pre-existing run-format revisions.
        /// </summary>
        public bool SupportsWhitespaceReanchoring { get; }

        public SourceRunModel(XElement para)
        {
            SupportsWhitespaceReanchoring = para.Elements()
                .Where(e => e.Name != W.pPr)
                .All(IsPlainDirectTextRun);

            int charOffset = 0;
            foreach (var child in para.Elements().Where(e => e.Name != W.pPr))
                WalkRunLevel(child, ref charOffset, ContainerChain.Empty);
        }

        private static bool IsPlainDirectTextRun(XElement element) =>
            element.Name == W.r &&
            element.Elements().All(child =>
                child.Name == W.t ||
                (child.Name == W.rPr && !child.Descendants(W.rPrChange).Any()));

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
                _hyperlinkTarget[runLevel] = ResolveLinkTarget(runLevel);
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
                     runLevel.Name == W.sdt || runLevel.Name == W.smartTag || runLevel.Name == W.fldSimple)
            {
                // Non-hyperlink container (sdt/smartTag/accepted ins-del wrapper/direct simple field): one ATOMIC
                // segment spanning its full inner text, emitted whole. A changed fldSimple is always marked as a
                // whole-paragraph structural replacement by FieldEnvelopeDigest; this branch only needs to carry
                // an unchanged simple field once while adjacent token edits use the correct text offset. Its char
                // span mirrors the reader/tokenizer's transparent recursion, including the special hyphens and
                // valid w:sym glyphs that each contribute one visible character. Counting only descendant w:t
                // nodes makes a suffix edit start one character early after a direct simple field containing one
                // of those run children, so the source slicer can drop or misplace the field/result.
                int start = charOffset;
                charOffset += VisibleTextLength(runLevel);
                _segments.Add(new Segment(runLevel, start, charOffset, SegmentKind.Container)
                {
                    Chain = chain,
                    // A result-less direct field has no diff token to claim it at an adjacent edit boundary.
                    // Keep it exactly once just like a structural marker; otherwise an unchanged REF/PAGE field
                    // can disappear merely because following prose changed.
                    AlwaysKeep = runLevel.Name == W.fldSimple && start == charOffset,
                });
            }
            else
            {
                // A non-run, non-container run-level element (bookmarkStart/End, proofErr, commentRangeStart…):
                // zero-width, atomic, kept whole. Bookmark markers are flagged AlwaysKeep so a boundary one is
                // never dropped (they are not diff tokens, so the token-driven boundary flags cannot see them).
                _segments.Add(new Segment(runLevel, charOffset, charOffset, SegmentKind.ZeroWidth)
                    { Chain = chain, AlwaysKeep = IsAlwaysKeepMarker(runLevel.Name) });
            }
        }

        /// <summary>A run-level structural marker that is zero-width, NOT a diff token, and must survive every
        /// edit boundary — a bookmark range endpoint or a COMMENT range endpoint. Dropping one would orphan the
        /// marker and dangle its cross-reference (a bookmark's REF field, a comment's <c>w:comment</c>
        /// definition). Carried through the fine token-diff path exactly like a bookmark, then reconciled to
        /// unique/paired/resolved by <see cref="NormalizeComments"/> (the comment analogue of
        /// <see cref="NormalizeBookmarks"/>). (Field plumbing is handled by <see cref="FieldPlumbingKeep"/>; the
        /// <c>commentReference</c> RUN — zero text — is flagged AlwaysKeep where it is walked in
        /// <see cref="WalkRun"/>.)</summary>
        private static bool IsAlwaysKeepMarker(XName name) =>
            name == W.bookmarkStart || name == W.bookmarkEnd ||
            name == W.commentRangeStart || name == W.commentRangeEnd;

        /// <summary>Field plumbing (<c>w:fldChar</c>/<c>w:instrText</c>/<c>w:delInstrText</c>) — zero-width, not a
        /// diff token, so it must never be dropped at an edit boundary (that orphans the field and dangles its
        /// cross-reference). <c>NormalizeFields</c> re-homes any plumbing that lands in the wrong revision
        /// context back to the field's own context.</summary>
        private static bool FieldPlumbingKeep(XName name) =>
            name == W.fldChar || name == W.instrText || name == W.delInstrText;

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
                else if (child.Name == W.noBreakHyphen || child.Name == W.softHyphen || child.Name == W.sym)
                {
                    // N7/N8: the IR reads these as a SINGLE text character (U+2011 / U+00AD / the symbol glyph), so
                    // the slicer MUST advance the char counter by 1 to stay aligned with the tokenizer — else a
                    // boundary slice over this run is off by one and drops an adjacent character (the reject of a
                    // "Company‑Controlled Intellectual" run dropped the "I"). Emitted whole as its source element.
                    int start = charOffset;
                    charOffset += 1;
                    _segments.Add(new Segment(run, start, charOffset, SegmentKind.RunOther) { OtherChild = child, Chain = chain });
                }
                else
                {
                    // tab/break/drawing/noteref/field-plumbing/commentReference — zero-width run child. Field
                    // plumbing (w:fldChar/w:instrText/w:delInstrText) AND a w:commentReference are AlwaysKeep:
                    // like a bookmark marker neither is a diff token, so one clustered at an edit boundary would
                    // otherwise be dropped — orphaning the field/comment and dangling its cross-reference (editing
                    // text BEFORE a REF field dropped the whole field; editing a commented word dropped the
                    // reference). NormalizeFields re-homes field plumbing; NormalizeComments reconciles the
                    // comment reference. The visible field RESULT is ordinary tokenized text and is unaffected.
                    _segments.Add(new Segment(run, charOffset, charOffset, SegmentKind.RunOther)
                        { OtherChild = child, Chain = chain,
                          AlwaysKeep = FieldPlumbingKeep(child.Name) || child.Name == W.commentReference });
                }
            }
            if (!any)
                _segments.Add(new Segment(run, charOffset, charOffset, SegmentKind.RunOther) { Chain = chain });
        }

        private static bool IsContainer(XName n) =>
            n == W.hyperlink || n == W.ins || n == W.del || n == W.sdt || n == W.smartTag;

        /// <summary>
        /// Visible character length of an atomic source container. This intentionally mirrors the reader's
        /// <see cref="IrReader.InlineWalker"/> / <see cref="IrReader.EmitRunChild"/> topology instead of simply
        /// summing descendant <c>w:t</c> nodes: textbox bodies and opaque containers can contain text in their
        /// raw XML yet are zero-width in the tokenizer's coordinate space. Literal text contributes its UTF-16
        /// length, the two special hyphen elements each contribute one character, and only a valid BMP
        /// <c>w:sym/@w:char</c> contributes one. An invalid symbol is modeled as zero-width opaque content and
        /// must not move a source-slice boundary.
        /// </summary>
        private static int VisibleTextLength(XElement container) => VisibleRunLevelTextLength(container, 0);

        private static int VisibleRunLevelTextLength(XElement element, int sdtDepth)
        {
            if (element.Name == W.r)
            {
                int length = 0;
                foreach (var child in element.Elements())
                    if (child.Name != W.rPr)
                        length += VisibleRunChildTextLength(child);
                return length;
            }

            if (element.Name == W.hyperlink || element.Name == W.ins || element.Name == W.del)
                return element.Elements().Sum(child => VisibleRunLevelTextLength(child, sdtDepth));

            if (element.Name == W.fldSimple)
                return element.Elements().Where(child => child.Name != W.fldData)
                    .Sum(child => VisibleRunLevelTextLength(child, sdtDepth));

            // Keep this depth cap aligned with IrReader.MaxSdtDepth. At the cap the reader preserves an inline
            // content-control envelope opaquely, so none of its descendant raw text has a tokenizer coordinate.
            if (element.Name == W.sdt)
            {
                if (sdtDepth >= 64)
                    return 0;
                var content = element.Element(W.sdtContent);
                return content is null
                    ? 0
                    : content.Elements().Sum(child => VisibleRunLevelTextLength(child, sdtDepth + 1));
            }

            if (element.Name == W.smartTag)
            {
                if (sdtDepth >= 64)
                    return 0;
                return element.Elements().Where(child => child.Name != W.smartTagPr)
                    .Sum(child => VisibleRunLevelTextLength(child, sdtDepth + 1));
            }

            // Non-run-level elements are either reader-dropped or modeled as one zero-width atomic inline:
            // drawing/pict/textbox, tabs/breaks/note refs, field plumbing, and arbitrary opaque XML all land here.
            return 0;
        }

        private static int VisibleRunChildTextLength(XElement child)
        {
            if (child.Name == W.t)
                return child.Value.Length;
            if (child.Name == W.noBreakHyphen || child.Name == W.softHyphen)
                return 1;
            return child.Name == W.sym && IsVisibleSym(child) ? 1 : 0;
        }

        private static bool IsVisibleSym(XElement sym)
        {
            var raw = (string?)sym.Attribute(W.w + "char");
            return raw is not null
                && int.TryParse(raw, System.Globalization.NumberStyles.HexNumber,
                    System.Globalization.CultureInfo.InvariantCulture, out var code)
                && code is >= 0x20 and <= 0xFFFF;
        }

        /// <summary>Resolve a source <c>w:hyperlink</c>'s target, mirroring <c>IrReader.BuildHyperlink</c>: an
        /// <c>@r:id</c> resolves against the owning part's hyperlink relationships to the external URI (the part is
        /// stashed as an annotation on the source tree root by <c>IrReader.Read</c>; a dangling/unresolvable id
        /// yields null); otherwise an <c>@w:anchor</c> internal link uses the convention <c>"#" + anchor</c>.
        /// Returns null when neither is present or an <c>r:id</c> cannot be resolved — the coalescer then declines
        /// to merge on the target basis (conservative), so a fully-replaced link only rejoins when its target is
        /// known and identical on both sides.</summary>
        private static string? ResolveLinkTarget(XElement hyperlink)
        {
            var relId = (string?)hyperlink.Attribute(R.id);
            if (relId != null)
            {
                var part = hyperlink.AncestorsAndSelf().Last().Annotation<OpenXmlPart>();
                if (part != null)
                    foreach (var rel in part.HyperlinkRelationships)
                        if (rel.Id == relId)
                            return rel.Uri?.ToString();
                return null;   // dangling/unresolvable r:id → unknown target.
            }
            var anchor = (string?)hyperlink.Attribute(W.anchor);
            return anchor != null ? "#" + anchor : null;
        }

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
                        {
                            wrapper.SetAttributeValue(SourceLinkId, ord);
                            // Also stamp the resolved target so the coalescer can merge a fully-replaced link
                            // (same target both sides) yet keep a genuine retarget split. A null target sets
                            // nothing (SetAttributeValue(_, null) is a no-op), leaving the run un-mergeable on the
                            // target basis — the conservative default.
                            if (_hyperlinkTarget.TryGetValue(shell, out var tgt) && tgt != null)
                                wrapper.SetAttributeValue(SourceLinkTarget, tgt);
                        }
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
                else if (seg.IsZeroWidth && seg.AlwaysKeep)
                    // A bookmark marker is taken anywhere in [start,end] (BOTH boundaries inclusive), regardless of
                    // the token-driven ownership flags — it is not a diff token, so those flags can't see it and a
                    // boundary bookmark would otherwise be dropped. _claimedZeroWidth keeps it once per model.
                    overlaps = seg.Start >= start && seg.Start <= end;
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

            /// <summary>True for a structural marker that must NEVER be dropped at an op boundary — a
            /// <c>w:bookmarkStart</c>/<c>w:bookmarkEnd</c>. Unlike a content zero-width (footnote ref, drawing,
            /// tab), a bookmark marker is NOT a diff token (rule N3 strips it from the token stream), so the
            /// token-driven <c>includeStart/EndZeroWidth</c> ownership flags are blind to it and a boundary
            /// bookmark would otherwise fall through the cracks (neither op claims it) — orphaning the bookmark
            /// and dangling every cross-reference that targets it. An always-keep marker is taken anywhere in
            /// the op's <c>[start,end]</c> (boundaries inclusive); the per-model <see cref="_claimedZeroWidth"/>
            /// set still guarantees exactly-once emission, and <c>NormalizeBookmarks</c> reconciles the
            /// del/ins-context duplicates a both-sides edit produces.</summary>
            public bool AlwaysKeep { get; init; }
        }
    }
}
