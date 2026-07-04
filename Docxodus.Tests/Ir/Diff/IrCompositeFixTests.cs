#nullable enable

using System.Collections.Generic;
using System.Linq;
using Docxodus;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Regression tests for two confirmed N-way consolidate composite-merge bugs:
/// <list type="bullet">
/// <item><b>B2</b> — under <see cref="ConflictResolution.StackAll"/>, a delete-vs-edit block conflict
/// stacked BOTH a base-restoring DeleteBlock and a base-restoring ModifyBlock, so the merged document's
/// reject produced the base paragraph TWICE (4 paras vs 3). The fix emits at most one base-anchored op
/// for the contested base block and renders the other competitors as base-ANCHORLESS InsertBlock ops
/// (contributing nothing on reject), so reject ≡ base under every policy.</item>
/// <item><b>B4</b> — when 2+ reviewers edit the same base TABLE (even DIFFERENT cells), the phantom
/// "all ops identical" consensus (table <c>BlockResultText</c> returned "" for every reviewer) silently
/// dropped all but the first reviewer's table and recorded NO conflict (data loss). The fix makes table
/// ModifyBlocks fall through to the block-level conflict branch, so the disagreement is recorded and the
/// reject restores the base table.</item>
/// </list>
/// </summary>
public class IrCompositeFixTests
{
    private static WmlDocument Consolidate(
        WmlDocument baseDoc, ConflictResolution policy,
        params (string Author, WmlDocument Doc)[] reviewers)
        => DocxDiff.Consolidate(
            baseDoc,
            reviewers.Select(r => new DocxDiffReviewer { Author = r.Author, Document = r.Doc }).ToList(),
            new DocxDiffConsolidateSettings { ConflictResolution = policy });

    private static IReadOnlyList<DocxDiffConflict> Conflicts(
        WmlDocument baseDoc, ConflictResolution policy,
        params (string Author, WmlDocument Doc)[] reviewers)
        => DocxDiff.GetConflicts(
            baseDoc,
            reviewers.Select(r => new DocxDiffReviewer { Author = r.Author, Document = r.Doc }).ToList(),
            new DocxDiffConsolidateSettings { ConflictResolution = policy });

    // ---- B2: StackAll delete-vs-edit must reject to base (no duplication) ----

    [Fact]
    public void Delete_vs_edit_stackAll_rejects_to_base()
    {
        var baseDoc = Docs.Para("First", "Second stays interesting", "Third");
        var alice = Docs.Para("First", "Third");                            // deletes para2
        var bob = Docs.Para("First", "Second EDITED interesting", "Third"); // edits para2 in place

        var merged = Consolidate(baseDoc, ConflictResolution.StackAll, ("Alice", alice), ("Bob", bob));

        // Reject must restore the base EXACTLY — no duplicated paragraph.
        Assert.Equal(Docs.PlainText(baseDoc),
            Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));

        // Accept must keep Bob's in-place edit (StackAll surfaces both competitors).
        Assert.Contains("EDITED", Docs.PlainText(RevisionAccepter.AcceptRevisions(merged)));

        // The delete-vs-edit disagreement is recorded.
        Assert.NotEmpty(Conflicts(baseDoc, ConflictResolution.StackAll, ("Alice", alice), ("Bob", bob)));
    }

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    public void Delete_vs_edit_baseWins_and_firstWins_still_clean(ConflictResolution policy)
    {
        var baseDoc = Docs.Para("First", "Second stays interesting", "Third");
        var alice = Docs.Para("First", "Third");
        var bob = Docs.Para("First", "Second EDITED interesting", "Third");

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));

        Assert.Equal(Docs.PlainText(baseDoc),
            Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
    }

    // ---- B4: multi-reviewer table edits must record a conflict, never silently drop ----

    private static string Cell(string text) =>
        $"<w:tc><w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p></w:tc>";

    private static string Row(params string[] cells) => $"<w:tr>{string.Concat(cells)}</w:tr>";

    private static string Table(params string[] rows) =>
        $"<w:tbl><w:tblPr/><w:tblGrid/>{string.Concat(rows)}</w:tbl>";

    /// <summary>A 2x2 table base; one paragraph above so the body has stable surrounding structure.</summary>
    private static WmlDocument TableBase() => IrTestDocuments.FromBodyXml(
        "<w:p><w:r><w:t>lead</w:t></w:r></w:p>" +
        Table(
            Row(Cell("a one"), Cell("b two")),
            Row(Cell("c three"), Cell("d four"))));

    private static WmlDocument TableVariant(string c00, string c01, string c10, string c11) =>
        IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>lead</w:t></w:r></w:p>" +
            Table(
                Row(Cell(c00), Cell(c01)),
                Row(Cell(c10), Cell(c11))));

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Multireviewer_disjoint_table_cell_edits_compose_no_drop(ConflictResolution policy)
    {
        var baseDoc = TableBase();
        var alice = TableVariant("a ALICE", "b two", "c three", "d four"); // edits cell (0,0)
        var bob = TableVariant("a one", "b two", "c three", "d BOB");      // edits cell (1,1) — disjoint

        // FOLLOW-ON B: disjoint cross-reviewer cell edits now COMPOSE inline — both land, NO conflict
        // (the prior behavior recorded a whole-table conflict; the deeper IrCompositeTableTests cover the
        // composed-content + per-reviewer attribution in detail).
        Assert.Empty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var accept = Docs.StructuralBody(RevisionAccepter.AcceptRevisions(merged));
        Assert.Contains("a ALICE", accept);
        Assert.Contains("d BOB", accept);

        // reject ≡ base must hold under every policy.
        Assert.Equal(Docs.StructuralBody(baseDoc),
            Docs.StructuralBody(RevisionProcessor.RejectRevisions(merged)));
    }

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Multireviewer_same_table_cell_edits_record_conflict(ConflictResolution policy)
    {
        var baseDoc = TableBase();
        var alice = TableVariant("a ALICE", "b two", "c three", "d four"); // edits cell (0,0)
        var bob = TableVariant("a BOB", "b two", "c three", "d four");     // edits SAME cell differently

        Assert.True(
            Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)).Count >= 1,
            "Multi-reviewer same-cell table edits were not recorded as a conflict.");

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        Assert.Equal(Docs.StructuralBody(baseDoc),
            Docs.StructuralBody(RevisionProcessor.RejectRevisions(merged)));
    }

    // ---- B4 generalization: NON-paragraph, non-table modifies (section breaks / opaque blocks) ----
    //
    // BlockResultText returns "" for ANY non-paragraph block, so before the generalization the empty-text
    // "all ops identical" consensus also fired for two reviewers DIFFERENTLY editing the same base section
    // break (a standalone body w:sectPr → IrSectionBreak block; ModifyBlock with BOTH TokenDiff AND TableDiff
    // null) — silently dropping all but the first reviewer and recording NO conflict. The generalized
    // IsUncomparableModify (op.Kind == ModifyBlock && op.TokenDiff == null) now covers tables AND
    // section-break/opaque modifies, routing them to the recorded block-level conflict. A paragraph modify
    // always carries TokenDiff != null and is unaffected (still composes / reaches genuine text consensus).

    /// <summary>A two-section base: a standalone first-section <c>w:sectPr</c> (its own body block) whose
    /// page size differs from the trailing section's, so it reads as an <see cref="Docxodus.Ir.IrSectionBreak"/>
    /// body block both reviewers can edit.</summary>
    private static WmlDocument TwoSectionBase() => IrTestDocuments.FromBodyXml(
        "<w:p><w:r><w:t>first section body</w:t></w:r></w:p>" +
        "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>" +
        "<w:p><w:r><w:t>second section body</w:t></w:r></w:p>" +
        "<w:sectPr><w:pgSz w:w=\"15840\" w:h=\"12240\"/></w:sectPr>");

    /// <summary>The same shape as <see cref="TwoSectionBase"/> but with the FIRST (standalone) section's
    /// page size overridden to <paramref name="firstSectionW"/>×<paramref name="firstSectionH"/>, so a
    /// reviewer "edits" the standalone section break.</summary>
    private static WmlDocument TwoSectionVariant(string firstSectionW, string firstSectionH) =>
        IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>first section body</w:t></w:r></w:p>" +
            $"<w:sectPr><w:pgSz w:w=\"{firstSectionW}\" w:h=\"{firstSectionH}\"/></w:sectPr>" +
            "<w:p><w:r><w:t>second section body</w:t></w:r></w:p>" +
            "<w:sectPr><w:pgSz w:w=\"15840\" w:h=\"12240\"/></w:sectPr>");

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Multireviewer_section_break_edits_record_conflict_not_silent_drop(ConflictResolution policy)
    {
        var baseDoc = TwoSectionBase();
        var alice = TwoSectionVariant("11000", "14000"); // changes the standalone section's page size one way
        var bob = TwoSectionVariant("9000", "13000");    // changes it a DIFFERENT way → genuine disagreement

        // Both reviewers' section-break edits must be SEEN: a recorded conflict, not a silent drop.
        Assert.True(
            Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)).Count >= 1,
            "Multi-reviewer section-break edits were silently dropped (conflictCount == 0).");

        // reject ≡ base must hold for the section-break-conflict output under every policy.
        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        Assert.Equal(Docs.StructuralBody(baseDoc),
            Docs.StructuralBody(RevisionProcessor.RejectRevisions(merged)));
    }

    // ---- B5: a reviewer's STRUCTURAL ops (move / merge / split) must not silently drop content ----
    //
    // The composite merger groups each reviewer's pairwise ops by base anchor (op.LeftAnchor) and routes
    // right-only InsertBlocks by preceding anchor. A null-LeftAnchor STRUCTURAL op — a MoveBlock/MoveModify
    // DESTINATION (RightAnchor set, LeftAnchor null) or a MergeBlock (RightAnchor set, SplitMergeAnchors =
    // left anchors, LeftAnchor null) — reaches NEITHER path, so it was DROPPED: the move-TO / merge result
    // vanished on accept with no conflict recorded (reject ≡ base still held — no corruption — but content
    // was lost). The fix LOWERS every reviewer Move/Split/Merge op to Insert/Delete/Modify before grouping,
    // so the moved/merged content survives accept (rendered as del/ins, NOT native move/split/merge markup
    // in v1; content fully preserved and round-trips). Paragraphs here are ≥4 words so move detection fires.

    private static string PlainAccept(WmlDocument merged) =>
        Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));

    private static int ParagraphCount(string plainText) =>
        plainText.Split('\n').Count(line => line.Trim().Length > 0);

    /// <summary>
    /// BUG 1 (move-disjoint). Base [P1,P2,P3,P4]; Alice relocates P2 to the end; Bob edits a word in P1.
    /// Alice and Bob touch DIFFERENT base blocks (P2 vs P1), so there is no block conflict and BOTH edits
    /// apply under every policy. On accept the moved paragraph P2 must survive (it was dropped before the
    /// lowering fix) and every paragraph must be present; reject ≡ base. Parameterized over all three
    /// policies (the disjoint edits make the result policy-invariant).
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Move_disjoint_accept_preserves_moved_paragraph(ConflictResolution policy)
    {
        const string p1 = "First paragraph alpha bravo";
        const string p2 = "Second paragraph charlie delta";
        const string p3 = "Third paragraph echo foxtrot";
        const string p4 = "Fourth paragraph golf hotel";

        var baseDoc = Docs.Para(p1, p2, p3, p4);
        var alice = Docs.Para(p1, p3, p4, p2);                       // P2 relocated to the end (a move)
        var bob = Docs.Para("First paragraph alpha EDITED", p2, p3, p4); // edits a word in P1, in place

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var accept = PlainAccept(merged);

        // The moved paragraph P2 must NOT be lost on accept (the key bug).
        Assert.Contains("charlie delta", accept);
        // Every base paragraph survives (P1 in Bob's edited form, P2/P3/P4 intact).
        Assert.Contains("First paragraph alpha", accept);
        Assert.Contains("EDITED", accept);
        Assert.Contains("Third paragraph echo foxtrot", accept);
        Assert.Contains("Fourth paragraph golf hotel", accept);
        // Paragraph COUNT preserved (no paragraph vanished).
        Assert.True(ParagraphCount(accept) >= 4,
            $"expected >= 4 paragraphs on accept, got {ParagraphCount(accept)}: [{accept}]");

        // reject ≡ base.
        Assert.Equal(Docs.PlainText(baseDoc),
            Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
    }

    /// <summary>
    /// BUG 2 (two-reviewers-move-same-block). Base [P1,P2,P3,P4,P5]; both reviewers relocate the SAME block
    /// P2 to DIFFERENT mid positions (the rigid P1,P3,P4,P5 spine makes P2 the unambiguous mover for both,
    /// so the contention is over one base block). Before the lowering fix the move DESTINATIONS were dropped
    /// (P2's relocations vanished on accept, no conflict). After lowering, both reviewers' relocating deletes
    /// of base P2 collide → a recorded placement conflict, while each reviewer's relocating insert survives
    /// independently — so under every policy NO paragraph is lost on accept and reject ≡ base.
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Two_reviewers_move_same_block_no_loss(ConflictResolution policy)
    {
        const string p1 = "First paragraph alpha bravo";
        const string p2 = "Second paragraph charlie delta";
        const string p3 = "Third paragraph echo foxtrot";
        const string p4 = "Fourth paragraph golf hotel";
        const string p5 = "Fifth paragraph india juliet";

        var baseDoc = Docs.Para(p1, p2, p3, p4, p5);
        var alice = Docs.Para(p1, p3, p2, p4, p5);  // P2 relocated between P3 and P4
        var bob = Docs.Para(p1, p3, p4, p2, p5);    // P2 relocated between P4 and P5 (a different spot)

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var accept = PlainAccept(merged);

        // No paragraph lost on accept — every base paragraph's content present (P2 relocated, not dropped).
        Assert.Contains("First paragraph alpha bravo", accept);
        Assert.Contains("Second paragraph charlie delta", accept);
        Assert.Contains("Third paragraph echo foxtrot", accept);
        Assert.Contains("Fourth paragraph golf hotel", accept);
        Assert.Contains("Fifth paragraph india juliet", accept);

        // reject ≡ base under every policy.
        Assert.Equal(Docs.PlainText(baseDoc),
            Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));

        // The same base block P2 is contested by two reviewers (a placement disagreement) → recorded conflict.
        Assert.True(
            Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)).Count > 0,
            "Two reviewers relocating the same base block must record a conflict (not silently drop a move).");
    }

    /// <summary>
    /// BUG 3 (merge-vs-edit). Base two paragraphs; Alice merges them into one; Bob edits a word in one.
    /// Before the lowering fix the MergeBlock (null LeftAnchor) was DROPPED entirely — accept showed only
    /// Bob's edit (the original two paragraphs, Alice's merge silently gone) with conflicts==0 (data loss).
    /// After lowering, Alice's merge becomes deletes (of the base blocks) + an insert (the merged result);
    /// the deleted base block Bob also edited is contested → a recorded conflict, and Alice's merged
    /// paragraph survives accept under EVERY policy (the merged result is a base-anchorless insert that no
    /// block-conflict policy suppresses). Either way: NO silent drop. reject ≡ base. Parameterized over all
    /// three policies.
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Merge_vs_edit_accept_preserves_merge(ConflictResolution policy)
    {
        var baseDoc = Docs.Para("Alpha alpha words here", "Beta beta words here");
        var alice = Docs.Para("Alpha alpha words here Beta beta words here"); // merges the two into one
        var bob = Docs.Para("Alpha alpha EDITED here", "Beta beta words here"); // edits a word in para 1

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var accept = PlainAccept(merged);

        // Alice's merge must NOT be silently dropped: a conflict is recorded (the merge contends with Bob's
        // edit of the same base block). conflicts == 0 with the merge missing was the original data-loss bug.
        Assert.True(
            Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)).Count > 0,
            "Alice's merge was silently dropped (no conflict recorded, merge absent from output).");

        // Alice's merged single paragraph (both halves contiguous on ONE line) survives accept — a state
        // unreachable when the merge is dropped (then the two base paragraphs stay on separate lines).
        Assert.True(
            accept.Split('\n').Any(line => line.Contains("Alpha alpha") && line.Contains("Beta beta words here")),
            $"Alice's merged paragraph (both halves on one line) absent from accept: [{accept}]");

        // reject ≡ base — no corruption regardless of how the merge is composed.
        Assert.Equal(Docs.PlainText(baseDoc),
            Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
    }

    /// <summary>
    /// BUG 4 (split-vs-edit regression guard). Base one paragraph; Alice splits it into two; Bob edits a
    /// word. Accept must preserve all content; reject ≡ base. SplitBlock carries a LeftAnchor so it reached
    /// the byBase path even before the fix — this guards that lowering split to delete+inserts does not
    /// regress content preservation. Parameterized over all three policies (the split's two inserted
    /// members are base-anchorless and survive accept under every policy).
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Split_vs_edit_accept_preserves_content(ConflictResolution policy)
    {
        var baseDoc = Docs.Para("Alpha alpha words here Beta beta words here");
        var alice = Docs.Para("Alpha alpha words here", "Beta beta words here"); // splits into two
        var bob = Docs.Para("Alpha alpha EDITED here Beta beta words here");     // edits a word in place

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var accept = PlainAccept(merged);

        // Split content survives accept.
        Assert.Contains("Alpha alpha", accept);
        Assert.Contains("Beta beta words here", accept);

        // reject ≡ base.
        Assert.Equal(Docs.PlainText(baseDoc),
            Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
    }

    // ---- Fix 3: novel-branch committed tests (3-reviewer contested relocation + MoveModify) ----

    /// <summary>
    /// THREE reviewers each relocate the SAME base block P2 to THREE DIFFERENT positions. This drives the
    /// contested-relocation branch with 3 competing move-source deletes anchored at P2 (and 3 independently
    /// routed relocating inserts). Under EVERY policy: no content loss on accept (every base paragraph's
    /// content present — P2 relocated, never dropped), a PLACEMENT conflict is recorded (the reviewers agree
    /// P2 leaves its origin but disagree where it lands — deliberately NOT policy-resolved into the op
    /// stream), and reject ≡ base. The rigid P1,P3,P4,P5 spine makes P2 the unambiguous mover for all three.
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Three_reviewers_move_same_block_records_conflict_no_loss(ConflictResolution policy)
    {
        const string p1 = "First paragraph alpha bravo";
        const string p2 = "Second paragraph charlie delta";
        const string p3 = "Third paragraph echo foxtrot";
        const string p4 = "Fourth paragraph golf hotel";
        const string p5 = "Fifth paragraph india juliet";

        var baseDoc = Docs.Para(p1, p2, p3, p4, p5);
        var alice = Docs.Para(p1, p3, p2, p4, p5);  // P2 → between P3 and P4
        var bob = Docs.Para(p1, p3, p4, p2, p5);    // P2 → between P4 and P5
        var carol = Docs.Para(p1, p3, p4, p5, p2);  // P2 → the very end

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob), ("Carol", carol));
        var accept = PlainAccept(merged);

        // No content loss on accept — every base paragraph present (P2's content relocated, not dropped).
        Assert.Contains("First paragraph alpha bravo", accept);
        Assert.Contains("Second paragraph charlie delta", accept);
        Assert.Contains("Third paragraph echo foxtrot", accept);
        Assert.Contains("Fourth paragraph golf hotel", accept);
        Assert.Contains("Fifth paragraph india juliet", accept);

        // A placement conflict is recorded (three reviewers disagree on P2's destination). The recorded
        // conflict identifies the contested block by its LEFT (base) content — see ContestedBlockText.
        var conflicts = Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob), ("Carol", carol));
        Assert.True(conflicts.Count > 0,
            "Three reviewers relocating the same base block must record a placement conflict.");

        // Fix 3 (competitor text): the contested-relocation conflict's competitors must NAME the contested
        // block via its LEFT (base) content — before the fix every competitor's ResultText was "" (the
        // lowered move-source DeleteBlock has no RightAnchor), leaving the conflict unable to say WHICH
        // block is contested. Find the placement conflict (a block-span [0,0) conflict whose competitors all
        // carry the relocated base block P2's text) and assert it identifies P2.
        var placement = conflicts.FirstOrDefault(c =>
            c.Competitors.Count >= 2 &&
            c.Competitors.All(comp => comp.ResultText.Contains("Second paragraph charlie delta")));
        Assert.True(placement is not null,
            "Contested-relocation conflict competitors must carry the relocated base block's LEFT content " +
            "(identifying P2), not empty result text.");

        // reject ≡ base under every policy.
        Assert.Equal(Docs.PlainText(baseDoc),
            Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
    }

    /// <summary>
    /// MoveModify (move-AND-edit by one reviewer). Base [P1,P2,P3,P4]; Alice relocates P2 to the end AND
    /// edits a word inside it in the same move (a fuzzy move-and-edit → an IR MovedModified alignment, whose
    /// destination op carries an in-move TokenDiff). The lowering drops the in-move TokenDiff and inserts the
    /// whole EDITED destination block while the paired source delete removes the original — so on accept the
    /// moved+edited content lands and the original is removed; reject ≡ base. Single-reviewer, so policy is
    /// irrelevant (StackAll exercises the single-touched path).
    /// </summary>
    [Fact]
    public void MoveModify_single_reviewer_lands_moved_and_edited_content()
    {
        const string p1 = "First paragraph alpha bravo";
        const string p2 = "Second paragraph charlie delta";
        const string p3 = "Third paragraph echo foxtrot";
        const string p4 = "Fourth paragraph golf hotel";

        var baseDoc = Docs.Para(p1, p2, p3, p4);
        // P2 relocated to the end AND a word edited inside it (charlie -> CHARLIE) — a move-and-edit.
        var alice = Docs.Para(p1, p3, p4, "Second paragraph CHARLIE delta");

        var merged = Consolidate(baseDoc, ConflictResolution.StackAll, ("Alice", alice));
        var accept = PlainAccept(merged);

        // The moved+edited content lands on accept (the edited token present, the moved block survives).
        Assert.Contains("CHARLIE", accept);
        Assert.Contains("Second paragraph", accept);
        // The ORIGINAL (un-edited) P2 is removed — its pre-edit token must not survive on accept.
        Assert.DoesNotContain("charlie delta", accept);
        // The untouched paragraphs survive.
        Assert.Contains("First paragraph alpha bravo", accept);
        Assert.Contains("Third paragraph echo foxtrot", accept);
        Assert.Contains("Fourth paragraph golf hotel", accept);

        // reject ≡ base.
        Assert.Equal(Docs.PlainText(baseDoc),
            Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
    }
}
