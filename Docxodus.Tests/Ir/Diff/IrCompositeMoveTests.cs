#nullable enable
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// FOLLOW-ON A: native move composition in the consolidate engine. A SINGLE reviewer's non-colliding
/// MoveBlock/MoveModifyBlock renders as a NATIVE move (w:moveFrom/w:moveTo, authored to that reviewer,
/// move-group-id globally namespaced). Colliding moves (move-vs-edit on the same base block; two reviewers
/// moving the same block) STAY lowered to del/ins as a recorded conflict. Uncontested Split/Merge now
/// compose natively too — see <see cref="IrCompositeSplitMergeTests"/>.
/// Gated on <see cref="IrDiffSettings.RenderMoves"/> (DetectMoves=false ⇒ moves still lower).
/// </summary>
public class IrCompositeMoveTests
{
    // Four ≥4-word paragraphs so a reorder is detected as a MoveBlock by the aligner.
    private const string P1 = "First paragraph alpha bravo charlie";
    private const string P2 = "Second paragraph delta echo foxtrot";
    private const string P3 = "Third paragraph golf hotel india";
    private const string P4 = "Fourth paragraph juliet kilo lima";

    private static readonly IrReaderOptions ReadOpts =
        new() { RetainSources = false, RevisionView = RevisionView.Accept };

    private static IrCompositeScript Merge(
        WmlDocument baseDoc, ConflictResolution policy, params (string Author, WmlDocument Doc)[] reviewers)
    {
        var diff = new DocxDiffSettings().ToIrDiffSettings();
        var baseIr = IrReader.Read(baseDoc, ReadOpts);
        var revIr = reviewers.Select(r => (r.Author, IrReader.Read(r.Doc, ReadOpts))).ToList();
        return IrCompositeMerger.Merge(baseIr, revIr, policy, diff);
    }

    private static IrCompositeScript Merge(WmlDocument baseDoc, params (string Author, WmlDocument Doc)[] reviewers)
        => Merge(baseDoc, ConflictResolution.BaseWins, reviewers);

    private static (IrCompositeOp Src, IrCompositeOp Dest)? FindNativeMovePair(IrCompositeScript s, int? gid = null)
    {
        var src = s.Operations.FirstOrDefault(o =>
            o.Op.Kind is IrEditOpKind.MoveBlock or IrEditOpKind.MoveModifyBlock
            && o.Op.IsMoveSource == true && (gid == null || o.Op.MoveGroupId == gid));
        if (src is null) return null;
        var dest = s.Operations.FirstOrDefault(o =>
            o.Op.Kind is IrEditOpKind.MoveBlock or IrEditOpKind.MoveModifyBlock
            && o.Op.IsMoveSource == false && o.Op.MoveGroupId == src.Op.MoveGroupId);
        return dest is null ? null : (src, dest);
    }

    // ---- 1. Single-reviewer non-colliding move → native MoveBlock src+dest sharing one MoveGroupId ----

    [Fact]
    public void Single_reviewer_non_colliding_move_emits_native_move_pair()
    {
        var b = Docs.Para(P1, P2, P3, P4);
        var alice = Docs.Para(P1, P3, P4, P2);   // P2 relocated to the end
        var s = Merge(b, ("Alice", alice));

        var pair = FindNativeMovePair(s);
        Assert.NotNull(pair);
        var (src, dest) = pair!.Value;

        // Source rides LeftAnchor; destination rides RightAnchor; one shared move group id.
        Assert.NotNull(src.Op.LeftAnchor);
        Assert.NotNull(dest.Op.RightAnchor);
        Assert.Equal(src.Op.MoveGroupId, dest.Op.MoveGroupId);
        Assert.NotNull(src.Op.MoveGroupId);

        // Authored to the mover (on both halves).
        Assert.Equal("Alice", src.Author);
        Assert.Equal("Alice", dest.Author);

        // The lowered del/ins shape must NOT appear for this relocated block: no plain DeleteBlock that
        // retains a relocation marker, and the move is genuinely native (a MoveBlock pair exists).
        Assert.DoesNotContain(s.Operations, o => o.Op.Kind == IrEditOpKind.DeleteBlock && o.Op.IsMoveSource == true);
    }

    [Fact]
    public void Single_reviewer_move_accept_relocates_once_reject_equals_base()
    {
        var b = Docs.Para(P1, P2, P3, P4);
        var alice = Docs.Para(P1, P3, P4, P2);
        var diff = new DocxDiffSettings().ToIrDiffSettings();
        var merged = DocxDiff.Consolidate(b, new[] { new DocxDiffReviewer { Document = alice, Author = "Alice" } });

        // reject ≡ base.
        Assert.Equal(Docs.PlainText(b), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));

        // accept relocates P2 to the end exactly once (no dup, no loss): the accepted body equals Alice's.
        Assert.Equal(Docs.PlainText(alice), Docs.PlainText(RevisionAccepter.AcceptRevisions(merged)));

        // The moved paragraph's distinctive words appear exactly once in accept.
        var accepted = Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));
        Assert.Equal(1, CountOccurrences(accepted, "foxtrot"));
    }

    // ---- 2. Two reviewers each move a DIFFERENT block → two move groups with DISTINCT global gids ----

    [Fact]
    public void Two_reviewers_moving_different_blocks_get_distinct_global_move_groups()
    {
        // 6 paragraphs so each reviewer relocates a different one without colliding.
        const string P5 = "Fifth paragraph mike november oscar";
        const string P6 = "Sixth paragraph papa quebec romeo";
        var b = Docs.Para(P1, P2, P3, P4, P5, P6);
        var alice = Docs.Para(P1, P3, P4, P5, P6, P2);   // Alice relocates P2 to end
        var bob = Docs.Para(P5, P1, P2, P3, P4, P6);     // Bob relocates P5 to front
        var s = Merge(b, ("Alice", alice), ("Bob", bob));

        var moveGroups = s.Operations
            .Where(o => o.Op.Kind is IrEditOpKind.MoveBlock or IrEditOpKind.MoveModifyBlock)
            .GroupBy(o => o.Op.MoveGroupId)
            .ToList();

        // Two distinct move groups.
        Assert.Equal(2, moveGroups.Count);
        Assert.Equal(2, moveGroups.Select(g => g.Key).Distinct().Count());

        // Each group has a source + a destination sharing the group's gid.
        foreach (var g in moveGroups)
        {
            Assert.Contains(g, o => o.Op.IsMoveSource == true);
            Assert.Contains(g, o => o.Op.IsMoveSource == false);
        }

        // One group authored to Alice, one to Bob.
        var authors = moveGroups.Select(g => g.First().Author).ToHashSet();
        Assert.Contains("Alice", authors);
        Assert.Contains("Bob", authors);

        // Global namespacing: gids are unique across reviewers (no two groups share an id).
        var gids = moveGroups.Select(g => g.Key).ToList();
        Assert.Equal(gids.Count, gids.Distinct().Count());

        // End-to-end round-trip: reject ≡ base (the 6-paragraph base restored), and accept carries each
        // moved block's distinguishing word EXACTLY ONCE (no duplication from a doubled move half, no loss
        // from a dropped one) — proving the two distinct global move groups render coherently together.
        var merged = DocxDiff.Consolidate(
            b,
            new[]
            {
                new DocxDiffReviewer { Document = alice, Author = "Alice" },
                new DocxDiffReviewer { Document = bob, Author = "Bob" },
            },
            new DocxDiffConsolidateSettings { ConflictResolution = ConflictResolution.BaseWins });

        Assert.Equal(Docs.PlainText(b), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));

        var accepted = Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));
        Assert.Equal(1, CountOccurrences(accepted, "delta"));     // Alice's moved P2 distinguishing word
        Assert.Equal(1, CountOccurrences(accepted, "november")); // Bob's moved P5 distinguishing word
    }

    // ---- 3. Move-vs-edit SAME base block → lowered + recorded conflict; no native move ----

    [Fact]
    public void Move_vs_edit_same_block_lowers_and_records_conflict_no_native_move()
    {
        var b = Docs.Para(P1, P2, P3, P4);
        var mover = Docs.Para(P1, P3, P4, P2);                     // Alice moves P2 to end
        var editor = Docs.Para(P1, "Second paragraph delta echo CHANGED", P3, P4); // Bob edits P2 in place
        var s = Merge(b, ("Alice", mover), ("Bob", editor));

        // No native move op survives for the contested block.
        Assert.DoesNotContain(s.Operations,
            o => o.Op.Kind is IrEditOpKind.MoveBlock or IrEditOpKind.MoveModifyBlock);

        // A conflict is recorded.
        Assert.NotEmpty(s.Conflicts);

        // reject ≡ base under all policies.
        AssertRejectEqualsBaseAllPolicies(b, ("Alice", mover), ("Bob", editor));
    }

    // ---- 4. Two reviewers move the SAME block → contested-relocation conflict; no native move ----

    [Fact]
    public void Two_reviewers_moving_same_block_contested_relocation_no_native_move()
    {
        // A 5-paragraph base with a clear MIDDLE mover (P3): both reviewers relocate P3, so the aligner
        // anchors BOTH move sources at P3's base block (they co-anchor → contested, not native-eligible).
        const string p5 = "Fifth paragraph mike november oscar";
        var b = Docs.Para(P1, P2, P3, P4, p5);
        var alice = Docs.Para(P1, P2, P4, p5, P3);   // Alice moves P3 to end
        var bob = Docs.Para(P3, P1, P2, P4, p5);     // Bob moves P3 to front
        var s = Merge(b, ("Alice", alice), ("Bob", bob));

        // Both movers co-anchor at P3's base source block → NOT native-eligible (touchers != {single}).
        Assert.DoesNotContain(s.Operations,
            o => o.Op.Kind is IrEditOpKind.MoveBlock or IrEditOpKind.MoveModifyBlock);

        // Contested relocation is recorded.
        Assert.NotEmpty(s.Conflicts);

        // reject ≡ base under all policies.
        AssertRejectEqualsBaseAllPolicies(b, ("Alice", alice), ("Bob", bob));
    }

    // ---- 4b. Merge collides with a move of a CONSUMED paragraph → no orphaned native move ----

    [Fact]
    public void Merge_collides_with_move_of_consumed_paragraph_no_orphan_move()
    {
        // Base: four ≥4-word paragraphs. L1/L2 carry sentence boundaries so the aligner detects a MERGE
        // when a reviewer combines them into one paragraph (the proven merge-detection shape).
        const string L1 = "Alpha alpha words here.";
        const string L2 = "Beta beta words here.";
        const string L3 = "Gamma gamma words here.";
        const string L4 = "Delta delta words here.";
        var b = Docs.Para(L1, L2, L3, L4);

        // Alice MERGES L1 + L2 into a single paragraph. A MergeBlock consumes L1 and L2 as its
        // SplitMergeAnchors (left anchors); MergeBlock's own LeftAnchor is null.
        var alice = Docs.Para(L1 + " " + L2, L3, L4);
        // Bob MOVES L1 to the end (a relocation whose move-SOURCE anchors at L1's base block).
        var bob = Docs.Para(L2, L3, L4, L1);

        // Precondition: the aligner MUST actually detect Alice's edit as a MergeBlock (otherwise this
        // fixture would not exercise the PlanMoves MergeBlock-toucher path at all).
        var aliceJson = DocxDiff.GetEditScriptJson(b, alice);
        Assert.Contains("\"kind\": \"MergeBlock\"", aliceJson);

        var s = Merge(b, ("Alice", alice), ("Bob", bob));

        // THE BUG (HEAD 8e33518): PlanMoves builds touchersByBaseAnchor from each reviewer's raw ops keyed
        // on op.LeftAnchor. A MergeBlock has LeftAnchor == null and carries its consumed paragraphs in
        // SplitMergeAnchors, so L1/L2 are NEVER registered as touched by Alice. Bob's move of L1 then looks
        // like a sole-toucher → it is misclassified NATIVE, emitting a w:moveTo destination with no coherent
        // paired w:moveFrom once Alice's merge (lowered to deletes of L1+L2) resolves → an orphaned half-move.
        // After the fix the merge registers L1/L2 as touched, so Bob's move of L1 is no longer a sole-toucher
        // → it lowers to del/ins. (The two deletes of L1 — Alice's merge-delete and Bob's relocation-delete —
        // are genuine consensus that L1 leaves its origin, so no conflict is recorded; the contract's
        // collision guarantee is "lowered / no content lost", and content IS fully preserved below.)
        Assert.DoesNotContain(s.Operations,
            o => o.Op.Kind is IrEditOpKind.MoveBlock or IrEditOpKind.MoveModifyBlock);

        // End-to-end under every policy: reject ≡ base, NO orphaned move markup (equal — indeed zero — counts
        // of moveFrom/moveTo range starts), and NO content lost (every base paragraph's distinguishing word
        // survives accept).
        foreach (var policy in new[] { ConflictResolution.BaseWins, ConflictResolution.FirstReviewerWins, ConflictResolution.StackAll })
        {
            var merged = DocxDiff.Consolidate(
                b,
                new[]
                {
                    new DocxDiffReviewer { Document = alice, Author = "Alice" },
                    new DocxDiffReviewer { Document = bob, Author = "Bob" },
                },
                new DocxDiffConsolidateSettings { ConflictResolution = policy });

            // reject restores the base exactly.
            Assert.Equal(Docs.PlainText(b), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));

            // No orphaned half-move: moveFrom/moveTo range starts are balanced, and zero — the move lowered.
            var xml = Docs.MainPartXml(merged);
            int moveFrom = CountOccurrences(xml, "moveFromRangeStart");
            int moveTo = CountOccurrences(xml, "moveToRangeStart");
            Assert.Equal(moveFrom, moveTo);
            Assert.Equal(0, moveTo);

            // No content lost: every base paragraph's distinguishing word is present on accept.
            var accept = Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));
            Assert.Contains("Alpha", accept);  // L1 (relocated by Bob and merged by Alice)
            Assert.Contains("Beta", accept);   // L2 (merged into L1 by Alice)
            Assert.Contains("Gamma", accept);  // L3 untouched
            Assert.Contains("Delta", accept);  // L4 untouched
        }
    }

    // ---- 5. DetectMoves=false ⇒ no native MoveBlock op (lowered) — gate test ----

    [Fact]
    public void DetectMoves_off_lowers_move_no_native_op()
    {
        var b = Docs.Para(P1, P2, P3, P4);
        var alice = Docs.Para(P1, P3, P4, P2);
        var diff = new DocxDiffSettings { DetectMoves = false }.ToIrDiffSettings();
        var baseIr = IrReader.Read(b, ReadOpts);
        var revIr = new[] { ("Alice", IrReader.Read(alice, ReadOpts)) }.ToList();
        var s = IrCompositeMerger.Merge(baseIr, revIr, ConflictResolution.BaseWins, diff);

        Assert.DoesNotContain(s.Operations,
            o => o.Op.Kind is IrEditOpKind.MoveBlock or IrEditOpKind.MoveModifyBlock);
        // It still round-trips (lowered del/ins).
        var merged = DocxDiff.Consolidate(b, new[] { new DocxDiffReviewer { Document = alice, Author = "Alice" } },
            new DocxDiffConsolidateSettings { Diff = new DocxDiffSettings { DetectMoves = false } });
        Assert.Equal(Docs.PlainText(b), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
    }

    // ---- 9. reject ≡ base for all 3 policies with a native move; verifier over the native-move case ----

    [Fact]
    public void Native_move_reject_equals_base_all_policies_and_verifier_holds()
    {
        var b = Docs.Para(P1, P2, P3, P4);
        var alice = Docs.Para(P1, P3, P4, P2);

        foreach (var policy in new[] { ConflictResolution.BaseWins, ConflictResolution.FirstReviewerWins, ConflictResolution.StackAll })
        {
            var merged = DocxDiff.Consolidate(b, new[] { new DocxDiffReviewer { Document = alice, Author = "Alice" } },
                new DocxDiffConsolidateSettings { ConflictResolution = policy });
            Assert.Equal(Docs.PlainText(b), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
            var acceptedText = Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));
            IrCompositeVerifier.Verify(b, new[] { ("Alice", alice) }, policy, acceptedText);
        }
    }

    // ----------------------------------------------------------------- helpers

    private static void AssertRejectEqualsBaseAllPolicies(WmlDocument b, params (string Author, WmlDocument Doc)[] reviewers)
    {
        var dd = reviewers.Select(r => new DocxDiffReviewer { Document = r.Doc, Author = r.Author }).ToList();
        foreach (var policy in new[] { ConflictResolution.BaseWins, ConflictResolution.FirstReviewerWins, ConflictResolution.StackAll })
        {
            var merged = DocxDiff.Consolidate(b, dd, new DocxDiffConsolidateSettings { ConflictResolution = policy });
            Assert.Equal(Docs.PlainText(b), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
        }
    }

    private static int CountOccurrences(string haystack, string needle)
    {
        int count = 0, idx = 0;
        while ((idx = haystack.IndexOf(needle, idx, System.StringComparison.Ordinal)) >= 0)
        {
            count++;
            idx += needle.Length;
        }
        return count;
    }
}
