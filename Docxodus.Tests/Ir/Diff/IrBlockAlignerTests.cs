#nullable enable

using System.Linq;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.1 Task 2 tests for <see cref="IrBlockAligner"/>: identity, single edit, insert/delete at
/// head/middle/tail, pure move (the headline capability), move + unrelated edit, format-only,
/// boilerplate non-false-move, adjacent swap, table-as-unit, empty docs, determinism, and a shared
/// invariants check applied to every case's result.
/// </summary>
/// <remarks>
/// Documents are built via <see cref="IrTestDocuments"/> + <see cref="IrReader"/> read with
/// <c>RetainSources = false</c> — the aligner needs only the reader-computed hashes, no provenance.
/// </remarks>
public class IrBlockAlignerTests
{
    private static readonly IrReaderOptions NoSources = new() { RetainSources = false };
    private static readonly IrDiffSettings Default = new();

    private static IrDocument Doc(params string[] paragraphTexts) =>
        IrReader.Read(IrTestDocuments.Create(paragraphTexts), NoSources);

    private static IrDocument FromXml(string bodyInnerXml) =>
        IrReader.Read(IrTestDocuments.FromBodyXml(bodyInnerXml), NoSources);

    private static IrBlockAlignment Align(IrDocument l, IrDocument r) =>
        IrBlockAligner.Align(l, r, Default);

    private static IrBlockAlignment Align(IrDocument l, IrDocument r, IrDiffSettings settings) =>
        IrBlockAligner.Align(l, r, settings);

    /// <summary>The aligner invariants the plan pins — see <see cref="IrAlignmentAsserts"/>.</summary>
    private static void AssertInvariants(IrDocument left, IrDocument right, IrBlockAlignment a) =>
        IrAlignmentAsserts.AssertInvariants(left, right, a);

    private static int Count(IrBlockAlignment a, IrAlignmentKind k) => IrAlignmentAsserts.Count(a, k);

    // ------------------------------------------------------------------ identity / edit

    [Fact]
    public void Identity_all_unchanged()
    {
        var l = Doc("alpha", "beta", "gamma");
        var r = Doc("alpha", "beta", "gamma");
        var a = Align(l, r);

        Assert.All(a.Entries, e => Assert.Equal(IrAlignmentKind.Unchanged, e.Kind));
        Assert.Equal(3, a.Entries.Count);
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Single_text_edit_is_modified()
    {
        // "beta" → "beta-edited" shares the word "beta", so the lone 1×1 residue force-pairs as
        // Modified. (Previously "BETA-edited": since the Word-matcher junction calibration the 1×1
        // residue requires ≥1 shared WORD token — the oracle keeps a FULL paragraph rewrite as
        // separate ins/del paragraphs ("24" ↔ "1.5 Line Spacing Demo"), and case-sensitive matching
        // makes "BETA" a full rewrite of "beta". The zero-shared case is pinned by
        // Full_rewrite_1x1_residue_stays_delete_insert below.)
        var l = Doc("alpha", "beta", "gamma");
        var r = Doc("alpha", "beta-edited", "gamma");
        var a = Align(l, r);

        Assert.Equal(2, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(1, Count(a, IrAlignmentKind.Modified));
        Assert.Equal(0, Count(a, IrAlignmentKind.Moved));
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Full_rewrite_1x1_residue_stays_delete_insert()
    {
        // Word-oracle data point: a replace-gap paragraph with ZERO shared word tokens is NOT paired
        // by Word's matcher — the corpus oracle for "24" ↔ "1.5 Line Spacing Demo" keeps an
        // ins-marked paragraph (right pPr) and a separate del-marked paragraph (left pPr). Forcing
        // the pair would token-interleave two unrelated texts inside one paragraph.
        var l = Doc("alpha", "twenty-four", "gamma");
        var r = Doc("alpha", "completely unrelated replacement text", "gamma");
        var a = Align(l, r);

        Assert.Equal(2, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(0, Count(a, IrAlignmentKind.Modified));
        Assert.Equal(1, Count(a, IrAlignmentKind.Deleted));
        Assert.Equal(1, Count(a, IrAlignmentKind.Inserted));
        var bodyRewrite = a.Entries.Where(e => e.Kind is IrAlignmentKind.Deleted or IrAlignmentKind.Inserted)
            .ToList();
        Assert.Equal(2, bodyRewrite.Count);
        Assert.NotNull(bodyRewrite[0].BodyFullRewriteGroupId);
        Assert.Equal(bodyRewrite[0].BodyFullRewriteGroupId, bodyRewrite[1].BodyFullRewriteGroupId);

        // Raw-block callers are nested-scope machinery (cells, notes, headers, textboxes): they use
        // the identical alignment classifications but never carry the body renderer provenance.
        var nested = IrBlockAligner.AlignBlocks(l.Body.Blocks, r.Body.Blocks, Default);
        Assert.All(nested.Entries, e => Assert.Null(e.BodyFullRewriteGroupId));
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Tail_full_rewrite_1x1_has_no_separate_paragraph_marker()
    {
        // The trailing section-break sentinel is not a body continuation: Word keeps a tail rewrite
        // as one mixed paragraph, so the explicit renderer provenance must remain absent.
        var l = Doc("anchor title", "obsolete amber stanza");
        var r = Doc("anchor title edited", "fresh quantum clause");
        var a = Align(l, r);

        Assert.Equal(1, Count(a, IrAlignmentKind.Deleted));
        Assert.Equal(1, Count(a, IrAlignmentKind.Inserted));
        Assert.All(a.Entries.Where(e => e.Kind is IrAlignmentKind.Deleted or IrAlignmentKind.Inserted),
            e => Assert.Null(e.BodyFullRewriteGroupId));
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Residue_pairing_respects_case_sensitivity()
    {
        // The lexical evidence used for a 1×1 residue must honor the same CaseInsensitive policy as
        // token MatchKeys. Default comparison treats beta → BETA-edited as a full rewrite; enabling
        // case-insensitive comparison makes beta the shared lexical evidence and pairs the paragraph.
        var l = Doc("alpha", "beta", "gamma");
        var r = Doc("alpha", "BETA-edited", "gamma");

        var sensitive = Align(l, r);
        Assert.Equal(0, Count(sensitive, IrAlignmentKind.Modified));
        Assert.Equal(1, Count(sensitive, IrAlignmentKind.Deleted));
        Assert.Equal(1, Count(sensitive, IrAlignmentKind.Inserted));
        AssertInvariants(l, r, sensitive);

        var insensitive = Align(l, r, new IrDiffSettings { CaseInsensitive = true });
        Assert.Equal(1, Count(insensitive, IrAlignmentKind.Modified));
        Assert.Equal(0, Count(insensitive, IrAlignmentKind.Deleted));
        Assert.Equal(0, Count(insensitive, IrAlignmentKind.Inserted));
        AssertInvariants(l, r, insensitive);
    }

    // ------------------------------------------------------------------ insert

    [Fact]
    public void Insert_at_start()
    {
        var l = Doc("alpha", "beta");
        var r = Doc("NEW", "alpha", "beta");
        var a = Align(l, r);

        Assert.Equal(2, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(1, Count(a, IrAlignmentKind.Inserted));
        Assert.Equal(IrAlignmentKind.Inserted, a.Entries[0].Kind);
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Insert_in_middle()
    {
        var l = Doc("alpha", "beta");
        var r = Doc("alpha", "NEW", "beta");
        var a = Align(l, r);

        Assert.Equal(2, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(1, Count(a, IrAlignmentKind.Inserted));
        Assert.Equal(IrAlignmentKind.Inserted, a.Entries[1].Kind);
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Insert_at_end()
    {
        var l = Doc("alpha", "beta");
        var r = Doc("alpha", "beta", "NEW");
        var a = Align(l, r);

        Assert.Equal(2, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(1, Count(a, IrAlignmentKind.Inserted));
        Assert.Equal(IrAlignmentKind.Inserted, a.Entries[^1].Kind);
        AssertInvariants(l, r, a);
    }

    // ------------------------------------------------------------------ delete

    [Fact]
    public void Delete_at_start()
    {
        var l = Doc("alpha", "beta", "gamma");
        var r = Doc("beta", "gamma");
        var a = Align(l, r);

        Assert.Equal(2, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(1, Count(a, IrAlignmentKind.Deleted));
        Assert.Equal(IrAlignmentKind.Deleted, a.Entries[0].Kind); // left-anchored: front deletion first
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Delete_in_middle()
    {
        var l = Doc("alpha", "beta", "gamma");
        var r = Doc("alpha", "gamma");
        var a = Align(l, r);

        Assert.Equal(2, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(1, Count(a, IrAlignmentKind.Deleted));
        // Left-anchored interleave: deletion of "beta" trails "alpha"'s entry, before "gamma".
        Assert.Equal(IrAlignmentKind.Unchanged, a.Entries[0].Kind);
        Assert.Equal(IrAlignmentKind.Deleted, a.Entries[1].Kind);
        Assert.Equal(IrAlignmentKind.Unchanged, a.Entries[2].Kind);
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Delete_at_end()
    {
        var l = Doc("alpha", "beta", "gamma");
        var r = Doc("alpha", "beta");
        var a = Align(l, r);

        Assert.Equal(2, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(1, Count(a, IrAlignmentKind.Deleted));
        Assert.Equal(IrAlignmentKind.Deleted, a.Entries[^1].Kind);
        AssertInvariants(l, r, a);
    }

    // ------------------------------------------------------------------ move (headline)

    [Fact]
    public void Pure_move_yields_exactly_one_moved_rest_unchanged()
    {
        // "gamma" relocated from the end to the front; everything else holds in order.
        var l = Doc("alpha", "beta", "gamma", "delta");
        var r = Doc("gamma", "alpha", "beta", "delta");
        var a = Align(l, r);

        Assert.Equal(1, Count(a, IrAlignmentKind.Moved));
        Assert.Equal(3, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(0, Count(a, IrAlignmentKind.Modified));
        Assert.Equal(0, Count(a, IrAlignmentKind.Inserted));
        Assert.Equal(0, Count(a, IrAlignmentKind.Deleted));

        var moved = a.Entries.Single(e => e.Kind == IrAlignmentKind.Moved);
        Assert.Equal("gamma", Text(moved.Right!));
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Move_and_unrelated_edit_classified_independently()
    {
        // "epsilon" relocates from the tail to the front (Moved); "beta" → edited text in place. The
        // edit stays inside a stable spine gap (between alpha and gamma) so it surfaces as Modified
        // independently of the move. (In M2.1 this relied on blind positional pairing; in M2.2 the edit
        // is the lone 1×1-gap residue, paired as Modified by the unambiguous-residue fallback. When an
        // edited paragraph instead RELOCATES into a different gap, M2.2's cross-gap fuzzy pass recovers it
        // as MovedModified — see Cross_gap_move_and_edit_is_moved_modified — rather than the M2.1
        // Delete+Insert.)
        var l = Doc("alpha", "beta", "gamma", "delta", "epsilon");
        var r = Doc("epsilon", "alpha", "beta-edited", "gamma", "delta");
        var a = Align(l, r);

        Assert.Equal(1, Count(a, IrAlignmentKind.Moved));
        Assert.Equal(1, Count(a, IrAlignmentKind.Modified));
        Assert.Equal(3, Count(a, IrAlignmentKind.Unchanged));

        var moved = a.Entries.Single(e => e.Kind == IrAlignmentKind.Moved);
        Assert.Equal("epsilon", Text(moved.Right!));
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Adjacent_swap_of_two_unique_paragraphs()
    {
        // Swap two adjacent unique paragraphs. LIS over the anchor pairs {(0→1),(1→0),(2→2)} has
        // length 2 (e.g. b@1→b'@1, c@2→c'@2 — wait: indices). The longest increasing subsequence by
        // right index keeps the chain that stays in order and drops the one that crosses it, so
        // exactly ONE of the swapped pair is Moved and the other stays Unchanged (plus the unmoved
        // tail). Pinned: 1 Moved + 2 Unchanged.
        var l = Doc("alpha", "beta", "gamma");
        var r = Doc("beta", "alpha", "gamma");
        var a = Align(l, r);

        Assert.Equal(1, Count(a, IrAlignmentKind.Moved));
        Assert.Equal(2, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(0, Count(a, IrAlignmentKind.Modified));
        AssertInvariants(l, r, a);
    }

    // ------------------------------------------------------------------ format-only

    [Fact]
    public void Bolding_a_paragraph_is_format_only()
    {
        var l = FromXml(
            "<w:p><w:r><w:t>alpha</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>beta</w:t></w:r></w:p>");
        var r = FromXml(
            "<w:p><w:r><w:t>alpha</w:t></w:r></w:p>" +
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>beta</w:t></w:r></w:p>");
        var a = Align(l, r);

        Assert.Equal(1, Count(a, IrAlignmentKind.FormatOnly));
        Assert.Equal(1, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(0, Count(a, IrAlignmentKind.Modified));
        Assert.Equal(0, Count(a, IrAlignmentKind.Moved));
        AssertInvariants(l, r, a);
    }

    // ------------------------------------------------------------------ boilerplate

    [Fact]
    public void Boilerplate_delete_one_of_ten_identical_no_false_moves()
    {
        var ten = Enumerable.Repeat("boilerplate", 10).ToArray();
        var nine = Enumerable.Repeat("boilerplate", 9).ToArray();
        var l = Doc(ten);
        var r = Doc(nine);
        var a = Align(l, r);

        Assert.Equal(9, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(1, Count(a, IrAlignmentKind.Deleted));
        Assert.Equal(0, Count(a, IrAlignmentKind.Moved));
        Assert.Equal(0, Count(a, IrAlignmentKind.Modified));
        AssertInvariants(l, r, a);
    }

    // ------------------------------------------------------------------ M2.2 Task 3: similarity pairing + fuzzy moves

    [Fact]
    public void Cross_gap_move_and_edit_is_moved_modified()
    {
        // A multi-word paragraph relocates from the TAIL to the FRONT and is edited in the same revision.
        // M2.1's exact-hash anchoring cannot recognize this (the content hash changed, so no off-spine
        // anchor), and the source/destination land in DIFFERENT spine gaps — so M2.1 produced Delete +
        // Insert. M2.2's cross-gap fuzzy pass re-pairs them: ≥3 words on both sides, similarity ≥ 0.8
        // (only one word changed of seven) → MovedModified.
        var l = Doc(
            "alpha", "beta", "gamma", "delta",
            "the quick brown fox jumps over hounds");
        var r = Doc(
            "the quick brown fox jumps over dogs",
            "alpha", "beta", "gamma", "delta");
        var a = Align(l, r);

        Assert.Equal(1, Count(a, IrAlignmentKind.MovedModified));
        Assert.Equal(4, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(0, Count(a, IrAlignmentKind.Moved));
        Assert.Equal(0, Count(a, IrAlignmentKind.Deleted));
        Assert.Equal(0, Count(a, IrAlignmentKind.Inserted));

        var mm = a.Entries.Single(e => e.Kind == IrAlignmentKind.MovedModified);
        Assert.Equal("the quick brown fox jumps over hounds", Text(mm.Left!));
        Assert.Equal("the quick brown fox jumps over dogs", Text(mm.Right!));
        // The destination entry sits at the moved block's RIGHT position (the front).
        Assert.Equal(IrAlignmentKind.MovedModified, a.Entries[0].Kind);
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Cross_gap_below_similarity_threshold_stays_delete_insert()
    {
        // Same shape as the MovedModified case, but the tail paragraph is REWRITTEN (shares too few words
        // with the front insertion: well under the 0.8 MoveSimilarityThreshold). No fuzzy move — the two
        // stay a clean Delete + Insert rather than a misleading "relocated + edited" claim.
        var l = Doc(
            "alpha", "beta", "gamma", "delta",
            "the quick brown fox jumps over hounds");
        var r = Doc(
            "an entirely unrelated sentence with different words throughout",
            "alpha", "beta", "gamma", "delta");
        var a = Align(l, r);

        Assert.Equal(0, Count(a, IrAlignmentKind.MovedModified));
        Assert.Equal(0, Count(a, IrAlignmentKind.Moved));
        Assert.Equal(1, Count(a, IrAlignmentKind.Deleted));
        Assert.Equal(1, Count(a, IrAlignmentKind.Inserted));
        Assert.Equal(4, Count(a, IrAlignmentKind.Unchanged));
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Cross_gap_below_minimum_token_count_stays_delete_insert()
    {
        // A highly-similar relocation, but BOTH sides have only two Word tokens — under the default
        // MoveMinimumTokenCount of 3. Short fragments are excluded from move detection (they coincidentally
        // match too many candidates), so this stays Delete + Insert despite the high similarity.
        var l = Doc("alpha", "beta", "gamma", "delta", "hello world");
        var r = Doc("hello earth", "alpha", "beta", "gamma", "delta");
        var a = Align(l, r);

        Assert.Equal(0, Count(a, IrAlignmentKind.MovedModified));
        Assert.Equal(0, Count(a, IrAlignmentKind.Moved));
        Assert.Equal(1, Count(a, IrAlignmentKind.Deleted));
        Assert.Equal(1, Count(a, IrAlignmentKind.Inserted));
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Cross_gap_exact_relocation_residue_classifies_as_moved_not_moved_modified()
    {
        // Defensive case for the exact-equal guard in DetectCrossGapMoves. Off-spine anchoring normally
        // catches an exact relocation as plain Moved before the cross-gap pass runs, so a score-1.0 +
        // equal-ContentHash residue is not expected to reach the cross-gap pass — but IF it does, it must
        // classify as Moved (no edit to re-diff), never MovedModified. We force a residue by making the
        // relocated content NON-UNIQUE on one side: "shared phrase here now" appears twice on the left
        // (so it is not a unique anchor and is NOT consumed by anchoring) and once on the right.
        var l = Doc(
            "shared phrase here now", "alpha", "beta", "gamma",
            "shared phrase here now");
        var r = Doc(
            "shared phrase here now", "alpha", "beta", "gamma");
        var a = Align(l, r);

        // One copy stays Unchanged (anchored); the surplus left copy is deleted. No false MovedModified.
        Assert.Equal(0, Count(a, IrAlignmentKind.MovedModified));
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void In_gap_distant_weak_pair_is_rejected_by_locality()
    {
        // A weakly-similar pair (≈0.45 Jaccard, above the base 0.35 floor) at OPPOSITE ends of a large
        // gap must NOT pair: Word's compare anchors insertions next to their matched neighbors and
        // deletes distant old content wholesale — pairing across the gap produces interleaved
        // "word salad" inside an unrelated paragraph. Eligibility is sim ≥ threshold + λ·displacement,
        // so the same texts DO pair when positionally adjacent (see the companion test below).
        var l = Doc(
            "alpha",
            "unrelated opening chatter entirely",
            "second block of miscellaneous filler",
            "third stretch of leftover writing",
            "fourth patch of assorted content",
            "fifth wall of other material",
            "kappa lam sig tau aa bb cc dd ee",
            "omega");
        var r = Doc(
            "alpha",
            "kappa lam sig tau vv ww xx yy zz",
            "omega");
        var a = Align(l, r);

        // The weak far pair is refused: the lone right paragraph is a pure insert, all six left
        // gap paragraphs are deletes.
        Assert.True(Count(a, IrAlignmentKind.Modified) == 0,
            "unexpected pairing: " + string.Join("; ", a.Entries.Select(e =>
                $"{e.Kind}[{(e.Left is null ? "-" : Text(e.Left))}↔{(e.Right is null ? "-" : Text(e.Right))}]")));
        Assert.Equal(1, Count(a, IrAlignmentKind.Inserted));
        Assert.Equal(6, Count(a, IrAlignmentKind.Deleted));
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void In_gap_adjacent_weak_pair_still_pairs()
    {
        // Same ≈0.45 similarity, but positionally aligned at the head of the gap (displacement ≈ 0):
        // the pair forms — locality only penalizes DISTANT weak pairs.
        var l = Doc(
            "alpha",
            "kappa lam sig tau aa bb cc dd ee",
            "unrelated closing chatter entirely",
            "omega");
        var r = Doc(
            "alpha",
            "kappa lam sig tau vv ww xx yy zz",
            "final different words altogether now",
            "omega");
        var a = Align(l, r);

        Assert.Contains(a.Entries, e =>
            e.Kind == IrAlignmentKind.Modified &&
            e.Left is not null && Text(e.Left).Contains("kappa lam sig") &&
            e.Right is not null && Text(e.Right).Contains("kappa lam sig"));
        AssertInvariants(l, r, a);
    }

    // --------------------------------------------------- junction pairing (Word-matcher parity)

    [Fact]
    public void Junction_pairs_titles_on_one_shared_content_word()
    {
        // Word-oracle data point: Word Compare pairs the replace-gap titles "Subtitle Style Demo" ↔
        // "Superscript Demo" into ONE mixed ins+del paragraph on a single shared word ("Demo",
        // word-Jaccard 0.25 — far below the 0.35 similarity threshold). The junction LCS reproduces
        // it; the bodies pair via the ordinary similarity pass.
        var l = Doc("Subtitle Style Demo", "This document demonstrates the Subtitle paragraph style.");
        var r = Doc("Superscript Demo", "This document demonstrates superscript formatting");
        var a = Align(l, r);

        Assert.Equal(2, Count(a, IrAlignmentKind.Modified));
        Assert.Equal(0, Count(a, IrAlignmentKind.Deleted));
        Assert.Equal(0, Count(a, IrAlignmentKind.Inserted));
        Assert.Contains(a.Entries, e => e.Kind == IrAlignmentKind.Modified &&
            Text(e.Left!) == "Subtitle Style Demo" && Text(e.Right!) == "Superscript Demo");
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Junction_pairs_labeled_date_and_year_suffixed_heading()
    {
        // Word Compare's product-roadmap ↔ project-plan redline joins the old heading to the
        // inserted date on their shared 2026. This is deliberately narrower than generic year
        // overlap: a labeled Date: Month day, year may bridge to a short year-suffixed heading.
        // The 8×7 replace gap keeps the Date paragraph at the same relative location as the corpus
        // regression.
        var l = Doc(
            "Product Roadmap 2026",
            "ablation1", "ablation2", "ablation3", "ablation4",
            "ablation5", "ablation6", "ablation7");
        var r = Doc(
            "Project Plan",
            "Date: February 1, 2026",
            "quartz1", "quartz2", "quartz3", "quartz4", "quartz5");
        var a = Align(l, r);

        Assert.Contains(a.Entries, e => e.Kind == IrAlignmentKind.Modified &&
            e.Left is not null && e.Right is not null &&
            Text(e.Left) == "Product Roadmap 2026" &&
            Text(e.Right) == "Date: February 1, 2026");
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Junction_does_not_pair_unlabeled_year_suffixed_titles()
    {
        // A shared calendar year by itself is not evidence: generic title-like paragraphs must
        // remain separate rather than becoming a fabricated Modified diagonal.
        var l = Doc(
            "Budget Summary 2026",
            "ablation1", "ablation2", "ablation3", "ablation4",
            "ablation5", "ablation6", "ablation7");
        var r = Doc(
            "Strategic Plan 2026",
            "quartz1", "quartz2", "quartz3", "quartz4", "quartz5", "quartz6");
        var a = Align(l, r);

        Assert.Equal(0, Count(a, IrAlignmentKind.Modified));
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Junction_labeled_date_bridge_does_not_grow_on_year_only_neighbor()
    {
        // The date bridge is LCS-only. Its adjacent 2026-bearing paragraphs do not gain generic
        // numeric evidence through diagonal growth.
        var l = Doc(
            "Product Roadmap 2026",
            "Adjacent Context 2026",
            "ablation1", "ablation2", "ablation3", "ablation4", "ablation5", "ablation6");
        var r = Doc(
            "Project Plan",
            "Date: February 1, 2026",
            "Different Neighbor 2026",
            "quartz1", "quartz2", "quartz3", "quartz4", "quartz5");
        var a = Align(l, r);

        Assert.Contains(a.Entries, e => e.Kind == IrAlignmentKind.Modified &&
            e.Left is not null && e.Right is not null &&
            Text(e.Left) == "Product Roadmap 2026" &&
            Text(e.Right) == "Date: February 1, 2026");
        Assert.DoesNotContain(a.Entries, e => e.Kind == IrAlignmentKind.Modified &&
            e.Left is not null && e.Right is not null &&
            Text(e.Left) == "Adjacent Context 2026" &&
            Text(e.Right) == "Different Neighbor 2026");
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Junction_keeps_short_shared_ordinals_out_of_weak_pairing()
    {
        // The numeric guard added in 19e0 exists to stop list scaffolding from manufacturing a
        // Modified diagonal. The calendar-year exception must not re-admit a short ordinal.
        var l = Doc(
            "Legacy Ledger 17",
            "ablation1", "ablation2", "ablation3", "ablation4",
            "ablation5", "ablation6", "ablation7");
        var r = Doc(
            "Project Plan",
            "Briefing Note 17",
            "quartz1", "quartz2", "quartz3", "quartz4", "quartz5");
        var a = Align(l, r);

        Assert.Equal(0, Count(a, IrAlignmentKind.Modified));
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Junction_declines_stopword_grade_overlap()
    {
        // Word-oracle data point (header_no_rels ↔ heading_1_bold): despite sharing "with"/"the",
        // Word keeps EVERY paragraph separate — stopword-grade overlap (word-Jaccard 0.091 for the
        // closest pair, below the calibrated 0.10 floor) is no pairing evidence. All left paragraphs
        // delete, all right paragraphs insert; nothing force-pairs (no 1×1 residue in a 3×3 gap).
        var l = Doc(
            "header-no-rels",
            "Second page",
            "Some content in the second section.... with just an empty p, this section isn't rendered?");
        var r = Doc(
            "Heading 1 Bold Demo",
            "This document shows Heading 1 style with extra bold emphasis.",
            "Heading 1 with bold creates the strongest document headers.");
        var a = Align(l, r);

        Assert.Equal(0, Count(a, IrAlignmentKind.Modified));
        Assert.Equal(3, Count(a, IrAlignmentKind.Deleted));
        Assert.Equal(3, Count(a, IrAlignmentKind.Inserted));
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Junction_growth_extends_below_a_paired_title()
    {
        // Word-oracle data point (heading_3 ↔ heading_4_right_italic): the demo titles pair via the
        // similarity pass; the bodies below them share only "Heading" (word-Jaccard 0.077 — below
        // any defensible global floor) yet Word still merges them into one mixed paragraph. The
        // diagonal growth phase reproduces it: a free pair sitting right under an established pair
        // needs only ≥1 shared word. The third paragraphs share NOTHING, so they stay separate
        // (the 1×1 residue does not force a zero-shared paragraph pair).
        var l = Doc(
            "Heading 3 Style Demo",
            "Demonstrating Heading 3 paragraph style.",
            "Section Sub-header");
        var r = Doc(
            "Heading 4 Right Italic Demo",
            "Heading 4 style with right alignment and italic formatting.",
            "This creates unique stylized subheadings for document sections.");
        var a = Align(l, r);

        Assert.Equal(2, Count(a, IrAlignmentKind.Modified));
        Assert.Equal(1, Count(a, IrAlignmentKind.Deleted));
        Assert.Equal(1, Count(a, IrAlignmentKind.Inserted));
        Assert.Contains(a.Entries, e => e.Kind == IrAlignmentKind.Modified &&
            Text(e.Left!).StartsWith("Demonstrating") && Text(e.Right!).StartsWith("Heading 4 style"));
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Junction_declines_function_word_only_overlap_above_the_floor()
    {
        // Word-oracle data point (potpourri ↔ product_roadmap): '2.2 Numbered (with nested)' does
        // NOT merge into 'Q1: Launch v2.0 with new dashboard' although their word-Jaccard (0.11 on
        // the shared "with") clears the 0.10 floor — a shared closed-class FUNCTION word is no
        // pairing evidence. Modeled minimally: same shared-"with" shape, Jaccard above the floor.
        // (A second unrelated left paragraph keeps this out of the laxer 1×1-residue path — the
        // corpus shape was a 76×8 replace gap.)
        var l = Doc("alpha", "Numbered (with nested)", "Second unrelated stanza entirely", "omega");
        var r = Doc("alpha", "Launch dashboards with telemetry", "omega");
        var a = Align(l, r);

        Assert.Equal(0, Count(a, IrAlignmentKind.Modified));
        Assert.Equal(2, Count(a, IrAlignmentKind.Deleted));
        Assert.Equal(1, Count(a, IrAlignmentKind.Inserted));
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Junction_pairs_on_containment_even_for_function_words()
    {
        // Word-oracle data point (word_tolerated_duplicate_ppr ↔ word_tolerated_misplaced_link):
        // Word merges the one-word paragraph "a" into "A) ST_OnOff values for <w:b> on a run:" —
        // when the shared words cover at least HALF of the smaller side, the paragraph is mostly
        // CONTAINED in its counterpart (an extension, not a replacement), and even function-word
        // overlap pairs. NB: the 3×2 gap here also proves this is the junction LCS, not the 1×1
        // residue.
        var l = Doc("alpha", "a", "x", "omega");
        var r = Doc("alpha", "sample on a run", "omega");
        var a = Align(l, r);

        Assert.Contains(a.Entries, e => e.Kind == IrAlignmentKind.Modified &&
            Text(e.Left!) == "a" && Text(e.Right!) == "sample on a run");
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Junction_growth_respects_size_parity()
    {
        // Word-oracle data point (justify_alignment ↔ large_font_size): the titles pair on "Demo",
        // but the 30-word justified body does NOT merge into the 7-word "This document demonstrates
        // large 24pt font size." on the boilerplate "This document demonstrates" — Word deletes it
        // wholesale. The growth size-parity guard (min ≥ ⅓·max word count) encodes that.
        var l = Doc(
            "Justify Alignment Demo",
            "This document demonstrates justified text alignment which spreads text evenly across " +
            "the full width of the line, creating clean left and right edges that are perfect for " +
            "formal documents and publications.");
        var r = Doc(
            "Large Font Size Demo",
            "This document demonstrates large 24pt font size.",
            "Large fonts are great for titles and presentations.");
        var a = Align(l, r);

        Assert.Equal(1, Count(a, IrAlignmentKind.Modified));
        Assert.Equal(1, Count(a, IrAlignmentKind.Deleted));
        Assert.Equal(2, Count(a, IrAlignmentKind.Inserted));
        var modified = a.Entries.Single(e => e.Kind == IrAlignmentKind.Modified);
        Assert.Equal("Justify Alignment Demo", Text(modified.Left!));
        Assert.Equal("Large Font Size Demo", Text(modified.Right!));
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void In_gap_leftover_tables_pair_positionally()
    {
        // Word merges an old table into the replacing new table (per-cell del+ins interleave) even
        // when the gap holds MORE tables on one side — the k-th leftover table pairs with the k-th,
        // the surplus inserts/deletes. The old rule required exactly one free table on EACH side and
        // lowered everything else to whole-table delete+insert stacks.
        var l = FromXml(
            "<w:p><w:r><w:t>alpha</w:t></w:r></w:p>" +
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid>" +
            "<w:tr><w:tc><w:p><w:r><w:t>old cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
            "<w:p><w:r><w:t>omega</w:t></w:r></w:p>");
        var r = FromXml(
            "<w:p><w:r><w:t>alpha</w:t></w:r></w:p>" +
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid>" +
            "<w:tr><w:tc><w:p><w:r><w:t>brand new cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/></w:tblGrid>" +
            "<w:tr><w:tc><w:p><w:r><w:t>second fresh table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
            "<w:p><w:r><w:t>omega</w:t></w:r></w:p>");
        var a = Align(l, r);

        // The left table pairs with the FIRST right table as Modified; the second right table inserts.
        Assert.Equal(1, Count(a, IrAlignmentKind.Modified));
        Assert.Equal(1, Count(a, IrAlignmentKind.Inserted));
        Assert.Equal(0, Count(a, IrAlignmentKind.Deleted));
        var modified = a.Entries.Single(e => e.Kind == IrAlignmentKind.Modified);
        Assert.True(modified.Left is Docxodus.Ir.IrTable && modified.Right is Docxodus.Ir.IrTable);
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void In_gap_cross_positioned_edit_pairs_as_modified()
    {
        // Two paragraphs are edited AND swapped WITHIN a single spine gap (between alpha and omega). M2.1's
        // blind positional pairing would have paired them by position — pairing edited-P1 with edited-P2's
        // slot and vice-versa, producing two low-quality Modified pairs. M2.2's in-gap similarity pairing
        // matches each edited paragraph to its true counterpart by score, so both surface as faithful
        // Modified pairs (each ≥ 0.5 similarity to its real original, far above its similarity to the
        // other). This is the upgrade of the M2.1 gap-positional limitation.
        var l = Doc(
            "alpha",
            "the quick brown fox jumps high",
            "a lazy sleepy dog rests here",
            "omega");
        var r = Doc(
            "alpha",
            "a lazy sleepy dog rests there",      // edit of P2 (one word), now in P1's slot
            "the quick brown fox leaps high",     // edit of P1 (one word), now in P2's slot
            "omega");
        var a = Align(l, r);

        Assert.Equal(2, Count(a, IrAlignmentKind.Modified));
        Assert.Equal(2, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(0, Count(a, IrAlignmentKind.Deleted));
        Assert.Equal(0, Count(a, IrAlignmentKind.Inserted));

        // Each Modified pair joins an edited paragraph to its TRUE original (by content), not its slot-mate.
        var modifies = a.Entries.Where(e => e.Kind == IrAlignmentKind.Modified).ToList();
        Assert.Contains(modifies, e =>
            Text(e.Left!).Contains("quick brown fox") && Text(e.Right!).Contains("quick brown fox"));
        Assert.Contains(modifies, e =>
            Text(e.Left!).Contains("lazy sleepy dog") && Text(e.Right!).Contains("lazy sleepy dog"));
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Boilerplate_adversarial_yields_zero_false_moves()
    {
        // A boilerplate-heavy edit: 8 identical clauses, one deleted, plus a short distinct edit. The
        // similarity + cross-gap passes must NOT manufacture moves out of the repeated boilerplate.
        var l = Doc(
            "Standard clause.", "Standard clause.", "Standard clause.", "Standard clause.",
            "Standard clause.", "Standard clause.", "Standard clause.", "Standard clause.",
            "unique closing remark goes here");
        var r = Doc(
            "Standard clause.", "Standard clause.", "Standard clause.", "Standard clause.",
            "Standard clause.", "Standard clause.", "Standard clause.",
            "unique closing remark goes here");
        var a = Align(l, r);

        Assert.Equal(0, Count(a, IrAlignmentKind.Moved));
        Assert.Equal(0, Count(a, IrAlignmentKind.MovedModified));
        Assert.Equal(1, Count(a, IrAlignmentKind.Deleted));  // the one removed boilerplate copy
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Cross_gap_move_detection_is_deterministic()
    {
        var l = Doc(
            "alpha", "beta", "gamma",
            "the quick brown fox jumps over hounds");
        var r = Doc(
            "the quick brown fox jumps over dogs",
            "alpha", "beta", "gamma");

        var a1 = Align(l, r);
        var a2 = Align(l, r);
        Assert.True(a1.Entries.SequenceEqual(a2.Entries),
            "Cross-gap fuzzy move detection must be deterministic across Align calls.");
        Assert.Equal(1, Count(a1, IrAlignmentKind.MovedModified));
        AssertInvariants(l, r, a1);
    }

    // ------------------------------------------------------------------ table as unit

    [Fact]
    public void Table_cell_edit_makes_table_block_modified()
    {
        const string tbl =
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"100\"/></w:tblGrid>" +
            "<w:tr><w:tc><w:p><w:r><w:t>{0}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        var l = FromXml("<w:p><w:r><w:t>intro</w:t></w:r></w:p>" + string.Format(tbl, "cell-old"));
        var r = FromXml("<w:p><w:r><w:t>intro</w:t></w:r></w:p>" + string.Format(tbl, "cell-new"));
        var a = Align(l, r);

        Assert.Equal(1, Count(a, IrAlignmentKind.Unchanged)); // the intro paragraph
        Assert.Equal(1, Count(a, IrAlignmentKind.Modified));  // the table as ONE unit
        var modified = a.Entries.Single(e => e.Kind == IrAlignmentKind.Modified);
        Assert.IsType<IrTable>(modified.Left);
        Assert.IsType<IrTable>(modified.Right);
        AssertInvariants(l, r, a);
    }

    /// <summary>
    /// M2.4b Workstream C grain LOCK (added in WS-D review follow-up). The unambiguous-table-residue rule
    /// (<see cref="IrBlockAligner"/>.FillOneGap) pairs the lone free-left table with the lone free-right table
    /// as Modified REGARDLESS of similarity — a table can only sensibly pair with a table, so a heavily-edited
    /// (here COMPLETELY UNRELATED) table is still ONE edited table, not a delete+insert of two tables. This
    /// test pins that choice for the extreme case: two tables that share NO cell content, isolated in a gap
    /// between two unchanged paragraphs. They MUST pair as ONE Modified table (not Deleted+Inserted), and the
    /// rendered grain MUST be clean all-rows delete+insert — every left cell's text Deleted and every right
    /// cell's text Inserted, with NO coincidental Equal island splitting the rows (the rows pair positionally
    /// into ModifyRows whose totally-different cells token-diff to whole del+ins). Locks both the pairing
    /// decision and the resulting revision grain against regression.
    /// </summary>
    [Fact]
    public void Unrelated_tables_in_a_gap_pair_as_modified_with_all_rows_del_ins_grain()
    {
        const string row = "<w:tr><w:tc><w:p><w:r><w:t>{0}</w:t></w:r></w:p></w:tc></w:tr>";
        string Table(string a, string b) =>
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"100\"/></w:tblGrid>" +
            string.Format(row, a) + string.Format(row, b) + "</w:tbl>";

        // Stable spine paragraphs bracket the table so it is an ISOLATED gap residue (1 free table each side).
        var l = FromXml("<w:p><w:r><w:t>head</w:t></w:r></w:p>" + Table("Apple", "Banana") +
                        "<w:p><w:r><w:t>tail</w:t></w:r></w:p>");
        var r = FromXml("<w:p><w:r><w:t>head</w:t></w:r></w:p>" + Table("Xylophone", "Zebra") +
                        "<w:p><w:r><w:t>tail</w:t></w:r></w:p>");

        var a = Align(l, r);
        Assert.Equal(2, Count(a, IrAlignmentKind.Unchanged));  // head + tail
        Assert.Equal(1, Count(a, IrAlignmentKind.Modified));   // the table as ONE unit (residue rule)
        Assert.Equal(0, Count(a, IrAlignmentKind.Deleted));    // NOT a whole-table delete+insert
        Assert.Equal(0, Count(a, IrAlignmentKind.Inserted));
        var modified = a.Entries.Single(e => e.Kind == IrAlignmentKind.Modified);
        Assert.IsType<IrTable>(modified.Left);
        Assert.IsType<IrTable>(modified.Right);
        AssertInvariants(l, r, a);

        // Rendered grain: every left cell text Deleted, every right cell text Inserted (no shared Equal island
        // because the cells share nothing). Compatible mode is what the GetRevisions surface uses.
        var script = IrEditScriptBuilder.Build(l, r, new IrDiffSettings { RevisionGranularity = RevisionGranularity.WmlComparerCompatible });
        var revs = IrRevisionRenderer.Render(script, l, r, new IrDiffSettings { RevisionGranularity = RevisionGranularity.WmlComparerCompatible });
        var deleted = string.Concat(revs.Where(x => x.Type == IrRevisionType.Deleted).Select(x => x.Text));
        var inserted = string.Concat(revs.Where(x => x.Type == IrRevisionType.Inserted).Select(x => x.Text));
        Assert.Contains("Apple", deleted);
        Assert.Contains("Banana", deleted);
        Assert.Contains("Xylophone", inserted);
        Assert.Contains("Zebra", inserted);
        // No FormatChanged/Moved noise and nothing left Equal — the change is wholly del+ins.
        Assert.DoesNotContain(revs, x => x.Type is IrRevisionType.Moved or IrRevisionType.FormatChanged);
    }

    // ------------------------------------------------------------------ empty docs

    [Fact]
    public void Empty_left_all_inserted()
    {
        var l = FromXml(string.Empty);
        var r = Doc("alpha", "beta");
        var a = Align(l, r);

        Assert.Equal(2, Count(a, IrAlignmentKind.Inserted));
        Assert.Equal(2, a.Entries.Count);
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Empty_right_all_deleted()
    {
        var l = Doc("alpha", "beta");
        var r = FromXml(string.Empty);
        var a = Align(l, r);

        Assert.Equal(2, Count(a, IrAlignmentKind.Deleted));
        Assert.Equal(2, a.Entries.Count);
        AssertInvariants(l, r, a);
    }

    [Fact]
    public void Both_empty_no_entries()
    {
        var l = FromXml(string.Empty);
        var r = FromXml(string.Empty);
        var a = Align(l, r);

        Assert.Empty(a.Entries);
        AssertInvariants(l, r, a);
    }

    // ------------------------------------------------------------------ determinism

    [Fact]
    public void Two_align_calls_are_sequence_equal()
    {
        var l = Doc("alpha", "beta", "gamma", "delta", "boilerplate", "boilerplate");
        var r = Doc("gamma", "alpha", "beta-edited", "boilerplate", "delta", "NEW");

        var a1 = Align(l, r);
        var a2 = Align(l, r);

        Assert.True(a1.Entries.SequenceEqual(a2.Entries),
            "Two Align calls on identical inputs must produce sequence-equal entries.");
        AssertInvariants(l, r, a1);
    }

    private static string Text(IrBlock b) =>
        b is IrParagraph p
            ? string.Concat(p.Inlines.OfType<IrTextRun>().Select(t => t.Text))
            : string.Empty;
}
