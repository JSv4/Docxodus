#nullable enable

using System.Collections.Generic;
using System.Linq;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.4 Task 1.3 tests for <see cref="IrCompositeMerger.ComposeTokenSpans"/>: the EXACT sub-block
/// merge that composes N reviewers' per-base-paragraph token diffs (all over the SAME base token
/// stream) into one authored token-op list plus the conflict list.
/// </summary>
/// <remarks>
/// <para><b>Real-differ fixtures.</b> Replacement/overlap fixtures are built from genuine
/// <see cref="IrTokenDiffer"/> output (via <see cref="DiffReal"/>: tokenize two real paragraphs read
/// by <see cref="IrReader"/>, then diff). The differ emits a replacement as a Delete FOLLOWED BY an
/// Insert anchored at <c>Delete.LeftEnd</c> (NOT <c>Delete.LeftStart</c>) — e.g. <c>"the quick fox"
/// → "the slow fox"</c> ⇒ <c>Equal(0,2) Delete(2,3|2,2) Insert(3,3|2,3) Equal(3,5)</c>. The merger
/// must pair each Delete with its immediately-following Insert to recover the replacement text; older
/// hand-built fixtures that anchored the Insert at the deleted token's LeftStart validated a bug the
/// differ never produces.</para>
/// <para><b>Tiny hand-built fixtures</b> (pure inserts) keep the differ convention exactly: an insert
/// is <c>Insert(p,p|rStart,rEnd)</c>; a pure delete is <c>Delete(a,b|x,x)</c>; a replacement is
/// <c>Delete(a,b|x,x) Insert(b,b|..)</c>.</para>
/// </remarks>
public class IrTokenComposeTests
{
    private static readonly IrRunFormat Plain = new() { Bold = false, UnmodeledDigest = default };
    private static readonly IrReaderOptions NoSources = new() { RetainSources = false };
    private static readonly IrDiffSettings DiffDefaults = new();

    // ---- real-differ helpers ------------------------------------------------

    /// <summary>Tokenize the first body paragraph of a one-paragraph DOCX built from <paramref name="text"/>.</summary>
    private static IReadOnlyList<IrDiffToken> Tokens(string text) =>
        IrDiffTokenizer.Tokenize(
            IrReader.Read(IrTestDocuments.Create(text), NoSources).Body.Blocks.OfType<IrParagraph>().First(),
            DiffDefaults);

    /// <summary>
    /// A reviewer tuple from a GENUINE token diff of <paramref name="baseText"/> → <paramref name="reviewerText"/>:
    /// the real <see cref="IrTokenDiff"/> and the real right-token list. The base token count is the
    /// length of <paramref name="baseText"/>'s token stream (shared across reviewers by construction).
    /// </summary>
    private static (int, string, IrTokenDiff, IReadOnlyList<IrDiffToken>) ReviewerReal(
        int idx, string author, string baseText, string reviewerText)
    {
        var rightTokens = Tokens(reviewerText);
        var diff = IrTokenDiffer.Diff(Tokens(baseText), rightTokens, DiffDefaults);
        return (idx, author, diff, rightTokens);
    }

    /// <summary>Base token count for <paramref name="baseText"/> — the shared base token stream length.</summary>
    private static int BaseCount(string baseText) => Tokens(baseText).Count;

    // ---- tiny hand-built helpers (differ-convention-exact) ------------------

    /// <summary>One Word token whose Text/MatchKey are <paramref name="text"/>.</summary>
    private static IrDiffToken Tok(string text) =>
        new(IrDiffTokenKind.Word, text, text, 0, text.Length, Plain);

    /// <summary>A right-token list of single Word tokens, one per supplied string.</summary>
    private static List<IrDiffToken> Rights(params string[] texts) =>
        texts.Select(Tok).ToList();

    /// <summary>An Insert op anchored at base position <paramref name="pos"/> spanning right tokens
    /// <paramref name="rStart"/>..<paramref name="rEnd"/> (differ shape: <c>Insert(p,p|rStart,rEnd)</c>).</summary>
    private static IrTokenOp Ins(int pos, int rStart, int rEnd) =>
        new(IrTokenOpKind.Insert, pos, pos, rStart, rEnd);

    /// <summary>A pure Delete op of the base token at <paramref name="pos"/> (differ shape:
    /// <c>Delete(pos,pos+1|rAnchor,rAnchor)</c>).</summary>
    private static IrTokenOp Del(int pos, int rAnchor) =>
        new(IrTokenOpKind.Delete, pos, pos + 1, rAnchor, rAnchor);

    private static (int, string, IrTokenDiff, IReadOnlyList<IrDiffToken>) Reviewer(
        int idx, string author, IReadOnlyList<IrDiffToken> rights, params IrTokenOp[] ops) =>
        (idx, author, new IrTokenDiff(IrNodeList.From(ops)), rights);

    // ---- pure-insert fixtures (hand-built, differ-shape) --------------------

    [Fact]
    public void Disjoint_token_inserts_compose_without_conflict()
    {
        // base count 5; R1 inserts "AAA" at pos 1; R2 inserts "BBB" at pos 3.
        var reviewers = new List<(int, string, IrTokenDiff, IReadOnlyList<IrDiffToken>)>
        {
            Reviewer(0, "Bob", Rights("AAA"), Ins(1, 0, 1)),
            Reviewer(1, "Fred", Rights("BBB"), Ins(3, 0, 1)),
        };

        var ops = IrCompositeMerger.ComposeTokenSpans(
            5, reviewers, ConflictResolution.BaseWins, "p:body:base", 1, out var conflicts);

        Assert.Empty(conflicts);
        var inserts = ops.Where(o => o.Op.Kind == IrTokenOpKind.Insert).ToList();
        Assert.Equal(2, inserts.Count);
        var authors = inserts.Select(o => o.Author).ToHashSet();
        Assert.Contains("Bob", authors);
        Assert.Contains("Fred", authors);
    }

    [Fact]
    public void Same_text_insert_by_two_reviewers_is_consensus()
    {
        var reviewers = new List<(int, string, IrTokenDiff, IReadOnlyList<IrDiffToken>)>
        {
            Reviewer(0, "Bob", Rights("SAME"), Ins(2, 0, 1)),
            Reviewer(1, "Fred", Rights("SAME"), Ins(2, 0, 1)),
        };

        var ops = IrCompositeMerger.ComposeTokenSpans(
            5, reviewers, ConflictResolution.BaseWins, "p:body:base", 1, out var conflicts);

        Assert.Empty(conflicts);
        var inserts = ops.Where(o => o.Op.Kind == IrTokenOpKind.Insert).ToList();
        Assert.Single(inserts);
        Assert.Equal("Bob", inserts[0].Author); // first reviewer in the group is the author
    }

    [Fact]
    public void Different_text_insert_at_same_anchor_conflicts()
    {
        var reviewers = new List<(int, string, IrTokenDiff, IReadOnlyList<IrDiffToken>)>
        {
            Reviewer(0, "Bob", Rights("X"), Ins(2, 0, 1)),
            Reviewer(1, "Fred", Rights("Y"), Ins(2, 0, 1)),
        };

        var ops = IrCompositeMerger.ComposeTokenSpans(
            5, reviewers, ConflictResolution.BaseWins, "p:body:base", 7, out var conflicts);

        Assert.Single(conflicts);
        Assert.Equal(7, conflicts[0].Id);
        Assert.Equal(2, conflicts[0].TokenStart);
        Assert.Equal(2, conflicts[0].TokenEnd);
        Assert.Equal(2, conflicts[0].Competitors.Count);
        // BaseWins → no insert at pos 2.
        Assert.Empty(ops.Where(o => o.Op.Kind == IrTokenOpKind.Insert));
    }

    // ---- real-differ replacement / overlap fixtures -------------------------

    [Fact]
    public void Real_differ_emits_delete_then_insert_at_left_end_for_replacement()
    {
        // Sanity-pin the convention the merger depends on: replacement = Delete then Insert anchored
        // at Delete.LeftEnd (one token past the deleted token), NOT at the deleted token's LeftStart.
        var (_, _, diff, _) = ReviewerReal(0, "Bob", "the quick fox", "the slow fox");
        var sig = string.Join(" ", diff.Ops.Select(
            o => $"{o.Kind}({o.LeftStart},{o.LeftEnd}|{o.RightStart},{o.RightEnd})"));
        Assert.Equal("Equal(0,2|0,2) Delete(2,3|2,2) Insert(3,3|2,3) Equal(3,5|3,5)", sig);
    }

    [Fact]
    public void Overlapping_replacement_conflicts_basewins_keeps_base()
    {
        // Real differ output: both reviewers replace base "quick" (token 2) with DIFFERENT words.
        // Each diff is Equal(0,2) Delete(2,3) Insert(3,3|..) Equal(3,5). The single deleted base token
        // is token 2; the replacement Insert lives at pos 3 (== Delete.LeftEnd) and is owned by the
        // delete path — it must NOT also raise an insert conflict at pos 3 (no double-count).
        const string baseText = "the quick fox";
        var reviewers = new List<(int, string, IrTokenDiff, IReadOnlyList<IrDiffToken>)>
        {
            ReviewerReal(0, "Bob", baseText, "the slow fox"),
            ReviewerReal(1, "Fred", baseText, "the fast fox"),
        };

        var ops = IrCompositeMerger.ComposeTokenSpans(
            BaseCount(baseText), reviewers, ConflictResolution.BaseWins, "p:body:base", 1, out var conflicts);

        // I2 (tightened): EXACTLY one conflict — the delete-with-differing-replacement at base token 2.
        // (The replacement Inserts at pos 3 do NOT separately conflict.)
        Assert.Single(conflicts);
        Assert.Equal(2, conflicts[0].TokenStart);
        Assert.Equal(3, conflicts[0].TokenEnd);
        Assert.Equal(2, conflicts[0].Competitors.Count);
        Assert.Equal(new[] { "slow", "fast" }, conflicts[0].Competitors.Select(c => c.ResultText).ToArray());

        // BaseWins → base token 2 survives as Equal; no Delete at LeftStart 2; no insert emitted.
        Assert.DoesNotContain(ops, o => o.Op.Kind == IrTokenOpKind.Delete && o.Op.LeftStart == 2);
        Assert.Contains(ops, o => o.Op.Kind == IrTokenOpKind.Equal && o.Op.LeftStart == 2);
        Assert.Empty(ops.Where(o => o.Op.Kind == IrTokenOpKind.Insert));
    }

    // ---- delete-conflict policy coverage (I1) -------------------------------

    private static List<(int, string, IrTokenDiff, IReadOnlyList<IrDiffToken>)> ReplaceQuickConflict()
    {
        const string baseText = "the quick fox";
        return new()
        {
            ReviewerReal(0, "Bob", baseText, "the slow fox"),
            ReviewerReal(1, "Fred", baseText, "the fast fox"),
        };
    }

    [Fact]
    public void Delete_with_differing_replacement_BaseWins_keeps_base()
    {
        var ops = IrCompositeMerger.ComposeTokenSpans(
            BaseCount("the quick fox"), ReplaceQuickConflict(),
            ConflictResolution.BaseWins, "p:body:base", 1, out var conflicts);

        Assert.Single(conflicts);
        // Base survives at token 2 as Equal; nothing inserted.
        Assert.Contains(ops, o => o.Op.Kind == IrTokenOpKind.Equal && o.Op.LeftStart == 2);
        Assert.DoesNotContain(ops, o => o.Op.Kind == IrTokenOpKind.Delete && o.Op.LeftStart == 2);
        Assert.Empty(ops.Where(o => o.Op.Kind == IrTokenOpKind.Insert));
    }

    [Fact]
    public void Delete_with_differing_replacement_FirstReviewerWins_applies_first()
    {
        var reviewers = ReplaceQuickConflict();
        var ops = IrCompositeMerger.ComposeTokenSpans(
            BaseCount("the quick fox"), reviewers,
            ConflictResolution.FirstReviewerWins, "p:body:base", 1, out var conflicts);

        Assert.Single(conflicts);
        // First deleter (Bob) wins: his Delete of token 2 plus his replacement insert "slow".
        Assert.Contains(ops, o => o.Op.Kind == IrTokenOpKind.Delete && o.Op.LeftStart == 2 && o.Author == "Bob");
        var inserts = ops.Where(o => o.Op.Kind == IrTokenOpKind.Insert).ToList();
        Assert.Single(inserts);
        Assert.Equal("Bob", inserts[0].Author);
        Assert.Equal("slow", RightOf(inserts[0], reviewers[0].Item4));
    }

    [Fact]
    public void Delete_with_differing_replacement_StackAll_one_delete_stacked_inserts()
    {
        var reviewers = ReplaceQuickConflict();
        var ops = IrCompositeMerger.ComposeTokenSpans(
            BaseCount("the quick fox"), reviewers,
            ConflictResolution.StackAll, "p:body:base", 1, out var conflicts);

        Assert.Single(conflicts);
        // StackAll: a base token can only be deleted once — exactly ONE Delete at token 2 (first
        // deleter), but BOTH replacement inserts are stacked (in reviewer order).
        var deletes = ops.Where(o => o.Op.Kind == IrTokenOpKind.Delete && o.Op.LeftStart == 2).ToList();
        Assert.Single(deletes);
        Assert.Equal("Bob", deletes[0].Author);
        var inserts = ops.Where(o => o.Op.Kind == IrTokenOpKind.Insert).ToList();
        Assert.Equal(2, inserts.Count);
        Assert.Equal(new[] { "Bob", "Fred" }, inserts.Select(o => o.Author).ToArray());
    }

    /// <summary>Concatenate the raw text of an authored Insert op's right span.</summary>
    private static string RightOf(IrAuthoredTokenOp authored, IReadOnlyList<IrDiffToken> rightTokens)
    {
        var op = authored.Op;
        return string.Concat(Enumerable.Range(op.RightStart, op.RightEnd - op.RightStart)
            .Select(i => rightTokens[i].Text));
    }

    // ---- pure delete + empty + policy-knob fixtures -------------------------

    [Fact]
    public void Pure_delete_consensus()
    {
        // Real differ: "the quick fox" → "the fox" deletes base tokens [2,4) ("quick" + its separator),
        // emitting a single multi-token Delete and NO insert (pure delete).
        const string baseText = "the quick fox";
        var reviewers = new List<(int, string, IrTokenDiff, IReadOnlyList<IrDiffToken>)>
        {
            ReviewerReal(0, "Bob", baseText, "the fox"),
            ReviewerReal(1, "Fred", baseText, "the fox"),
        };

        var ops = IrCompositeMerger.ComposeTokenSpans(
            BaseCount(baseText), reviewers, ConflictResolution.BaseWins, "p:body:base", 1, out var conflicts);

        Assert.Empty(conflicts);
        // The deleted base tokens 2 and 3 each resolve to a (consensus) Delete authored to Bob; no
        // replacement insert exists for a pure delete.
        var deletes = ops.Where(o => o.Op.Kind == IrTokenOpKind.Delete).ToList();
        Assert.NotEmpty(deletes);
        Assert.All(deletes, d => Assert.Equal("Bob", d.Author));
        Assert.Empty(ops.Where(o => o.Op.Kind == IrTokenOpKind.Insert));
    }

    [Fact]
    public void Empty_reviewer_diffs_yield_all_equal_no_conflicts()
    {
        // M3: no reviewers → every base token survives as Equal, no conflicts.
        var ops = IrCompositeMerger.ComposeTokenSpans(
            4,
            new List<(int, string, IrTokenDiff, IReadOnlyList<IrDiffToken>)>(),
            ConflictResolution.BaseWins, "p:body:base", 1, out var conflicts);

        Assert.Empty(conflicts);
        Assert.Equal(4, ops.Count);
        Assert.All(ops, o => Assert.Equal(IrTokenOpKind.Equal, o.Op.Kind));
        Assert.Equal(new[] { 0, 1, 2, 3 }, ops.Select(o => o.Op.LeftStart).ToArray());
    }

    [Fact]
    public void FirstReviewerWins_and_StackAll_emit_expected()
    {
        List<(int, string, IrTokenDiff, IReadOnlyList<IrDiffToken>)> Build() => new()
        {
            Reviewer(0, "Bob", Rights("X"), Ins(2, 0, 1)),
            Reviewer(1, "Fred", Rights("Y"), Ins(2, 0, 1)),
        };

        var firstWins = IrCompositeMerger.ComposeTokenSpans(
            5, Build(), ConflictResolution.FirstReviewerWins, "p:body:base", 1, out var c1);
        Assert.Single(c1);
        var fwInserts = firstWins.Where(o => o.Op.Kind == IrTokenOpKind.Insert).ToList();
        Assert.Single(fwInserts);
        Assert.Equal("Bob", fwInserts[0].Author);

        var stackAll = IrCompositeMerger.ComposeTokenSpans(
            5, Build(), ConflictResolution.StackAll, "p:body:base", 1, out var c2);
        Assert.Single(c2);
        var saInserts = stackAll.Where(o => o.Op.Kind == IrTokenOpKind.Insert).ToList();
        Assert.Equal(2, saInserts.Count);
        Assert.Equal(new[] { "Bob", "Fred" }, saInserts.Select(o => o.Author).ToArray());
    }
}
