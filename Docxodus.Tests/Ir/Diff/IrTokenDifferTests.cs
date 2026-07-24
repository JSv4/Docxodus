#nullable enable

using System.Collections.Generic;
using System.Linq;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// M2.2 Task 1 tests for <see cref="IrTokenDiffer"/>: Myers token diff over MatchKeys plus the
/// per-token <see cref="IrRunFormat"/> format post-pass. Every case runs <see cref="IrTokenDiffAsserts"/>
/// over its result so the totality/coverage/per-kind invariants are enforced uniformly.
/// </summary>
/// <remarks>
/// Most cases build synthetic token lists directly (so the test pins exact Myers output without DOCX
/// round-tripping). One case tokenizes two real paragraphs via <see cref="IrTestDocuments"/> +
/// <see cref="IrReader"/> with <c>RetainSources = false</c>, proving the differ works end-to-end on
/// reader-produced tokens with no provenance.
/// </remarks>
public class IrTokenDifferTests
{
    private static readonly IrDiffSettings Default = new();

    // --- synthetic token builders ----------------------------------------

    /// <summary>Build a Word token whose Text/MatchKey are <paramref name="word"/> and whose Format is
    /// <paramref name="format"/> (defaults to a plain non-bold run format).</summary>
    private static IrDiffToken W(string word, IrRunFormat? format = null) =>
        new(IrDiffTokenKind.Word, word, word, 0, word.Length, format ?? Plain);

    /// <summary>A whole-paragraph token list from space-joined words (one Word + one Separator each).</summary>
    private static List<IrDiffToken> Words(params string[] words)
    {
        var tokens = new List<IrDiffToken>();
        for (int i = 0; i < words.Length; i++)
        {
            if (i > 0)
                tokens.Add(new IrDiffToken(IrDiffTokenKind.Separator, " ", " ", 0, 1, Plain));
            tokens.Add(W(words[i]));
        }
        return tokens;
    }

    /// <summary>A single non-whitespace Separator token (e.g. "." or "-") — content, not connective.</summary>
    private static IrDiffToken Sep(string s) =>
        new(IrDiffTokenKind.Separator, s, s, 0, s.Length, Plain);

    private static readonly IrRunFormat Plain = new() { Bold = false, UnmodeledDigest = default };
    private static readonly IrRunFormat Bold = new() { Bold = true, UnmodeledDigest = default };

    private static IrTokenDiff Diff(IReadOnlyList<IrDiffToken> l, IReadOnlyList<IrDiffToken> r)
    {
        var d = IrTokenDiffer.Diff(l, r, Default);
        IrTokenDiffAsserts.AssertInvariants(l, r, d);
        return d;
    }

    /// <summary>Compact op signature: "Kind(Lstart,Lend|Rstart,Rend)" joined by spaces.</summary>
    private static string Sig(IrTokenDiff d) =>
        string.Join(" ", d.Ops.Select(o => $"{o.Kind}({o.LeftStart},{o.LeftEnd}|{o.RightStart},{o.RightEnd})"));

    // --- core diff cases --------------------------------------------------

    [Fact]
    public void Single_word_change_in_the_middle()
    {
        var left = Words("the", "quick", "fox");
        var right = Words("the", "slow", "fox");
        var d = Diff(left, right);

        // "the " equal, "quick"->"slow" delete+insert, " fox" equal. The Insert's empty-left anchor
        // sits at the post-delete left cursor (3), the deterministic Delete-before-Insert convention.
        Assert.Equal(
            "Equal(0,2|0,2) Delete(2,3|2,2) Insert(3,3|2,3) Equal(3,5|3,5)",
            Sig(d));
    }

    [Fact]
    public void Prefix_edit_only()
    {
        var left = Words("alpha", "beta", "gamma");
        var right = Words("zeta", "beta", "gamma");
        var d = Diff(left, right);
        Assert.Equal(
            "Delete(0,1|0,0) Insert(1,1|0,1) Equal(1,5|1,5)",
            Sig(d));
    }

    [Fact]
    public void Suffix_edit_only()
    {
        var left = Words("alpha", "beta", "gamma");
        var right = Words("alpha", "beta", "omega");
        var d = Diff(left, right);
        Assert.Equal(
            "Equal(0,4|0,4) Delete(4,5|4,4) Insert(5,5|4,5)",
            Sig(d));
    }

    [Fact]
    public void Pure_insertion_at_end()
    {
        var left = Words("a", "b");
        var right = Words("a", "b", "c");
        var d = Diff(left, right);
        // "a b" equal, then " c" inserted.
        Assert.Equal("Equal(0,3|0,3) Insert(3,3|3,5)", Sig(d));
    }

    [Fact]
    public void Pure_deletion_at_end()
    {
        var left = Words("a", "b", "c");
        var right = Words("a", "b");
        var d = Diff(left, right);
        Assert.Equal("Equal(0,3|0,3) Delete(3,5|3,3)", Sig(d));
    }

    /// <summary>Word texts (Kind==Word) that fall inside an Equal op — i.e. retained across the diff.</summary>
    private static List<string> RetainedWords(IrTokenDiff d, IReadOnlyList<IrDiffToken> left)
    {
        var retained = new List<string>();
        foreach (var op in d.Ops)
            if (op.Kind == IrTokenOpKind.Equal)
                for (int i = op.LeftStart; i < op.LeftEnd; i++)
                    if (left[i].Kind == IrDiffTokenKind.Word)
                        retained.Add(left[i].Text);
        return retained;
    }

    [Fact]
    public void Content_anchored_retains_shared_word_amid_divergent_phrases()
    {
        // Two sentences sharing the content words "Open", "Sans", "a". Word anchors the intra-paragraph
        // diff on content words and retains them (rendering ins/del around the retained anchors). A
        // whitespace-anchored Myers instead spends its matches on the abundant identical spaces and
        // DROPS the interior "a" (delete+insert it), scattering the ink. Content-anchored diffing keeps
        // the shared content words retained.
        var left = Words("Open", "Sans", "bold", "underline", "creates", "a", "distinctive", "modern", "heading", "style");
        var right = Words("Open", "Sans", "is", "a", "humanist", "sans", "serif", "typeface");
        var d = Diff(left, right);
        var retained = RetainedWords(d, left);
        Assert.Contains("Open", retained);
        Assert.Contains("Sans", retained);
        Assert.Contains("a", retained);   // the interior shared content word must be retained, not dropped
    }

    /// <summary>Concatenate several token sub-lists into one paragraph token list.</summary>
    private static List<IrDiffToken> Concat(params List<IrDiffToken>[] parts)
    {
        var all = new List<IrDiffToken>();
        foreach (var p in parts) all.AddRange(p);
        return all;
    }

    [Fact]
    public void Content_anchored_retains_interior_word_in_real_whitespace_crowded_pair()
    {
        // The documented real regression pair (docs root cause):
        //   left : "Open Sans bold underline creates a distinctive modern heading style."
        //   right: "Open Sans is a humanist sans-serif typeface for digital use."
        // Both share the interior content word "a" (and "Open", "Sans"). With enough diverging words
        // on both sides plus the shared trailing "." (and the "-" in "sans-serif"), a whitespace-keyed
        // Myers spends its LCS budget on the abundant identical spaces and DROPS the interior "a",
        // deleting+re-inserting it. Word instead anchors on the CONTENT word "a" and retains it.
        // Content-anchored diffing keeps "a" in an Equal span.
        var left = Concat(
            Words("Open", "Sans", "bold", "underline", "creates", "a", "distinctive", "modern", "heading", "style"),
            new List<IrDiffToken> { Sep(".") });
        var right = Concat(
            Words("Open", "Sans", "is", "a", "humanist", "sans"),
            new List<IrDiffToken> { Sep("-") },
            Words("serif", "typeface", "for", "digital", "use"),
            new List<IrDiffToken> { Sep(".") });
        var d = Diff(left, right);
        var retained = RetainedWords(d, left);
        Assert.Contains("Open", retained);
        Assert.Contains("Sans", retained);
        Assert.Contains("a", retained);   // interior shared content word must be RETAINED, not dropped
    }

    [Fact]
    public void All_changed_no_common_tokens()
    {
        var left = Words("one", "two");
        var right = Words("three", "four");
        var d = Diff(left, right);
        // No shared CONTENT tokens => whole-left delete, whole-right insert. The interior separator " "
        // is connective, so under content-anchored diffing it never anchors: it lands inside a segment
        // bounded by non-matching content and is absorbed into the del/ins band rather than surviving as
        // Equal (whitespace ink follows the changed content, matching Word). Coverage asserts already run.
        Assert.Contains(IrTokenOpKind.Delete, d.Ops.Select(o => o.Kind));
        Assert.Contains(IrTokenOpKind.Insert, d.Ops.Select(o => o.Kind));
    }

    [Fact]
    public void All_changed_truly_disjoint_tokens()
    {
        // Disjoint single tokens (no shared separator) => one Delete + one Insert.
        var left = new List<IrDiffToken> { W("xxx") };
        var right = new List<IrDiffToken> { W("yyy") };
        var d = Diff(left, right);
        Assert.Equal("Delete(0,1|0,0) Insert(1,1|0,1)", Sig(d));
    }

    [Fact]
    public void All_equal_identical_lists()
    {
        var left = Words("same", "text", "here");
        var right = Words("same", "text", "here");
        var d = Diff(left, right);
        Assert.Equal("Equal(0,5|0,5)", Sig(d));
        Assert.Single(d.Ops);
    }

    [Fact]
    public void Separator_only_change()
    {
        // "a-b" (hyphen separator) vs "a b" (space separator): words equal, separator token changes.
        var left = new List<IrDiffToken>
        {
            W("a"),
            new(IrDiffTokenKind.Separator, "-", "-", 0, 1, Plain),
            W("b"),
        };
        var right = new List<IrDiffToken>
        {
            W("a"),
            new(IrDiffTokenKind.Separator, " ", " ", 0, 1, Plain),
            W("b"),
        };
        var d = Diff(left, right);
        // "a" equal, "-"->" " delete+insert, "b" equal.
        Assert.Equal(
            "Equal(0,1|0,1) Delete(1,2|1,1) Insert(2,2|1,2) Equal(2,3|2,3)",
            Sig(d));
    }

    // --- format post-pass -------------------------------------------------

    [Fact]
    public void Bold_word_becomes_FormatChanged_span_exactly_over_that_word()
    {
        // Content identical; the middle word goes bold. The Separator tokens around it keep Plain
        // format, so ONLY the bolded word's position is FormatChanged.
        var left = new List<IrDiffToken>
        {
            W("plain"),
            new(IrDiffTokenKind.Separator, " ", " ", 0, 1, Plain),
            W("word", Plain),
            new(IrDiffTokenKind.Separator, " ", " ", 0, 1, Plain),
            W("tail"),
        };
        var right = new List<IrDiffToken>
        {
            W("plain"),
            new(IrDiffTokenKind.Separator, " ", " ", 0, 1, Plain),
            W("word", Bold),
            new(IrDiffTokenKind.Separator, " ", " ", 0, 1, Plain),
            W("tail"),
        };
        var d = Diff(left, right);
        Assert.Equal(
            "Equal(0,2|0,2) FormatChanged(2,3|2,3) Equal(3,5|3,5)",
            Sig(d));
    }

    [Fact]
    public void Format_changes_separated_by_unchanged_separator_stay_distinct_spans()
    {
        // Words "a" and "b" both go bold but the separator between them keeps Plain format, so the
        // FormatChanged run is BROKEN by the still-Equal separator — two FormatChanged spans, not one.
        var left = Words("a", "b", "c"); // all Plain via the Words builder
        var right = new List<IrDiffToken>
        {
            W("a", Bold),
            new(IrDiffTokenKind.Separator, " ", " ", 0, 1, Plain),
            W("b", Bold),
            new(IrDiffTokenKind.Separator, " ", " ", 0, 1, Plain),
            W("c", Plain),
        };
        var d = Diff(left, right);
        Assert.Equal(
            "FormatChanged(0,1|0,1) Equal(1,2|1,2) FormatChanged(2,3|2,3) Equal(3,5|3,5)",
            Sig(d));
    }

    [Fact]
    public void Contiguous_format_changes_with_no_separator_merge()
    {
        // Two adjacent Word tokens (no separator between) both go bold => single FormatChanged span.
        var left = new List<IrDiffToken> { W("aa", Plain), W("bb", Plain) };
        var right = new List<IrDiffToken> { W("aa", Bold), W("bb", Bold) };
        var d = Diff(left, right);
        Assert.Equal("FormatChanged(0,2|0,2)", Sig(d));
    }

    // --- empty sides ------------------------------------------------------

    [Fact]
    public void Empty_left_yields_one_insert()
    {
        var right = Words("a", "b");
        var d = Diff(new List<IrDiffToken>(), right);
        Assert.Equal("Insert(0,0|0,3)", Sig(d));
    }

    [Fact]
    public void Empty_right_yields_one_delete()
    {
        var left = Words("a", "b");
        var d = Diff(left, new List<IrDiffToken>());
        Assert.Equal("Delete(0,3|0,0)", Sig(d));
    }

    [Fact]
    public void Empty_both_yields_no_ops()
    {
        var d = Diff(new List<IrDiffToken>(), new List<IrDiffToken>());
        Assert.Empty(d.Ops);
    }

    // --- adversarial repeated words --------------------------------------

    [Fact]
    public void Repeated_words_insert_one_more()
    {
        // "the the the" -> "the the the the": Myers finds the LCS of the three and inserts one "the"
        // (with its separator) at the greedy-preferred position. Pinned exactly.
        var left = Words("the", "the", "the");
        var right = Words("the", "the", "the", "the");
        var d = Diff(left, right);
        // The 5 left tokens ("the the the") match the right prefix; the extra " the" lands as a
        // trailing Insert. Pinned to catch any regression in the Myers tie-break / backtrace.
        Assert.Equal("Equal(0,5|0,5) Insert(5,5|5,7)", Sig(d));
    }

    [Fact]
    public void Repeated_words_with_distinct_tail()
    {
        // "the the the a the the" vs "the the a the the the": deterministic, invariants hold, and the
        // op shape is pinned so a regression in the anchor tie-break is caught.
        var left = Words("the", "the", "the", "a", "the", "the");
        var right = Words("the", "the", "a", "the", "the", "the");
        var d = Diff(left, right);
        // Stable, deterministic signature (regression pin). The char-weighted anchor rule keeps all five
        // "the" anchors (matched char weight 15) rather than pairing the lone "a" (weight 13) — both are
        // equally-minimal two-edit scripts, so this pins the char-weight tie-break: the right "a " is
        // inserted early and the left "a " deleted, leaving the surrounding "the"s Equal.
        Assert.Equal(
            "Equal(0,4|0,4) Insert(4,4|4,6) Equal(4,6|6,8) Delete(6,8|8,8) Equal(8,11|8,11)",
            Sig(d));
    }

    [Fact]
    public void Determinism_same_inputs_same_ops()
    {
        var left = Words("the", "the", "the", "a", "the", "the");
        var right = Words("the", "the", "a", "the", "the", "the");
        var d1 = IrTokenDiffer.Diff(left, right, Default);
        var d2 = IrTokenDiffer.Diff(left, right, Default);
        Assert.Equal(d1, d2); // record equality over the whole IrTokenDiff
    }

    // --- real IR end-to-end ----------------------------------------------

    [Fact]
    public void Real_IR_paragraphs_diff_end_to_end()
    {
        var options = new IrReaderOptions { RetainSources = false };
        var leftPara = IrReader.Read(IrTestDocuments.Create("The quick brown fox"), options)
            .Body.Blocks.OfType<IrParagraph>().First();
        var rightPara = IrReader.Read(IrTestDocuments.Create("The slow brown fox"), options)
            .Body.Blocks.OfType<IrParagraph>().First();

        var left = IrDiffTokenizer.Tokenize(leftPara, Default);
        var right = IrDiffTokenizer.Tokenize(rightPara, Default);
        var d = Diff(left, right);

        // The differ classified some change (the "quick"->"slow" word) while keeping the rest equal.
        Assert.Contains(IrTokenOpKind.Delete, d.Ops.Select(o => o.Kind));
        Assert.Contains(IrTokenOpKind.Insert, d.Ops.Select(o => o.Kind));
        Assert.Contains(IrTokenOpKind.Equal, d.Ops.Select(o => o.Kind));
    }
}
