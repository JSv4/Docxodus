#nullable enable

using System.Collections.Generic;
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

    /// <summary>The aligner invariants the plan pins — run against EVERY case's output.</summary>
    private static void AssertInvariants(IrDocument left, IrDocument right, IrBlockAlignment a)
    {
        var leftSeen = new List<IrBlock>();
        var rightSeen = new List<IrBlock>();

        foreach (var e in a.Entries)
        {
            switch (e.Kind)
            {
                case IrAlignmentKind.Inserted:
                    Assert.Null(e.Left);
                    Assert.NotNull(e.Right);
                    break;
                case IrAlignmentKind.Deleted:
                    Assert.NotNull(e.Left);
                    Assert.Null(e.Right);
                    break;
                case IrAlignmentKind.Unchanged:
                    Assert.NotNull(e.Left);
                    Assert.NotNull(e.Right);
                    Assert.Equal(e.Left!.ContentHash, e.Right!.ContentHash);
                    Assert.Equal(e.Left!.FormatFingerprint, e.Right!.FormatFingerprint);
                    break;
                case IrAlignmentKind.FormatOnly:
                    Assert.NotNull(e.Left);
                    Assert.NotNull(e.Right);
                    Assert.Equal(e.Left!.ContentHash, e.Right!.ContentHash);
                    Assert.NotEqual(e.Left!.FormatFingerprint, e.Right!.FormatFingerprint);
                    break;
                case IrAlignmentKind.Moved:
                    Assert.NotNull(e.Left);
                    Assert.NotNull(e.Right);
                    Assert.Equal(e.Left!.ContentHash, e.Right!.ContentHash);
                    break;
                case IrAlignmentKind.Modified:
                    Assert.NotNull(e.Left);
                    Assert.NotNull(e.Right);
                    break;
                case IrAlignmentKind.MovedModified:
                    Assert.Fail("MovedModified must never be produced in M2.1.");
                    break;
            }

            if (e.Left is not null)
                leftSeen.Add(e.Left);
            if (e.Right is not null)
                rightSeen.Add(e.Right);
        }

        // Every left/right body block appears in exactly one entry (totality + no duplication).
        AssertSameMultiset(left.Body.Blocks, leftSeen, "left");
        AssertSameMultiset(right.Body.Blocks, rightSeen, "right");
    }

    private static void AssertSameMultiset(IReadOnlyList<IrBlock> expected, List<IrBlock> seen, string side)
    {
        Assert.Equal(expected.Count, seen.Count);
        // Reference identity: the aligner must return the very block instances from the input lists.
        var pool = new List<IrBlock>(expected);
        foreach (var b in seen)
        {
            int idx = pool.FindIndex(x => ReferenceEquals(x, b));
            Assert.True(idx >= 0, $"{side} block appeared that was not in the input (or appeared twice).");
            pool.RemoveAt(idx);
        }
        Assert.Empty(pool);
    }

    private static int Count(IrBlockAlignment a, IrAlignmentKind k) => a.Entries.Count(e => e.Kind == k);

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
        var l = Doc("alpha", "beta", "gamma");
        var r = Doc("alpha", "BETA-edited", "gamma");
        var a = Align(l, r);

        Assert.Equal(2, Count(a, IrAlignmentKind.Unchanged));
        Assert.Equal(1, Count(a, IrAlignmentKind.Modified));
        Assert.Equal(0, Count(a, IrAlignmentKind.Moved));
        AssertInvariants(l, r, a);
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
        // independently of the move. (When a move instead reshuffles the gap boundaries so the edited
        // pair lands in DIFFERENT gaps, the exact-hash aligner classifies the edit as Delete+Insert
        // rather than Modified — a documented M2.1 consequence of gap-positional pairing.)
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
