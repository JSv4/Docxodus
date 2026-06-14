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
    public void Multireviewer_disjoint_table_cell_edits_record_conflict_not_silent_drop(ConflictResolution policy)
    {
        var baseDoc = TableBase();
        var alice = TableVariant("a ALICE", "b two", "c three", "d four"); // edits cell (0,0)
        var bob = TableVariant("a one", "b two", "c three", "d BOB");      // edits cell (1,1) — disjoint

        // Both reviewers' table edits must be SEEN: a recorded conflict, not a silent drop.
        Assert.True(
            Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)).Count >= 1,
            "Multi-reviewer table cell edits were silently dropped (conflictCount == 0).");

        // reject ≡ base must hold for the table-conflict output under every policy.
        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
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
}
