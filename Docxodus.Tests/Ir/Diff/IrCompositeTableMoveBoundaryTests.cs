#nullable enable

using System.Collections.Generic;
using System.Linq;
using Docxodus;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Issue #229 — a reviewer's paragraph move that crosses a table boundary must not contest the whole
/// table block. Base is <c>paragraph A</c> + a 2x2 table + <c>paragraph B</c>. When Alice relocates
/// <c>paragraph B</c> to just before the table (crossing the boundary) and Bob edits a disjoint table
/// cell, both edits must apply independently — a native/lowered paragraph move plus a per-cell compose —
/// with <c>conflicts == 0</c> and <c>reject ≡ base</c>. Before the fix the block aligner reads the TABLE
/// as the displaced block, so the table collapses to a whole-table block conflict and Bob's cell edit is
/// surfaced in the conflict rather than composed per-cell.
/// </summary>
public class IrCompositeTableMoveBoundaryTests
{
    // Distinctive ≥4-word paragraphs so a reorder is detectable as a move.
    private const string A = "Alpha paragraph one bravo charlie";
    private const string B = "Bravo paragraph two delta echo";

    private static string Cell(string text) =>
        $"<w:tc><w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p></w:tc>";

    private static string Row(params string[] cells) => $"<w:tr>{string.Concat(cells)}</w:tr>";

    private static string Table(params string[] rows) =>
        $"<w:tbl><w:tblPr/><w:tblGrid/>{string.Concat(rows)}</w:tbl>";

    private static string Para(string text) =>
        $"<w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p>";

    private static WmlDocument Base(string c11 = "d four") => IrTestDocuments.FromBodyXml(
        Para(A) +
        Table(
            Row(Cell("a one"), Cell("b two")),
            Row(Cell("c three"), Cell(c11))) +
        Para(B));

    /// <summary>Alice: paragraph B relocated to just BEFORE the table (crossing the table boundary).</summary>
    private static WmlDocument AliceMovesBBeforeTable() => IrTestDocuments.FromBodyXml(
        Para(A) +
        Para(B) +
        Table(
            Row(Cell("a one"), Cell("b two")),
            Row(Cell("c three"), Cell("d four"))));

    /// <summary>Bob: edits one disjoint table cell (cell(1,1) "d four" → "d BOB"); paragraph order unchanged.</summary>
    private static WmlDocument BobEditsCell() => IrTestDocuments.FromBodyXml(
        Para(A) +
        Table(
            Row(Cell("a one"), Cell("b two")),
            Row(Cell("c three"), Cell("d BOB"))) +
        Para(B));

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

    private static string Accept(WmlDocument merged) =>
        Docs.PlainTextWithTables(RevisionAccepter.AcceptRevisions(merged));

    private static string Reject(WmlDocument merged) =>
        Docs.StructuralBody(RevisionProcessor.RejectRevisions(merged));

    /// <summary>
    /// #229 — a paragraph move crossing a table boundary + a disjoint table-cell edit both apply, no
    /// conflict, no content loss, reject ≡ base. This is the acceptance criterion.
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Paragraph_move_across_table_boundary_composes_with_disjoint_cell_edit(ConflictResolution policy)
    {
        var baseDoc = Base();
        var alice = AliceMovesBBeforeTable();
        var bob = BobEditsCell();

        // (1) No conflict: the move and the disjoint cell edit are independent.
        Assert.Empty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));

        // (2) Bob's disjoint cell edit is composed (lands on accept).
        var accept = Accept(merged);
        Assert.Contains("d BOB", accept);
        Assert.DoesNotContain("d four", accept);

        // (3) Alice's move landed: paragraph B now precedes the table's cell content on accept.
        Assert.True(
            accept.IndexOf(B, System.StringComparison.Ordinal) < accept.IndexOf("a one", System.StringComparison.Ordinal),
            $"paragraph B should precede the table after the accepted move; accept body: [{accept}]");
        // No duplication of the moved paragraph.
        Assert.Equal(1, CountOccurrences(accept, B));

        // (4) reject ≡ base under every policy — no content loss.
        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
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
