#nullable enable

using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// FOLLOW-ON B: per-cell table composition in the consolidate engine. Today multi-reviewer edits to the
/// SAME base table route to a whole-block conflict (BaseWins keeps the base table; disjoint cell edits are
/// NOT composed). These tests pin the NEW behavior: DISJOINT cross-reviewer table-cell edits COMPOSE inline
/// (Alice edits cell(0,0), Bob edits cell(1,1) → both land, attributed); only SAME-cell edits by ≥2
/// reviewers → recorded conflict per policy. STOP boundaries (MovedRow / column-count / cell-tcPr changes)
/// fall back to the whole-table block conflict.
/// </summary>
public class IrCompositeTableTests
{
    // ------------------------------------------------------------------ fixtures

    private static string Cell(string text) =>
        $"<w:tc><w:p><w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r></w:p></w:tc>";

    private static string Row(params string[] cells) => $"<w:tr>{string.Concat(cells)}</w:tr>";

    private static string Table(params string[] rows) =>
        $"<w:tbl><w:tblPr/><w:tblGrid/>{string.Concat(rows)}</w:tbl>";

    private static string Lead => "<w:p><w:r><w:t>lead</w:t></w:r></w:p>";

    /// <summary>A 2x2 table base; one paragraph above so the body has stable surrounding structure.</summary>
    private static WmlDocument Base2x2() => IrTestDocuments.FromBodyXml(
        Lead +
        Table(
            Row(Cell("a one"), Cell("b two")),
            Row(Cell("c three"), Cell("d four"))));

    private static WmlDocument Variant2x2(string c00, string c01, string c10, string c11) =>
        IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell(c00), Cell(c01)),
                Row(Cell(c10), Cell(c11))));

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

    private static IReadOnlyList<DocxDiffConsolidatedRevision> Revisions(
        WmlDocument baseDoc, ConflictResolution policy,
        params (string Author, WmlDocument Doc)[] reviewers)
        => DocxDiff.GetConsolidatedRevisions(
            baseDoc,
            reviewers.Select(r => new DocxDiffReviewer { Author = r.Author, Document = r.Doc }).ToList(),
            new DocxDiffConsolidateSettings { ConflictResolution = policy });

    private static string Json(
        WmlDocument baseDoc, ConflictResolution policy,
        params (string Author, WmlDocument Doc)[] reviewers)
        => DocxDiff.GetConsolidatedEditScriptJson(
            baseDoc,
            reviewers.Select(r => new DocxDiffReviewer { Author = r.Author, Document = r.Doc }).ToList(),
            new DocxDiffConsolidateSettings { ConflictResolution = policy });

    private static string Accept(WmlDocument merged) =>
        Docs.PlainTextWithTables(RevisionAccepter.AcceptRevisions(merged));

    private static string Reject(WmlDocument merged) =>
        Docs.StructuralBody(RevisionProcessor.RejectRevisions(merged));

    // ------------------------------------------------------------------ 1. disjoint cells compose

    /// <summary>
    /// #1 — Disjoint cells compose. 2x2 base, Alice edits cell(0,0), Bob edits cell(1,1) → both land
    /// authored, Conflicts == 0. (RED before per-cell composition: whole-table conflict / one dropped.)
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Disjoint_cells_compose_both_land_no_conflict(ConflictResolution policy)
    {
        var baseDoc = Base2x2();
        var alice = Variant2x2("a ALICE", "b two", "c three", "d four"); // cell(0,0)
        var bob = Variant2x2("a one", "b two", "c three", "d BOB");      // cell(1,1) — disjoint

        // No conflict: disjoint cell edits compose.
        Assert.Empty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        // Accept yields BOTH edits regardless of policy (disjoint → policy-invariant).
        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var accept = Accept(merged);
        Assert.Contains("a ALICE", accept);
        Assert.Contains("d BOB", accept);
        Assert.Contains("b two", accept);
        Assert.Contains("c three", accept);

        // reject ≡ base.
        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
    }

    // ------------------------------------------------------------------ 2. consensus cell

    /// <summary>#2 — Consensus cell: both edit cell(0,0) identically → one cell, no conflict.</summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Consensus_cell_no_conflict(ConflictResolution policy)
    {
        var baseDoc = Base2x2();
        var alice = Variant2x2("a SAME", "b two", "c three", "d four");
        var bob = Variant2x2("a SAME", "b two", "c three", "d four"); // identical edit

        Assert.Empty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var accept = Accept(merged);
        Assert.Contains("a SAME", accept);
        // The consensus edit appears once (no duplicated "SAME").
        Assert.Equal(1, CountOccurrences(accept, "SAME"));

        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
    }

    // ------------------------------------------------------------------ 3. same-cell different-words conflict

    /// <summary>
    /// #3 — Same-cell different-words conflict. Alice "the SLOW fox" / Bob "the QUICK fox" in cell(0,0) →
    /// a token conflict recorded at the cell-paragraph anchor; resolved per policy.
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Same_cell_different_words_conflict(ConflictResolution policy)
    {
        var baseDoc = Variant2x2("the fox", "b two", "c three", "d four");
        var alice = Variant2x2("the SLOW fox", "b two", "c three", "d four");
        var bob = Variant2x2("the QUICK fox", "b two", "c three", "d four");

        var conflicts = Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        Assert.NotEmpty(conflicts);
        // The conflict competitors carry both reviewers' competing inserted words.
        var competitorTexts = conflicts.SelectMany(c => c.Competitors.Select(x => x.ResultText)).ToList();
        Assert.Contains(competitorTexts, t => t.Contains("SLOW"));
        Assert.Contains(competitorTexts, t => t.Contains("QUICK"));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var accept = Accept(merged);
        switch (policy)
        {
            case ConflictResolution.BaseWins:
                Assert.DoesNotContain("SLOW", accept);
                Assert.DoesNotContain("QUICK", accept);
                break;
            case ConflictResolution.FirstReviewerWins:
                Assert.Contains("SLOW", accept);
                Assert.DoesNotContain("QUICK", accept);
                break;
            case ConflictResolution.StackAll:
                Assert.Contains("SLOW", accept);
                Assert.Contains("QUICK", accept);
                break;
        }

        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
    }

    // ------------------------------------------------------------------ 4. same-cell disjoint words (recursion)

    /// <summary>
    /// #4 — Same-cell disjoint words: both reviewers edit DIFFERENT words of one cell paragraph → compose
    /// inside the cell (recursion proof: token composition runs at cell-paragraph granularity).
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Same_cell_disjoint_words_compose_inside_cell(ConflictResolution policy)
    {
        var baseDoc = Variant2x2("alpha beta gamma", "b two", "c three", "d four");
        var alice = Variant2x2("ALPHA beta gamma", "b two", "c three", "d four"); // first word
        var bob = Variant2x2("alpha beta GAMMA", "b two", "c three", "d four");   // last word

        // Disjoint words inside the SAME cell paragraph compose: no conflict.
        Assert.Empty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var accept = Accept(merged);
        Assert.Contains("ALPHA", accept);
        Assert.Contains("GAMMA", accept);
        Assert.Contains("beta", accept);

        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
    }

    // ------------------------------------------------------------------ 5. disjoint InsertRow by two reviewers

    /// <summary>#5 — Disjoint InsertRow by two reviewers → both rows appear, no conflict.</summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Disjoint_inserted_rows_both_appear(ConflictResolution policy)
    {
        var baseDoc = Base2x2();
        var alice = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("a one"), Cell("b two")),
                Row(Cell("c three"), Cell("d four")),
                Row(Cell("e ALICE"), Cell("f alice")))); // appends a row
        var bob = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("a one"), Cell("b two")),
                Row(Cell("c three"), Cell("d four")),
                Row(Cell("g BOB"), Cell("h bob")))); // appends a different row

        Assert.Empty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var accept = Accept(merged);
        Assert.Contains("e ALICE", accept);
        Assert.Contains("g BOB", accept);

        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
    }

    // ------------------------------------------------------------------ 6. DeleteRow (A) vs ModifyRow other row (B)

    /// <summary>#6 — DeleteRow (Alice deletes row 1) vs ModifyRow different row (Bob edits row 0) → both land.</summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void DeleteRow_vs_modify_other_row_both_land(ConflictResolution policy)
    {
        // 3-row base so a clean row delete leaves a stable spine.
        var baseDoc = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("r0 a"), Cell("r0 b")),
                Row(Cell("r1 a"), Cell("r1 b")),
                Row(Cell("r2 a"), Cell("r2 b"))));
        var alice = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("r0 a"), Cell("r0 b")),
                Row(Cell("r2 a"), Cell("r2 b")))); // deletes row 1
        var bob = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("r0 a EDIT"), Cell("r0 b")),
                Row(Cell("r1 a"), Cell("r1 b")),
                Row(Cell("r2 a"), Cell("r2 b")))); // edits row 0 cell(0,0)

        Assert.Empty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var accept = Accept(merged);
        Assert.Contains("EDIT", accept);          // Bob's row-0 edit landed
        Assert.DoesNotContain("r1 a", accept);    // Alice's row-1 delete landed (row gone on accept)

        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
    }

    // ------------------------------------------------------------------ 7. DeleteRow vs ModifyRow SAME row

    /// <summary>#7 — DeleteRow vs ModifyRow SAME row → row conflict, policy-resolved, reject ≡ base.</summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void DeleteRow_vs_modify_same_row_conflict(ConflictResolution policy)
    {
        var baseDoc = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("r0 a"), Cell("r0 b")),
                Row(Cell("r1 a"), Cell("r1 b")),
                Row(Cell("r2 a"), Cell("r2 b"))));
        var alice = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("r0 a"), Cell("r0 b")),
                Row(Cell("r2 a"), Cell("r2 b")))); // deletes row 1
        var bob = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("r0 a"), Cell("r0 b")),
                Row(Cell("r1 a EDIT"), Cell("r1 b")),
                Row(Cell("r2 a"), Cell("r2 b")))); // edits row 1 (the one Alice deletes)

        Assert.NotEmpty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
    }

    // ------------------------------------------------------------------ 8. MovedRow fallback (STOP boundary)

    /// <summary>
    /// #8 — MovedRow present → block-level conflict fallback (STOP boundary). When any reviewer's table diff
    /// has a MovedRow, the per-cell compose bails to the existing whole-table block conflict. No silent loss;
    /// reject ≡ base.
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void MovedRow_falls_back_to_block_conflict(ConflictResolution policy)
    {
        // A 3-row table with UNIQUE row hashes so the row aligner anchors them; Alice reorders rows
        // (a MovedRow), Bob edits a cell — two reviewers touch the same base table.
        var baseDoc = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("alpha row")),
                Row(Cell("bravo row")),
                Row(Cell("charlie row"))));
        var alice = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("bravo row")),
                Row(Cell("alpha row")),   // alpha moved down (a row move)
                Row(Cell("charlie row"))));
        var bob = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("alpha row")),
                Row(Cell("bravo row EDIT")),
                Row(Cell("charlie row"))));

        // Falls back to a recorded whole-table conflict — no silent loss.
        Assert.NotEmpty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
    }

    // ------------------------------------------------------------------ 9. column-count change fallback (gate)

    /// <summary>
    /// #9 — Column add/remove → block-level conflict fallback (gate test). When a reviewer adds a column
    /// (changing the cell count of a row) while another edits a different cell, the positional per-cell render
    /// (which clones the base row's cell skeleton, count-stable in v1) cannot compose the column change — so it
    /// bails to a recorded whole-table block conflict. No silent loss: a conflict is recorded under every
    /// policy, and BaseWins (which keeps the base table) rejects to base.
    /// <para>(A pure <c>w:tcPr</c> width change is NOT modeled by the IR — it leaves every hash identical, so
    /// the reviewer's table reads as EqualBlock (no touch) and composes trivially as single-source; only a
    /// genuine cell-count change is a detectable STOP boundary. reject ≡ base is NOT asserted for
    /// FirstReviewerWins/StackAll here: those emit the reviewer's column-changed table, and a 2-way column
    /// ADD is not yet marked as a tracked insertion — a pre-existing whole-table-renderer limitation,
    /// independent of this follow-on. The fallback's job is no-silent-loss, which holds.)</para>
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Column_count_change_falls_back_to_block_conflict(ConflictResolution policy)
    {
        var baseDoc = Base2x2();
        var alice = IrTestDocuments.FromBodyXml(
            Lead +
            Table(
                Row(Cell("a one"), Cell("b two"), Cell("NEW col")), // adds a 3rd column to row 0
                Row(Cell("c three"), Cell("d four"))));
        var bob = Variant2x2("a one", "b two", "c three", "d BOB"); // edits cell(1,1) text

        // Column add → fall back to recorded whole-table conflict (no silent loss).
        Assert.NotEmpty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        // BaseWins keeps the base table → reject ≡ base (the fallback's reject contract on the base-kept path).
        if (policy == ConflictResolution.BaseWins)
        {
            var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
            Assert.Equal(Docs.StructuralBody(baseDoc), Reject(merged));
        }
    }

    // ------------------------------------------------------------------ 10. all 3 policies over same-cell conflict

    /// <summary>#10 — All three policies over a same-cell conflict (accept text + competitors).</summary>
    [Fact]
    public void Same_cell_conflict_all_policies()
    {
        var baseDoc = Variant2x2("base word", "b two", "c three", "d four");
        var alice = Variant2x2("ALICE word", "b two", "c three", "d four");
        var bob = Variant2x2("BOB word", "b two", "c three", "d four");

        // BaseWins: neither competitor's word in accept.
        var bw = Accept(Consolidate(baseDoc, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob)));
        Assert.DoesNotContain("ALICE", bw);
        Assert.DoesNotContain("BOB", bw);

        // FirstReviewerWins: Alice's word, not Bob's.
        var fw = Accept(Consolidate(baseDoc, ConflictResolution.FirstReviewerWins, ("Alice", alice), ("Bob", bob)));
        Assert.Contains("ALICE", fw);
        Assert.DoesNotContain("BOB", fw);

        // StackAll: both.
        var sa = Accept(Consolidate(baseDoc, ConflictResolution.StackAll, ("Alice", alice), ("Bob", bob)));
        Assert.Contains("ALICE", sa);
        Assert.Contains("BOB", sa);

        // Competitors recorded with both authors.
        var conflicts = Conflicts(baseDoc, ConflictResolution.StackAll, ("Alice", alice), ("Bob", bob));
        var authors = conflicts.SelectMany(c => c.Competitors.Select(x => x.Author)).Distinct().ToList();
        Assert.Contains("Alice", authors);
        Assert.Contains("Bob", authors);
    }

    // ------------------------------------------------------------------ 11. reject ≡ base all policies

    /// <summary>
    /// #11 — reject ≡ base for all three policies (markup → RejectRevisions → StructuralBody == base),
    /// across disjoint + same-cell + delete-row scenarios.
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Reject_equals_base_all_policies(ConflictResolution policy)
    {
        var baseDoc = Base2x2();

        var disjoint = Consolidate(baseDoc, policy,
            ("Alice", Variant2x2("a ALICE", "b two", "c three", "d four")),
            ("Bob", Variant2x2("a one", "b two", "c three", "d BOB")));
        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(disjoint));

        var sameCell = Consolidate(baseDoc, policy,
            ("Alice", Variant2x2("a ALICE", "b two", "c three", "d four")),
            ("Bob", Variant2x2("a BOB", "b two", "c three", "d four")));
        Assert.Equal(Docs.StructuralBody(baseDoc), Reject(sameCell));
    }

    // ------------------------------------------------------------------ 12. renderer single w:tbl, authored cells

    /// <summary>
    /// #12 — Renderer: composed table = single w:tbl, per-cell w:ins/w:del authored to right reviewers, base
    /// cells verbatim; accept yields all disjoint edits.
    /// </summary>
    [Fact]
    public void Renderer_composed_table_single_tbl_authored_cells()
    {
        var baseDoc = Base2x2();
        var alice = Variant2x2("a ALICE", "b two", "c three", "d four");
        var bob = Variant2x2("a one", "b two", "c three", "d BOB");

        var merged = Consolidate(baseDoc, ConflictResolution.StackAll, ("Alice", alice), ("Bob", bob));
        var xml = Docs.MainPartXml(merged);

        // Exactly one table.
        Assert.Equal(1, CountOccurrences(xml, "<w:tbl>"));
        // Both reviewers' edits are present as tracked insertions attributed to them.
        Assert.Contains("w:author=\"Alice\"", xml);
        Assert.Contains("w:author=\"Bob\"", xml);
        // Untouched cells survive verbatim.
        Assert.Contains("b two", xml);
        Assert.Contains("c three", xml);

        var accept = Accept(merged);
        Assert.Contains("a ALICE", accept);
        Assert.Contains("d BOB", accept);
    }

    // ------------------------------------------------------------------ 13. revisions per reviewer / per conflict

    /// <summary>
    /// #13 — Revisions: disjoint-cell → revisions per reviewer; same-cell conflict → competing revisions
    /// linked to the cell conflict id.
    /// </summary>
    [Fact]
    public void Revisions_disjoint_per_reviewer_and_samecell_linked_to_conflict()
    {
        var baseDoc = Base2x2();

        // Disjoint: a revision for each reviewer, attributed.
        var disjointRevs = Revisions(baseDoc, ConflictResolution.StackAll,
            ("Alice", Variant2x2("a ALICE", "b two", "c three", "d four")),
            ("Bob", Variant2x2("a one", "b two", "c three", "d BOB")));
        Assert.Contains(disjointRevs, r => r.Author == "Alice");
        Assert.Contains(disjointRevs, r => r.Author == "Bob");

        // Same-cell conflict: competing revisions carry a conflict id matching a recorded conflict.
        var scBase = Variant2x2("base word", "b two", "c three", "d four");
        var scAlice = Variant2x2("ALICE word", "b two", "c three", "d four");
        var scBob = Variant2x2("BOB word", "b two", "c three", "d four");
        var conflicts = Conflicts(scBase, ConflictResolution.StackAll, ("Alice", scAlice), ("Bob", scBob));
        var revs = Revisions(scBase, ConflictResolution.StackAll, ("Alice", scAlice), ("Bob", scBob));
        var conflictIds = conflicts.Select(c => c.Id).ToHashSet();
        Assert.Contains(revs, r => r.ConflictId is { } id && conflictIds.Contains(id));
    }

    // ------------------------------------------------------------------ 14. JSON merged tableDiff + authoredRows

    /// <summary>
    /// #14 — JSON: composed table op serializes merged tableDiff + additive authoredRows; deterministic.
    /// </summary>
    [Fact]
    public void Json_composed_table_serializes_tableDiff_and_authoredRows_deterministic()
    {
        var baseDoc = Base2x2();
        var alice = Variant2x2("a ALICE", "b two", "c three", "d four");
        var bob = Variant2x2("a one", "b two", "c three", "d BOB");

        var json1 = Json(baseDoc, ConflictResolution.StackAll, ("Alice", alice), ("Bob", bob));
        var json2 = Json(baseDoc, ConflictResolution.StackAll, ("Alice", alice), ("Bob", bob));

        // Deterministic.
        Assert.Equal(json1, json2);
        // The merged tableDiff is present (apply/json truth).
        Assert.Contains("\"tableDiff\"", json1);
        Assert.Contains("\"rowOps\"", json1);
        // The additive authoredRows attribution view is present.
        Assert.Contains("\"authoredRows\"", json1);
    }

    // ------------------------------------------------------------------ 15. verifier over disjoint + conflict

    /// <summary>
    /// #15 — Verifier proves the table compose: the composite apply-reconstruction (cell ops applied to the
    /// base table) matches the rendered accepted body, for disjoint and same-cell conflict cases, all policies.
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Verifier_proves_table_compose_all_policies(ConflictResolution policy)
    {
        var baseDoc = Base2x2();

        var disjointMerged = Consolidate(baseDoc, policy,
            ("Alice", Variant2x2("a ALICE", "b two", "c three", "d four")),
            ("Bob", Variant2x2("a one", "b two", "c three", "d BOB")));
        IrCompositeVerifier.Verify(
            baseDoc,
            new (string, WmlDocument)[]
            {
                ("Alice", Variant2x2("a ALICE", "b two", "c three", "d four")),
                ("Bob", Variant2x2("a one", "b two", "c three", "d BOB")),
            },
            policy,
            Docs.AcceptStructuralBody(disjointMerged));

        var scBase = Variant2x2("base word", "b two", "c three", "d four");
        var scAlice = Variant2x2("ALICE word", "b two", "c three", "d four");
        var scBob = Variant2x2("BOB word", "b two", "c three", "d four");
        var scMerged = Consolidate(scBase, policy, ("Alice", scAlice), ("Bob", scBob));
        IrCompositeVerifier.Verify(
            scBase,
            new (string, WmlDocument)[] { ("Alice", scAlice), ("Bob", scBob) },
            policy,
            Docs.AcceptStructuralBody(scMerged));
    }

    // ------------------------------------------------------------------ 16. image in a composed cell

    private const string ANs = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string RNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    /// <summary>An inline-picture <c>w:drawing</c> whose <c>a:blip</c> references <paramref name="embedId"/>
    /// (mirrors the shape <c>IrNoteImageSdtTests.Drawing</c> uses).</summary>
    private static string Drawing(string embedId) =>
        "<w:drawing>" +
          "<wp:inline>" +
            "<wp:extent cx=\"100\" cy=\"100\"/>" +
            "<wp:docPr id=\"1\" name=\"Picture 1\"/>" +
            "<a:graphic><a:graphicData>" +
              "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                $"<pic:blipFill><a:blip r:embed=\"{embedId}\"/></pic:blipFill>" +
              "</pic:pic>" +
            "</a:graphicData></a:graphic>" +
          "</wp:inline>" +
        "</w:drawing>";

    /// <summary>A distinct tiny PNG: the shared <see cref="IrTestDocuments.TinyPng"/> bytes with the final
    /// byte replaced by <paramref name="tag"/>, so each reviewer's inserted image has DIFFERENT bytes even
    /// though both embed under the SAME rel id <c>rId51</c> in their respective (independent) packages — the
    /// cross-package rel-id COLLISION the stale whole-row media bucket mis-resolves.</summary>
    private static byte[] TaggedPng(byte tag)
    {
        var b = (byte[])IrTestDocuments.TinyPng.Clone();
        b[b.Length - 1] = tag;
        return b;
    }

    private static byte[] AlicePng => TaggedPng(0xA1);   // Alice inserts this into cell(0,0)
    private static byte[] BobPng => TaggedPng(0xB2);      // Bob inserts this into cell(0,1)

    /// <summary>An image-only paragraph embedding the SAME rel id <c>rId51</c> (so two reviewers' packages
    /// collide on that id while holding different bytes).</summary>
    private static string ImagePara() => $"<w:p><w:r>{Drawing("rId51")}</w:r></w:p>";

    /// <summary>A 2x2 table; row 0 cell(0,0) appends <paramref name="c00Extra"/>, cell(0,1) appends
    /// <paramref name="c01Extra"/> after a single base text paragraph; row 1 is plain text. (An inserted
    /// image lives in its OWN appended paragraph so it renders as a whole-block insert that round-trips —
    /// a trailing inline image inside an edited text paragraph is a separate, pre-existing slice concern.)</summary>
    private static string ImageCellTableBody(string c00Extra, string c01Extra) =>
        Lead +
        "<w:tbl><w:tblPr/><w:tblGrid/>" +
          "<w:tr>" +
            $"<w:tc><w:p><w:r><w:t xml:space=\"preserve\">logo</w:t></w:r></w:p>{c00Extra}</w:tc>" +
            $"<w:tc><w:p><w:r><w:t xml:space=\"preserve\">pic</w:t></w:r></w:p>{c01Extra}</w:tc>" +
          "</w:tr>" +
          Row(Cell("c three"), Cell("d four")) +
        "</w:tbl>";

    /// <summary>
    /// #16 — Image in a COMPOSED cell, cross-reviewer same-row, colliding rel id. Base row 0 is plain text in
    /// both cells. Alice (the FIRST table-toucher → the composed-table op's leftover <c>RightSourceId</c> = 0)
    /// appends an image paragraph (<see cref="AlicePng"/>) to cell(0,0); Bob appends an image paragraph
    /// (<see cref="BobPng"/>) to cell(0,1) — BOTH embedding the SAME id <c>rId51</c> in their independent
    /// packages. The two disjoint cells fall in ONE composed ModifyRow whose cells are sourced from DIFFERENT
    /// reviewers. With the (removed) buggy whole-row registration, Bob's image clone — correctly bucketed to
    /// reviewer 1 inside the per-cell render — was RE-registered under the leftover reviewer-0 (Alice) bucket;
    /// the bucket-0 import (which runs FIRST: buckets are key-ordered) then resolved Bob's clone's
    /// <c>rId51</c> against ALICE's package and pulled <see cref="AlicePng"/> into Bob's cell — a WRONG image
    /// that still resolves to a valid part. This pins the per-cell-only registration: every output embed
    /// resolves to an ImagePart whose bytes are the EXPECTED per-cell image (resolved set ==
    /// {AlicePng, BobPng}, not a collapsed {AlicePng}); accept keeps both edits + both correct images; reject
    /// ≡ base (both image inserts gone).
    /// </summary>
    [Theory]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Image_in_composed_cell_resolves_to_correct_per_cell_part(ConflictResolution policy)
    {
        var baseDoc = IrTestDocuments.FromBodyXmlWithImageParts(ImageCellTableBody("", ""));
        var alice = IrTestDocuments.FromBodyXmlWithImageParts(
            ImageCellTableBody(ImagePara(), ""), ("rId51", AlicePng));   // image into cell(0,0)
        var bob = IrTestDocuments.FromBodyXmlWithImageParts(
            ImageCellTableBody("", ImagePara()), ("rId51", BobPng));     // image into cell(0,1)

        // Disjoint cells (same row) compose: no conflict.
        Assert.Empty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));

        // (1) Opens cleanly; every body embed resolves to a real ImagePart in the OUTPUT package.
        AssertAllEmbedsResolve(merged);

        // (2) Accept keeps both reviewers' image inserts AND both correct per-cell images.
        var accepted = RevisionAccepter.AcceptRevisions(merged);
        AssertAllEmbedsResolve(accepted);
        // The two row-0 cells keep their OWN images: the set of resolved image-byte hashes is exactly
        // {AlicePng, BobPng}. The buggy whole-row bucket pulled Alice's bytes into Bob's cell → {AlicePng}
        // only (a single hash), which this set-equality catches.
        Assert.Equal(
            new HashSet<IrHash> { IrHash.Compute(AlicePng), IrHash.Compute(BobPng) },
            ResolvedImageHashes(accepted));

        // (3) reject ≡ base (both image inserts removed → no embeds).
        var rejected = RevisionProcessor.RejectRevisions(merged);
        Assert.Equal(Docs.StructuralBody(baseDoc), Docs.StructuralBody(rejected));
        Assert.Empty(BodyEmbedIds(rejected));
    }

    /// <summary>Open the doc and assert every <c>a:blip/@r:embed</c> in the main part body resolves to an
    /// <see cref="ImagePart"/> reachable from the main part by that relationship id.</summary>
    private static void AssertAllEmbedsResolve(WmlDocument d)
    {
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        var main = wDoc.MainDocumentPart;
        Assert.NotNull(main);

        var embeds = BodyEmbedIds(d);
        Assert.NotEmpty(embeds);   // the fixture always carries an image
        foreach (var embedId in embeds)
        {
            var rel = main!.GetPartById(embedId);
            Assert.IsType<ImagePart>(rel);
        }
    }

    /// <summary>Distinct <c>a:blip/@r:embed</c> ids in the main part body.</summary>
    private static List<string> BodyEmbedIds(WmlDocument d) =>
        XDocument.Parse(Docs.MainPartXml(d))
            .Descendants((XNamespace)ANs + "blip")
            .Select(b => b.Attribute((XNamespace)RNs + "embed")?.Value)
            .Where(v => !string.IsNullOrEmpty(v))
            .Select(v => v!)
            .Distinct()
            .ToList();

    /// <summary>The set of content hashes of the ImageParts that the body's embeds resolve to in the OUTPUT
    /// package (so a wrong-package import is caught by the bytes, not just by part existence).</summary>
    private static HashSet<IrHash> ResolvedImageHashes(WmlDocument d)
    {
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var wDoc = WordprocessingDocument.Open(ms, false);
        var main = wDoc.MainDocumentPart!;
        var hashes = new HashSet<IrHash>();
        foreach (var embedId in BodyEmbedIds(d))
        {
            var part = (ImagePart)main.GetPartById(embedId);
            using var s = part.GetStream(FileMode.Open, FileAccess.Read);
            using var mem = new MemoryStream();
            s.CopyTo(mem);
            hashes.Add(IrHash.Compute(mem.ToArray()));
        }
        return hashes;
    }

    // ------------------------------------------------------------------ helpers

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
