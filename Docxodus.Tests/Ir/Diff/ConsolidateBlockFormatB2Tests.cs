#nullable enable

using System.Collections.Generic;
using System.Linq;
using Docxodus;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Consolidate B2 — N-way merge of the table-shell (<c>w:tcPrChange</c>/<c>w:trPrChange</c>/<c>w:tblPrChange</c>/
/// <c>w:tblGridChange</c>/<c>w:tblPrExChange</c>) and section (<c>w:sectPrChange</c>) block-format families,
/// plus the text+format safety rule. Before B2 these edits were the pinned Consolidate ceiling: a reviewer's
/// shell/section-only edit was ignored (silently dropped from accept) or, for a cell shell, swapped in with no
/// marker (so <c>reject ≠ base</c>). These tests assert the NEW behavior at the PROPERTY-BYTE level via
/// <see cref="Docs.ShellSection"/> — the format-blind text projections (<see cref="Docs.PlainText"/>/
/// <see cref="Docs.StructuralBody"/>) cannot see a lost shell.
/// </summary>
public class ConsolidateBlockFormatB2Tests
{
    // ------------------------------------------------------------------ fixtures

    // A 1-row, 2-cell table with explicit tblPr/tblGrid/trPr/tcPr shells, framed by a lead paragraph and a
    // tail paragraph + trailing sectPr, so every block-format family has a distinct, mutable shell.
    private static string Body(
        string tblW = "5000", string trHeight = "300", string gridCol0 = "2500",
        string tcW00 = "2500", string tblPrEx = "", string pgMarTop = "1440") =>
        "<w:p><w:r><w:t>lead</w:t></w:r></w:p>" +
        "<w:tbl>" +
            $"<w:tblPr><w:tblW w:w=\"{tblW}\" w:type=\"dxa\"/></w:tblPr>" +
            $"<w:tblGrid><w:gridCol w:w=\"{gridCol0}\"/><w:gridCol w:w=\"2500\"/></w:tblGrid>" +
            "<w:tr>" +
                tblPrEx +
                $"<w:trPr><w:trHeight w:val=\"{trHeight}\"/></w:trPr>" +
                $"<w:tc><w:tcPr><w:tcW w:w=\"{tcW00}\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>" +
                "<w:tc><w:tcPr><w:tcW w:w=\"2500\" w:type=\"dxa\"/></w:tcPr><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>" +
            "</w:tr>" +
        "</w:tbl>" +
        "<w:p><w:r><w:t>tail</w:t></w:r></w:p>" +
        $"<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"{pgMarTop}\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"1440\"/></w:sectPr>";

    private static WmlDocument Base() => IrTestDocuments.FromBodyXml(Body());

    // ------------------------------------------------------------------ consolidate helpers

    private static WmlDocument Consolidate(WmlDocument baseDoc, ConflictResolution policy,
        params (string Author, WmlDocument Doc)[] reviewers)
        => DocxDiff.Consolidate(baseDoc,
            reviewers.Select(r => new DocxDiffReviewer { Author = r.Author, Document = r.Doc }).ToList(),
            new DocxDiffConsolidateSettings { ConflictResolution = policy });

    private static IReadOnlyList<DocxDiffConflict> Conflicts(WmlDocument baseDoc, ConflictResolution policy,
        params (string Author, WmlDocument Doc)[] reviewers)
        => DocxDiff.GetConflicts(baseDoc,
            reviewers.Select(r => new DocxDiffReviewer { Author = r.Author, Document = r.Doc }).ToList(),
            new DocxDiffConsolidateSettings { ConflictResolution = policy });

    private static WmlDocument Accept(WmlDocument merged) => RevisionAccepter.AcceptRevisions(merged);
    private static WmlDocument Reject(WmlDocument merged) => RevisionProcessor.RejectRevisions(merged);
    private static string Xml(WmlDocument d) => Docs.MainPartXml(d);

    /// <summary>The core B2 property-byte round-trip for a SINGLE reviewer who changed exactly one shell/section:
    /// accept ≡ the reviewer (winner), reject ≡ base, and the native change marker is present.</summary>
    private static void AssertSingleReviewerMerge(WmlDocument baseDoc, WmlDocument reviewer, string changeMarker)
    {
        var merged = Consolidate(baseDoc, ConflictResolution.BaseWins, ("Alice", reviewer));

        // Native markup emitted.
        Assert.Contains(changeMarker, Xml(merged));
        // accept ≡ winner and reject ≡ base at the property-byte level (shells + section).
        Assert.Equal(Docs.ShellSection(reviewer), Docs.ShellSection(Accept(merged)));
        Assert.Equal(Docs.ShellSection(baseDoc), Docs.ShellSection(Reject(merged)));
        // Text is untouched by a shell/section-only edit.
        Assert.Equal(Docs.StructuralBody(baseDoc), Docs.StructuralBody(Reject(merged)));
    }

    // ------------------------------------------------------------------ Phase 1: table-shell single-reviewer merge

    [Fact]
    public void Cell_tcPr_only_edit_merges_with_marker_and_round_trips()
        => AssertSingleReviewerMerge(Base(), IrTestDocuments.FromBodyXml(Body(tcW00: "3000")), "w:tcPrChange");

    [Fact]
    public void Row_trPr_only_edit_merges_with_marker_and_round_trips()
        => AssertSingleReviewerMerge(Base(), IrTestDocuments.FromBodyXml(Body(trHeight: "500")), "w:trPrChange");

    [Fact]
    public void Table_tblPr_only_edit_merges_with_marker_and_round_trips()
        => AssertSingleReviewerMerge(Base(), IrTestDocuments.FromBodyXml(Body(tblW: "6000")), "w:tblPrChange");

    [Fact]
    public void Table_tblGrid_only_edit_merges_with_marker_and_round_trips()
        => AssertSingleReviewerMerge(Base(), IrTestDocuments.FromBodyXml(Body(gridCol0: "3000")), "w:tblGridChange");

    [Fact]
    public void Row_tblPrEx_only_edit_merges_with_marker_and_round_trips()
        => AssertSingleReviewerMerge(
            Base(),
            IrTestDocuments.FromBodyXml(Body(tblPrEx: "<w:tblPrEx><w:tblCellMar><w:left w:w=\"120\" w:type=\"dxa\"/></w:tblCellMar></w:tblPrEx>")),
            "w:tblPrExChange");

    // ------------------------------------------------------------------ Phase 2: section single-reviewer merge

    [Fact]
    public void Trailing_sectPr_only_edit_merges_with_marker_and_round_trips()
        => AssertSingleReviewerMerge(Base(), IrTestDocuments.FromBodyXml(Body(pgMarTop: "2880")), "w:sectPrChange");

    [Fact]
    public void Trailing_sectPr_change_reports_section_consolidated_revision()
    {
        var revs = DocxDiff.GetConsolidatedRevisions(Base(),
            new[] { new DocxDiffReviewer { Author = "Alice", Document = IrTestDocuments.FromBodyXml(Body(pgMarTop: "2880")) } });
        Assert.Contains(revs, r => r.FormatChange is { } fc && fc.Scope == DocxDiffFormatChangeScope.Section && r.Author == "Alice");
    }

    [Fact]
    public void Inline_sectPr_change_merges_with_marker_and_round_trips()
    {
        // A mid-document inline section break (w:pPr/w:sectPr) whose page margin changes rides B1's paragraph
        // FormatOnly path (BlockSignature includes the section key under the section slice).
        static string InlineBody(string top) =>
            $"<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"{top}\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"1440\"/></w:sectPr></w:pPr><w:r><w:t>a</w:t></w:r></w:p>" +
            "<w:p><w:r><w:t>b</w:t></w:r></w:p>" +
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>";
        var baseDoc = IrTestDocuments.FromBodyXml(InlineBody("1440"));
        var alice = IrTestDocuments.FromBodyXml(InlineBody("2880"));

        var merged = Consolidate(baseDoc, ConflictResolution.BaseWins, ("Alice", alice));
        Assert.Contains("w:sectPrChange", Xml(merged));
        Assert.Equal(Docs.ShellSection(alice), Docs.ShellSection(Accept(merged)));
        Assert.Equal(Docs.ShellSection(baseDoc), Docs.ShellSection(Reject(merged)));
    }

    // ------------------------------------------------------------------ Phase 1/2: consensus + conflict per family

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Agreeing_cell_shell_edits_reach_consensus_no_conflict(ConflictResolution policy)
    {
        var baseDoc = Base();
        var v = IrTestDocuments.FromBodyXml(Body(tcW00: "3000"));
        Assert.Empty(Conflicts(baseDoc, policy, ("Alice", v), ("Bob", v)));

        var merged = Consolidate(baseDoc, policy, ("Alice", v), ("Bob", v));
        Assert.Equal(Docs.ShellSection(v), Docs.ShellSection(Accept(merged)));
        Assert.Equal(Docs.ShellSection(baseDoc), Docs.ShellSection(Reject(merged)));
    }

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Competing_cell_shell_edits_record_conflict_resolved_per_policy(ConflictResolution policy)
    {
        var baseDoc = Base();
        var alice = IrTestDocuments.FromBodyXml(Body(tcW00: "3000"));
        var bob = IrTestDocuments.FromBodyXml(Body(tcW00: "4000"));

        var conflicts = Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        Assert.NotEmpty(conflicts);
        var authors = conflicts.SelectMany(c => c.Competitors.Select(x => x.Author)).Distinct().ToList();
        Assert.Contains("Alice", authors);
        Assert.Contains("Bob", authors);

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var winner = policy == ConflictResolution.BaseWins ? baseDoc : alice;
        Assert.Equal(Docs.ShellSection(winner), Docs.ShellSection(Accept(merged)));
        // reject restores base shells regardless of policy.
        Assert.Equal(Docs.ShellSection(baseDoc), Docs.ShellSection(Reject(merged)));
    }

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Competing_trailing_sectPr_edits_record_conflict_resolved_per_policy(ConflictResolution policy)
    {
        var baseDoc = Base();
        var alice = IrTestDocuments.FromBodyXml(Body(pgMarTop: "2880"));
        var bob = IrTestDocuments.FromBodyXml(Body(pgMarTop: "720"));

        Assert.NotEmpty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var winner = policy == ConflictResolution.BaseWins ? baseDoc : alice;
        Assert.Equal(Docs.ShellSection(winner), Docs.ShellSection(Accept(merged)));
        Assert.Equal(Docs.ShellSection(baseDoc), Docs.ShellSection(Reject(merged)));
    }

    [Fact]
    public void Disjoint_families_compose_cell_shell_and_section_both_land()
    {
        var baseDoc = Base();
        var alice = IrTestDocuments.FromBodyXml(Body(tcW00: "3000"));   // cell shell only
        var bob = IrTestDocuments.FromBodyXml(Body(pgMarTop: "2880"));  // section only

        Assert.Empty(Conflicts(baseDoc, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob));
        // The composite carries Alice's cell shell AND Bob's section.
        var expected = IrTestDocuments.FromBodyXml(Body(tcW00: "3000", pgMarTop: "2880"));
        Assert.Equal(Docs.ShellSection(expected), Docs.ShellSection(Accept(merged)));
        Assert.Equal(Docs.ShellSection(baseDoc), Docs.ShellSection(Reject(merged)));
    }

    // ------------------------------------------------------------------ Phase 1c: multi-reviewer row/table-level shells

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void Agreeing_row_shell_edits_reach_consensus_no_conflict(ConflictResolution policy)
    {
        var baseDoc = Base();
        var v = IrTestDocuments.FromBodyXml(Body(trHeight: "500"));
        Assert.Empty(Conflicts(baseDoc, policy, ("Alice", v), ("Bob", v)));

        var merged = Consolidate(baseDoc, policy, ("Alice", v), ("Bob", v));
        Assert.Equal(Docs.ShellSection(v), Docs.ShellSection(Accept(merged)));
        Assert.Equal(Docs.ShellSection(baseDoc), Docs.ShellSection(Reject(merged)));
    }

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    public void Competing_row_shell_edits_record_conflict(ConflictResolution policy)
    {
        var baseDoc = Base();
        var alice = IrTestDocuments.FromBodyXml(Body(trHeight: "500"));
        var bob = IrTestDocuments.FromBodyXml(Body(trHeight: "700"));

        Assert.NotEmpty(Conflicts(baseDoc, policy, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, policy, ("Alice", alice), ("Bob", bob));
        var winner = policy == ConflictResolution.BaseWins ? baseDoc : alice;
        Assert.Equal(Docs.ShellSection(winner), Docs.ShellSection(Accept(merged)));
        Assert.Equal(Docs.ShellSection(baseDoc), Docs.ShellSection(Reject(merged)));
    }

    [Fact]
    public void Disjoint_table_and_row_shells_compose_both_land()
    {
        var baseDoc = Base();
        var alice = IrTestDocuments.FromBodyXml(Body(tblW: "6000"));    // table-level shell only
        var bob = IrTestDocuments.FromBodyXml(Body(trHeight: "500"));   // row-level shell only

        Assert.Empty(Conflicts(baseDoc, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob));
        var expected = IrTestDocuments.FromBodyXml(Body(tblW: "6000", trHeight: "500"));
        Assert.Equal(Docs.ShellSection(expected), Docs.ShellSection(Accept(merged)));
        Assert.Equal(Docs.ShellSection(baseDoc), Docs.ShellSection(Reject(merged)));
    }

    [Fact]
    public void Cell_shell_and_row_shell_by_different_reviewers_compose()
    {
        var baseDoc = Base();
        var alice = IrTestDocuments.FromBodyXml(Body(tcW00: "3000"));   // cell shell (ModifyBlock table)
        var bob = IrTestDocuments.FromBodyXml(Body(trHeight: "500"));   // row shell (FormatOnly table)

        Assert.Empty(Conflicts(baseDoc, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob)));

        var merged = Consolidate(baseDoc, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob));
        var expected = IrTestDocuments.FromBodyXml(Body(tcW00: "3000", trHeight: "500"));
        Assert.Equal(Docs.ShellSection(expected), Docs.ShellSection(Accept(merged)));
        Assert.Equal(Docs.ShellSection(baseDoc), Docs.ShellSection(Reject(merged)));
    }

    // ------------------------------------------------------------------ Phase 3: text+format is never silently dropped

    /// <summary>
    /// A reviewer edits a paragraph's TEXT and its FORMAT (pPr). v1 decision: this is conflict-routed, never a
    /// silent format drop. When another reviewer edits a DISJOINT text span of the same paragraph, the two texts
    /// do not compose past the format change — the format-carrying block is a RECORDED conflict.
    /// </summary>
    [Fact]
    public void Text_plus_pPr_edit_is_recorded_conflict_never_silent_drop()
    {
        var baseDoc = IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:jc w:val=\"left\"/></w:pPr><w:r><w:t xml:space=\"preserve\">The cat sat. The dog ran.</w:t></w:r></w:p>");
        // Alice edits sentence 1 AND changes alignment left→center.
        var alice = IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t xml:space=\"preserve\">The CAT sat. The dog ran.</w:t></w:r></w:p>");
        // Bob edits sentence 2 only (no format change).
        var bob = IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:jc w:val=\"left\"/></w:pPr><w:r><w:t xml:space=\"preserve\">The cat sat. The dog RAN.</w:t></w:r></w:p>");

        var conflicts = Conflicts(baseDoc, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob));
        Assert.NotEmpty(conflicts);  // Alice's format+text edit must surface as a conflict, not vanish.

        // reject restores the base paragraph exactly (text + pPr).
        var merged = Consolidate(baseDoc, ConflictResolution.BaseWins, ("Alice", alice), ("Bob", bob));
        Assert.Equal(Docs.PlainText(baseDoc), Docs.PlainText(Reject(merged)));
        Assert.Equal(Docs.ShellSection(baseDoc), Docs.ShellSection(Reject(merged)));
    }

    /// <summary>
    /// The positive case: a SINGLE reviewer editing a paragraph's text AND its pPr tracks BOTH — the text as
    /// w:ins/w:del and the pPr as w:pPrChange (no conflict, nothing dropped). Only a CROSS-reviewer text+pPr
    /// collision conflict-routes (v1 decision; true inline text+format compose is deferred to B3).
    /// </summary>
    [Fact]
    public void Single_reviewer_text_and_pPr_edit_tracks_both_no_conflict()
    {
        var baseDoc = IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:jc w:val=\"left\"/></w:pPr><w:r><w:t xml:space=\"preserve\">The cat sat.</w:t></w:r></w:p>");
        var alice = IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t xml:space=\"preserve\">The CAT sat.</w:t></w:r></w:p>");

        Assert.Empty(Conflicts(baseDoc, ConflictResolution.BaseWins, ("Alice", alice)));
        var merged = Consolidate(baseDoc, ConflictResolution.BaseWins, ("Alice", alice));
        Assert.Contains("w:pPrChange", Xml(merged));
        Assert.Equal(Docs.PlainText(alice), Docs.PlainText(Accept(merged)));       // text edit lands
        Assert.Equal(Docs.ShellSection(alice), Docs.ShellSection(Accept(merged))); // pPr change lands
        Assert.Equal(Docs.PlainText(baseDoc), Docs.PlainText(Reject(merged)));     // reject ≡ base text
        Assert.Equal(Docs.ShellSection(baseDoc), Docs.ShellSection(Reject(merged)));// reject ≡ base pPr
    }
}
