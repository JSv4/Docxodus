#nullable enable
using System.Linq;
using Docxodus;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Task 2.2 — the composite markup renderer's EqualBlock / InsertBlock / DeleteBlock /
/// single-reviewer-ModifyBlock paths. These fixtures use DISJOINT edits so each touched base block has
/// exactly one contributing reviewer (the single-modify path); the composed-paragraph branch
/// (<c>AuthoredTokens != null</c>) is Task 2.3 and is never exercised here. The invariant mirrors the
/// two-way renderer: reject-all yields the BASE document, accept-all yields each reviewer's accepted edits,
/// and every revision is attributed to its reviewer.
/// </summary>
public class IrCompositeMarkupRendererTests
{
    [Fact]
    public void Reject_all_equals_base_for_disjoint_two_reviewer_edit()
    {
        var baseDoc = Docs.Para("alpha one", "beta two", "gamma three");
        var r1 = Docs.Para("alpha one EDITED", "beta two", "gamma three");
        var r2 = Docs.Para("alpha one", "beta two", "gamma three EDITED");
        var script = IrCompositeMergerTests.MergeOf(baseDoc, ("Bob", r1), ("Fred", r2));
        var merged = IrCompositeMarkupRenderer.Render(script, baseDoc,
            new[] { ("Bob", r1), ("Fred", r2) }, new DocxDiffSettings().ToIrDiffSettings());

        Assert.Equal(Docs.PlainText(baseDoc), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
        var accepted = Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));
        Assert.Contains("alpha one EDITED", accepted);
        Assert.Contains("gamma three EDITED", accepted);
        var xml = Docs.MainPartXml(merged);
        Assert.Contains("w:author=\"Bob\"", xml);
        Assert.Contains("w:author=\"Fred\"", xml);
    }

    [Fact]
    public void Inserted_and_deleted_blocks_round_trip()
    {
        var baseDoc = Docs.Para("alpha", "beta", "gamma");
        var r1 = Docs.Para("alpha", "beta", "inserted by bob", "gamma");   // insert
        var r2 = Docs.Para("alpha", "gamma");                              // delete beta
        var script = IrCompositeMergerTests.MergeOf(baseDoc, ("Bob", r1), ("Fred", r2));
        var merged = IrCompositeMarkupRenderer.Render(script, baseDoc,
            new[] { ("Bob", r1), ("Fred", r2) }, new DocxDiffSettings().ToIrDiffSettings());
        Assert.Equal(Docs.PlainText(baseDoc), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
        var accepted = Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));
        Assert.Contains("inserted by bob", accepted);
        Assert.DoesNotContain("beta", accepted);
    }

    [Fact]
    public void All_equal_no_reviewer_edits_round_trips_to_base()
    {
        var baseDoc = Docs.Para("alpha", "beta", "gamma");
        var r1 = Docs.Para("alpha", "beta", "gamma");
        var script = IrCompositeMergerTests.MergeOf(baseDoc, ("Bob", r1));
        var merged = IrCompositeMarkupRenderer.Render(script, baseDoc,
            new[] { ("Bob", r1) }, new DocxDiffSettings().ToIrDiffSettings());

        Assert.Equal(Docs.PlainText(baseDoc), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
        Assert.Equal(Docs.PlainText(baseDoc), Docs.PlainText(RevisionAccepter.AcceptRevisions(merged)));
    }

    [Fact]
    public void Single_reviewer_modify_attributes_its_author()
    {
        var baseDoc = Docs.Para("alpha one", "beta two");
        var r1 = Docs.Para("alpha one EDITED", "beta two");
        var script = IrCompositeMergerTests.MergeOf(baseDoc, ("Bob", r1));
        var merged = IrCompositeMarkupRenderer.Render(script, baseDoc,
            new[] { ("Bob", r1) }, new DocxDiffSettings().ToIrDiffSettings());

        Assert.Equal(Docs.PlainText(baseDoc), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
        Assert.Contains("alpha one EDITED", Docs.PlainText(RevisionAccepter.AcceptRevisions(merged)));
        Assert.Contains("w:author=\"Bob\"", Docs.MainPartXml(merged));
    }
}
