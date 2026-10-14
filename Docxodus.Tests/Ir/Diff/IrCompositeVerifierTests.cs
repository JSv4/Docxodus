#nullable enable

using Docxodus;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Exercises <see cref="IrCompositeVerifier"/> (T5.1): independently reconstruct the policy-resolved
/// accepted body from the composite script's op semantics and assert it text-equals what the renderer
/// produced (<c>PlainText(AcceptRevisions(DocxDiff.Consolidate(...)))</c>). A mismatch is a real bug —
/// the script ops would not faithfully describe the rendered output.
/// </summary>
public class IrCompositeVerifierTests
{
    [Fact]
    public void Composite_apply_reconstructs_accepted_body_disjoint()
    {
        var b = Docs.Para("the quick brown fox", "second line here");
        var r1 = Docs.Para("the SLOW brown fox", "second line here");
        var r2 = Docs.Para("the quick brown fox", "second line HERE now");
        var merged = DocxDiff.Consolidate(b, new[]{
            new DocxDiffReviewer{Document=r1,Author="Bob"}, new DocxDiffReviewer{Document=r2,Author="Fred"}});
        var acceptedText = Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));
        IrCompositeVerifier.Verify(b, new[]{("Bob",r1),("Fred",r2)}, ConflictResolution.BaseWins, acceptedText);
    }

    [Fact]
    public void Composite_apply_reconstructs_accepted_body_for_all_policies_on_conflict()
    {
        var b = Docs.Para("the quick brown fox");
        var r1 = Docs.Para("the SLOW brown fox");
        var r2 = Docs.Para("the FAST brown fox");
        foreach (var p in new[]{ConflictResolution.BaseWins, ConflictResolution.FirstReviewerWins, ConflictResolution.StackAll})
        {
            var merged = DocxDiff.Consolidate(b, new[]{
                new DocxDiffReviewer{Document=r1,Author="Bob"}, new DocxDiffReviewer{Document=r2,Author="Fred"}},
                new DocxDiffConsolidateSettings{ConflictResolution=p});
            var acceptedText = Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));
            IrCompositeVerifier.Verify(b, new[]{("Bob",r1),("Fred",r2)}, p, acceptedText);
        }
    }

    [Fact]
    public void Composite_apply_reconstructs_with_insert_and_delete_blocks()
    {
        var b = Docs.Para("alpha", "beta", "gamma");
        var r1 = Docs.Para("alpha", "beta", "inserted line", "gamma");
        var r2 = Docs.Para("alpha", "gamma");
        var merged = DocxDiff.Consolidate(b, new[]{
            new DocxDiffReviewer{Document=r1,Author="Bob"}, new DocxDiffReviewer{Document=r2,Author="Fred"}});
        var acceptedText = Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));
        IrCompositeVerifier.Verify(b, new[]{("Bob",r1),("Fred",r2)}, ConflictResolution.BaseWins, acceptedText);
    }
}
