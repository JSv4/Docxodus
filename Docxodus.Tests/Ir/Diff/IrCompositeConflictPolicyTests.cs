#nullable enable
using Docxodus;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// Task 2.4 — the three conflict-resolution policies, proven end-to-end through the
/// composite markup renderer. Scenario: a single overlapping word replacement —
/// base "the quick brown fox", Bob → "the SLOW brown fox", Fred → "the FAST brown fox"
/// (both edit word 2 differently). The merger (<see cref="IrCompositeMerger"/>) owns the
/// policy and emits the authored op stream; the renderer faithfully renders it. The
/// per-policy contract is asserted against the ACCEPTED body, and the universal invariant
/// (reject ≡ base, regardless of policy) is asserted across all three.
/// </summary>
public class IrCompositeConflictPolicyTests
{
    private static WmlDocument Render(ConflictResolution policy)
    {
        var baseDoc = Docs.Para("the quick brown fox");
        var r1 = Docs.Para("the SLOW brown fox");
        var r2 = Docs.Para("the FAST brown fox");
        var script = IrCompositeMergerTests.MergeOf(policy, baseDoc, ("Bob", r1), ("Fred", r2));
        return IrCompositeMarkupRenderer.Render(script, baseDoc, new[] { ("Bob", r1), ("Fred", r2) },
            new DocxDiffSettings().ToIrDiffSettings());
    }

    private static string Accepted(ConflictResolution p) =>
        Docs.PlainText(RevisionAccepter.AcceptRevisions(Render(p)));

    private static string Rejected(ConflictResolution p) =>
        Docs.PlainText(RevisionProcessor.RejectRevisions(Render(p)));

    [Fact]
    public void BaseWins_keeps_base_word()
    {
        var accepted = Accepted(ConflictResolution.BaseWins);
        Assert.Contains("quick", accepted);
        Assert.DoesNotContain("SLOW", accepted);
        Assert.DoesNotContain("FAST", accepted);
    }

    [Fact]
    public void FirstReviewerWins_applies_first_only()
    {
        var accepted = Accepted(ConflictResolution.FirstReviewerWins);
        Assert.Contains("SLOW", accepted);
        Assert.DoesNotContain("FAST", accepted);
        Assert.DoesNotContain("quick", accepted);
    }

    [Fact]
    public void StackAll_includes_both()
    {
        var accepted = Accepted(ConflictResolution.StackAll);
        Assert.Contains("SLOW", accepted);
        Assert.Contains("FAST", accepted);
    }

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void All_policies_reject_to_base(ConflictResolution policy)
    {
        Assert.Equal("the quick brown fox", Rejected(policy));
    }

    [Theory]
    [InlineData(ConflictResolution.BaseWins)]
    [InlineData(ConflictResolution.FirstReviewerWins)]
    [InlineData(ConflictResolution.StackAll)]
    public void All_policies_record_the_conflict(ConflictResolution policy)
    {
        var baseDoc = Docs.Para("the quick brown fox");
        var r1 = Docs.Para("the SLOW brown fox");
        var r2 = Docs.Para("the FAST brown fox");
        var script = IrCompositeMergerTests.MergeOf(policy, baseDoc, ("Bob", r1), ("Fred", r2));
        Assert.NotEmpty(script.Conflicts);
    }
}
