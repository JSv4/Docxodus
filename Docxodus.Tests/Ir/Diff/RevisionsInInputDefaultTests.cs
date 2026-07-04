#nullable enable

using System.Linq;
using Docxodus;
using Xunit;
using static Docxodus.Tests.Ir.Diff.RevisionsInInputFixtures;

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// CHARACTERIZATION tests that PIN the engine's DEFAULT handling of inputs that already carry tracked changes
/// (deliverable 2 of the revisionsInInput campaign). They document the behavior as it ships TODAY — they are
/// expected to pass against the un-flagged engine — so any future drift is caught and the contract documented
/// in <c>docs/architecture/ir_diff_engine.md</c> stays honest.
///
/// <para><b>The pinned contract (rule N13 + carry-over).</b> The IR reader resolves revisions with
/// <c>RevisionView.Accept</c> BEFORE building, so the diff is computed over the ACCEPTED VIEW of each side: a
/// document's own <c>w:ins</c>/<c>w:del</c> never appear as their own diff, and the produced body carries only
/// THIS diff's revisions. BUT the markup renderer assembles the output on a clone of the LEFT package and only
/// rebuilds the body (+ changed notes), so pre-existing revision markup in CARRIED-OVER parts (headers/footers,
/// UNCHANGED footnotes/endnotes, styles, comments) survives verbatim into the result — the leak that
/// <see cref="DocxDiffSettings.PreAcceptInputRevisions"/> eliminates.</para>
/// </summary>
public class RevisionsInInputDefaultTests
{
    [Fact]
    public void Default_flattens_body_revisions_but_leaks_carried_over_scope_revisions()
    {
        // Body insertions/deletions are authored by "BodyPrior"; the header + footnote insertions (identical on
        // both sides, so an unchanged carry-over) are authored by "ScopePrior".
        var left = MultiScopeRevisionDoc(
            "Body alpha", "BodyPrior", "ins-a", "del-a", "ScopePrior", "hdr-prior", "fn-prior");
        var right = MultiScopeRevisionDoc(
            "Body gamma", "BodyPrior", "ins-g", "del-g", "ScopePrior", "hdr-prior", "fn-prior");

        var result = DocxDiff.Compare(left, right, new DocxDiffSettings { AuthorForRevisions = "TheDiff" });
        var authors = RevisionAuthorsAllScopes(result);

        // Body revisions are FLATTENED into the accepted view before diffing: "BodyPrior" never appears.
        Assert.DoesNotContain("BodyPrior", authors);
        // The diff's own revisions are present and attributed to the diff author.
        Assert.Contains("TheDiff", authors);
        // ...but the carried-over header/footnote revisions LEAK verbatim — "ScopePrior" survives. (Pinned
        // limitation of the default; fixed by PreAcceptInputRevisions.)
        Assert.Contains("ScopePrior", authors);
    }

    [Fact]
    public void Default_round_trips_the_body_to_the_accepted_view_of_each_side()
    {
        var left = MultiScopeRevisionDoc(
            "Body alpha", "BodyPrior", "ins-a", "del-a", "ScopePrior", "hdr-prior", "fn-prior");
        var right = MultiScopeRevisionDoc(
            "Body gamma", "BodyPrior", "ins-g", "del-g", "ScopePrior", "hdr-prior", "fn-prior");

        var result = DocxDiff.Compare(left, right, new DocxDiffSettings { AuthorForRevisions = "TheDiff" });

        // At the BODY level the default already round-trips to the ACCEPTED view of each side.
        Assert.Equal(
            Docs.PlainText(AcceptAll(left)),
            Docs.PlainText(RevisionProcessor.RejectRevisions(result)));
        Assert.Equal(
            Docs.PlainText(AcceptAll(right)),
            Docs.PlainText(RevisionAccepter.AcceptRevisions(result)));
    }

    [Fact]
    public void Default_leak_breaks_the_header_round_trip()
    {
        // The consequence of the leak: a leaked prior INSERTION in the header is REJECTED by reject-all, so the
        // reject view drops text the accepted view of the left actually contains — i.e. the round-trip does NOT
        // hold in the header scope under the default. (This is exactly why the flag exists.)
        var left = MultiScopeRevisionDoc(
            "Body alpha", "BodyPrior", "ins-a", "del-a", "ScopePrior", "hdr-prior", "fn-prior");
        var right = MultiScopeRevisionDoc(
            "Body gamma", "BodyPrior", "ins-g", "del-g", "ScopePrior", "hdr-prior", "fn-prior");

        var result = DocxDiff.Compare(left, right, new DocxDiffSettings { AuthorForRevisions = "TheDiff" });

        var acceptedLeftHeader = HeaderText(AcceptAll(left));                       // "Header hdr-prior"
        var rejectedResultHeader = HeaderText(RevisionProcessor.RejectRevisions(result)); // leaked ins stripped

        Assert.Contains("hdr-prior", acceptedLeftHeader);
        Assert.DoesNotContain("hdr-prior", rejectedResultHeader);
        Assert.NotEqual(acceptedLeftHeader, rejectedResultHeader); // round-trip broken in the header scope
    }

    [Fact]
    public void Default_GetRevisions_does_not_surface_the_inputs_own_pre_existing_revisions()
    {
        // Rule N13 at the consumer-revisions level: because the IR accepts both inputs before diffing, comparing
        // two redlines yields exactly the revisions of comparing their ACCEPTED views — the inputs' own
        // w:ins/w:del do not show up as their own diffs.
        var baseDoc = Docs.Para("The quick brown fox jumps over the lazy dog.");
        var redlineA = DocxDiff.Compare(baseDoc, Docs.Para("The quick red fox jumps over the lazy dog."),
            new DocxDiffSettings { AuthorForRevisions = "Alice" });
        var redlineB = DocxDiff.Compare(baseDoc, Docs.Para("The quick brown fox leaps over the sleepy dog."),
            new DocxDiffSettings { AuthorForRevisions = "Bob" });

        var overRedlines = DocxDiff.GetRevisions(redlineA, redlineB,
            new DocxDiffSettings { AuthorForRevisions = "D" });
        var overAcceptedViews = DocxDiff.GetRevisions(AcceptAll(redlineA), AcceptAll(redlineB),
            new DocxDiffSettings { AuthorForRevisions = "D" });

        Assert.Equal(
            overAcceptedViews.Select(r => $"{r.Type}:{r.Text}").ToList(),
            overRedlines.Select(r => $"{r.Type}:{r.Text}").ToList());
    }

    [Fact]
    public void PreAcceptInputRevisions_does_not_flatten_a_revision_inside_a_comment_definition()
    {
        // The flag's mechanism is RevisionProcessor.AcceptRevisions, which processes body/headers/footers/
        // notes/styles but NOT the comments part — so a tracked change inside a comment DEFINITION is the one
        // scope the flag does not clean. Pinned here so the boundary is specified, not silent.
        var left = CommentWithRevisionDoc("Body one", "CmtReviewer");
        var right = CommentWithRevisionDoc("Body two", "CmtReviewer");

        var flagged = DocxDiff.Compare(left, right,
            new DocxDiffSettings { PreAcceptInputRevisions = true, AuthorForRevisions = "TheDiff" });

        // The comment-internal insertion's author SURVIVES even with the flag on (documented boundary).
        Assert.Contains("CmtReviewer", RevisionAuthorsAllScopes(flagged));
    }

    [Fact]
    public void PreAcceptInputRevisions_does_not_flatten_a_revision_inside_a_glossary_docpart()
    {
        // The SECOND uncovered scope: RevisionProcessor.AcceptRevisions never processes the GlossaryDocumentPart
        // (building blocks / AutoText), and the markup renderer clones the whole left package (glossary included)
        // verbatim — so a tracked change inside a glossary entry survives the flag. Pinned so the boundary is
        // specified, not silent.
        var left = GlossaryWithRevisionDoc("Body one", "GlossReviewer");
        var right = GlossaryWithRevisionDoc("Body two", "GlossReviewer");

        var flagged = DocxDiff.Compare(left, right,
            new DocxDiffSettings { PreAcceptInputRevisions = true, AuthorForRevisions = "TheDiff" });

        Assert.Contains("GlossReviewer", RevisionAuthorsAllScopes(flagged));
    }
}
