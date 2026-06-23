#nullable enable
using System.Linq;
using Docxodus;
using Xunit;
namespace Docxodus.Tests;
public class DocxDiffConsolidateApiTests
{
    [Fact]
    public void Settings_default_policy_is_base_wins()
    {
        var s = new DocxDiffConsolidateSettings();
        Assert.Equal(ConflictResolution.BaseWins, s.ConflictResolution);
        Assert.NotNull(s.Diff);
    }

    [Fact]
    public void Reviewer_holds_document_and_author()
    {
        var r = new DocxDiffReviewer { Author = "Bob" };
        Assert.Equal("Bob", r.Author);
    }

    [Fact]
    public void Consolidate_two_reviewers_round_trips()
    {
        var b = Docxodus.Tests.Ir.Diff.Docs.Para("alpha one", "beta two", "gamma three");
        var r1 = Docxodus.Tests.Ir.Diff.Docs.Para("alpha one EDITED", "beta two", "gamma three");
        var r2 = Docxodus.Tests.Ir.Diff.Docs.Para("alpha one", "beta two", "gamma three EDITED");
        var merged = DocxDiff.Consolidate(b, new[]
        {
            new DocxDiffReviewer { Document = r1, Author = "Bob" },
            new DocxDiffReviewer { Document = r2, Author = "Fred" },
        });
        Assert.Equal(Docxodus.Tests.Ir.Diff.Docs.PlainText(b),
            Docxodus.Tests.Ir.Diff.Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
        var accepted = Docxodus.Tests.Ir.Diff.Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));
        Assert.Contains("alpha one EDITED", accepted);
        Assert.Contains("gamma three EDITED", accepted);
    }

    [Fact]
    public void GetConflicts_reports_overlapping_edit_with_competitors()
    {
        var b = Docxodus.Tests.Ir.Diff.Docs.Para("the quick brown fox");
        var r1 = Docxodus.Tests.Ir.Diff.Docs.Para("the SLOW brown fox");
        var r2 = Docxodus.Tests.Ir.Diff.Docs.Para("the FAST brown fox");
        var conflicts = DocxDiff.GetConflicts(b, new[]
        {
            new DocxDiffReviewer { Document = r1, Author = "Bob" },
            new DocxDiffReviewer { Document = r2, Author = "Fred" },
        });
        Assert.NotEmpty(conflicts);
        Assert.Contains(conflicts.SelectMany(c => c.Competitors), x => x.Author == "Bob");
        Assert.Contains(conflicts.SelectMany(c => c.Competitors), x => x.Author == "Fred");
    }

    [Fact]
    public void Consolidate_null_reviewer_element_throws_argument_exception_not_nre()
    {
        // Regression (engine audit): a null element inside a non-null reviewers list NRE'd before any guard.
        // The boundary must surface a clear ArgumentException, not a downstream NullReferenceException.
        var b = Docxodus.Tests.Ir.Diff.Docs.Para("x");
        Assert.Throws<System.ArgumentException>(() =>
            DocxDiff.Consolidate(b, new DocxDiffReviewer[] { null! }));
    }

    [Fact]
    public void Consolidate_null_diff_settings_throws_argument_exception_not_nre()
    {
        // Regression (engine audit): DocxDiffConsolidateSettings.Diff is a settable property; a null value
        // NRE'd at `s.Diff.ToIrDiffSettings()`. The boundary must surface a clear ArgumentException.
        var b = Docxodus.Tests.Ir.Diff.Docs.Para("x");
        var r = new DocxDiffReviewer { Document = Docxodus.Tests.Ir.Diff.Docs.Para("y"), Author = "Bob" };
        Assert.Throws<System.ArgumentException>(() =>
            DocxDiff.Consolidate(b, new[] { r }, new DocxDiffConsolidateSettings { Diff = null! }));
    }

    [Fact]
    public void Consolidate_multireviewer_note_edit_fails_loud_instead_of_silently_dropping_it()
    {
        // Regression (engine audit, v1 limitation made LOUD): the merger does not build composite note-scope
        // ops, so a reviewer's footnote/endnote CONTENT edit was silently dropped from the consolidated
        // document. For an N>=2 consolidate (no single-call fallback) it must fail fast
        // (NotSupportedException) rather than lose the edit silently. r1 edits only the body; r2 edits the
        // footnote — so the merge would silently drop r2's note edit.
        const string bodyA = "<w:p><w:r><w:t xml:space=\"preserve\">shared body</w:t></w:r></w:p>";
        const string bodyB = "<w:p><w:r><w:t xml:space=\"preserve\">shared body edited</w:t></w:r></w:p>";
        var b = Docxodus.Tests.Ir.IrTestDocuments.FromBodyXmlWithFootnote(bodyA, "original note");
        var r1 = Docxodus.Tests.Ir.IrTestDocuments.FromBodyXmlWithFootnote(bodyB, "original note");
        var r2 = Docxodus.Tests.Ir.IrTestDocuments.FromBodyXmlWithFootnote(bodyA, "edited note text");
        Assert.Throws<System.NotSupportedException>(() =>
            DocxDiff.Consolidate(b, new[]
            {
                new DocxDiffReviewer { Document = r1, Author = "Bob" },
                new DocxDiffReviewer { Document = r2, Author = "Fred" },
            }));
    }

    [Fact]
    public void Consolidate_zero_reviewers_returns_base()
    {
        var b = Docxodus.Tests.Ir.Diff.Docs.Para("x");
        var merged = DocxDiff.Consolidate(b, System.Array.Empty<DocxDiffReviewer>());
        Assert.Equal(Docxodus.Tests.Ir.Diff.Docs.PlainText(b), Docxodus.Tests.Ir.Diff.Docs.PlainText(merged));
    }
}
