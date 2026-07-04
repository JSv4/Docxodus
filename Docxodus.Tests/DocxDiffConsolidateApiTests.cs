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
    public void Consolidate_multireviewer_note_edit_composes_instead_of_failing_or_dropping()
    {
        // Was (v1): the merger did not build composite note-scope ops, so an N>=2 consolidate with a
        // reviewer note edit failed fast (NotSupportedException) rather than silently dropping it. The
        // N-way note merge now COMPOSES it: r1's body edit and r2's footnote edit both land on accept,
        // and reject restores the base note text.
        const string bodyA = "<w:p><w:r><w:t xml:space=\"preserve\">shared body</w:t></w:r></w:p>";
        const string bodyB = "<w:p><w:r><w:t xml:space=\"preserve\">shared body edited</w:t></w:r></w:p>";
        var b = Docxodus.Tests.Ir.IrTestDocuments.FromBodyXmlWithFootnote(bodyA, "original note");
        var r1 = Docxodus.Tests.Ir.IrTestDocuments.FromBodyXmlWithFootnote(bodyB, "original note");
        var r2 = Docxodus.Tests.Ir.IrTestDocuments.FromBodyXmlWithFootnote(bodyA, "edited note text");

        var merged = DocxDiff.Consolidate(b, new[]
        {
            new DocxDiffReviewer { Document = r1, Author = "Bob" },
            new DocxDiffReviewer { Document = r2, Author = "Fred" },
        });

        var accepted = RevisionAccepter.AcceptRevisions(merged);
        var rejected = RevisionProcessor.RejectRevisions(merged);

        Assert.Contains("shared body edited", Docxodus.Tests.Ir.Diff.Docs.PlainText(accepted));
        Assert.Contains("edited note text", FootnotePartText(accepted));
        Assert.DoesNotContain("original note", FootnotePartText(accepted));

        Assert.Equal(Docxodus.Tests.Ir.Diff.Docs.PlainText(b), Docxodus.Tests.Ir.Diff.Docs.PlainText(rejected));
        Assert.Contains("original note", FootnotePartText(rejected));
        Assert.DoesNotContain("edited note text", FootnotePartText(rejected));
    }

    /// <summary>The concatenated text of the footnotes part (all notes).</summary>
    private static string FootnotePartText(WmlDocument doc)
    {
        using var ms = new System.IO.MemoryStream(doc.DocumentByteArray);
        using var w = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(ms, false);
        var root = w.MainDocumentPart?.FootnotesPart?.GetXDocument().Root;
        if (root == null) return string.Empty;
        System.Xml.Linq.XNamespace ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        return string.Concat(root.Descendants(ns + "t").Select(t => t.Value));
    }

    [Fact]
    public void Consolidate_zero_reviewers_returns_base()
    {
        var b = Docxodus.Tests.Ir.Diff.Docs.Para("x");
        var merged = DocxDiff.Consolidate(b, System.Array.Empty<DocxDiffReviewer>());
        Assert.Equal(Docxodus.Tests.Ir.Diff.Docs.PlainText(b), Docxodus.Tests.Ir.Diff.Docs.PlainText(merged));
    }
}
