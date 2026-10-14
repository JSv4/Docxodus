#nullable enable
using System;
using System.Linq;
using Docxodus;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

public class CompositeInvariantTests
{
    [Fact]
    public void Deterministic_output_is_byte_identical()
    {
        var b = Docs.Para("alpha one", "beta two");
        var r1 = Docs.Para("alpha one EDITED", "beta two");
        var r2 = Docs.Para("alpha one", "beta two EDITED");
        var revs = new[] { new DocxDiffReviewer { Document = r1, Author = "Bob" },
                           new DocxDiffReviewer { Document = r2, Author = "Fred" } };
        var a = DocxDiff.Consolidate(b, revs).DocumentByteArray;
        var c = DocxDiff.Consolidate(b, revs).DocumentByteArray;
        Assert.Equal(a, c);
    }

    [Fact]
    public void Reviewer_order_independent_for_disjoint_edits()
    {
        var b = Docs.Para("alpha one", "beta two");
        var bob = new DocxDiffReviewer { Document = Docs.Para("alpha one EDITED", "beta two"), Author = "Bob" };
        var fred = new DocxDiffReviewer { Document = Docs.Para("alpha one", "beta two EDITED"), Author = "Fred" };
        var t1 = Docs.PlainText(RevisionAccepter.AcceptRevisions(DocxDiff.Consolidate(b, new[] { bob, fred })));
        var t2 = Docs.PlainText(RevisionAccepter.AcceptRevisions(DocxDiff.Consolidate(b, new[] { fred, bob })));
        Assert.Equal(t1, t2);
    }

    [Fact]
    public void Five_reviewers_each_distinct_edit_all_present()
    {
        var b = Docs.Para("a", "b", "c", "d", "e");
        var revs = Enumerable.Range(0, 5).Select(i =>
        {
            var paras = new[] { "a", "b", "c", "d", "e" };
            paras[i] += " EDITED" + i;
            return new DocxDiffReviewer { Document = Docs.Para(paras), Author = "R" + i };
        }).ToList();
        var accepted = Docs.PlainText(RevisionAccepter.AcceptRevisions(DocxDiff.Consolidate(b, revs)));
        for (int i = 0; i < 5; i++) Assert.Contains("EDITED" + i, accepted);
    }

    [Fact]
    public void Zero_reviewers_returns_base()
    {
        var b = Docs.Para("x");
        var merged = DocxDiff.Consolidate(b, Array.Empty<DocxDiffReviewer>());
        Assert.Equal(Docs.PlainText(b), Docs.PlainText(merged));
    }

    [Fact]
    public void One_reviewer_accept_equals_reviewer_text()
    {
        var b = Docs.Para("alpha one");
        var r = Docs.Para("alpha one EDITED");
        var cons = Docs.PlainText(RevisionAccepter.AcceptRevisions(
            DocxDiff.Consolidate(b, new[] { new DocxDiffReviewer { Document = r, Author = "Bob" } })));
        Assert.Equal(Docs.PlainText(r), cons);
    }

    [Fact]
    public void Reviewers_with_no_edits_produce_clean_base()
    {
        var b = Docs.Para("alpha one", "beta two");
        var noop1 = new DocxDiffReviewer { Document = Docs.Para("alpha one", "beta two"), Author = "Bob" };
        var noop2 = new DocxDiffReviewer { Document = Docs.Para("alpha one", "beta two"), Author = "Fred" };
        var merged = DocxDiff.Consolidate(b, new[] { noop1, noop2 });
        Assert.Equal(Docs.PlainText(b), Docs.PlainText(RevisionAccepter.AcceptRevisions(merged)));
        Assert.Equal(Docs.PlainText(b), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
    }
}
