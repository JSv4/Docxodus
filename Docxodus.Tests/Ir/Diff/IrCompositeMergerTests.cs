#nullable enable
using System.Linq;
using Docxodus;
using Docxodus.Ir;
using Docxodus.Ir.Diff;
using Xunit;

namespace Docxodus.Tests.Ir.Diff;

public class IrCompositeMergerTests
{
    internal static readonly IrReaderOptions ReadOpts =
        new() { RetainSources = false, RevisionView = RevisionView.Accept };

    [Fact]
    public void GroupByBaseAnchor_colocates_per_reviewer_ops()
    {
        var baseDoc = Docs.Para("alpha one", "beta two");
        var r1 = Docs.Para("alpha one EDITED", "beta two");
        var diff = new DocxDiffSettings().ToIrDiffSettings();
        var baseIr = IrReader.Read(baseDoc, ReadOpts);
        var s1 = IrEditScriptBuilder.Build(baseIr, IrReader.Read(r1, ReadOpts), diff);
        var grouped = IrCompositeMerger.GroupByBaseAnchor(new[] { s1 });
        Assert.Contains(grouped.Values, list => list.Any(x => x.Op.Kind == IrEditOpKind.ModifyBlock));
    }

    [Fact]
    public void Disjoint_block_edits_from_two_reviewers_both_appear_attributed()
    {
        var b = Docs.Para("alpha one", "beta two", "gamma three");
        var r1 = Docs.Para("alpha one EDITED", "beta two", "gamma three");
        var r2 = Docs.Para("alpha one", "beta two", "gamma three EDITED");
        var s = MergeOf(b, ("Bob", r1), ("Fred", r2));
        var mods = s.Operations.Where(o => o.Op.Kind == IrEditOpKind.ModifyBlock).ToList();
        Assert.Equal(2, mods.Count);
        Assert.Contains(mods, m => m.Author == "Bob" || (m.AuthoredTokens?.Any(a => a.Author == "Bob") ?? false));
        Assert.Contains(mods, m => m.Author == "Fred" || (m.AuthoredTokens?.Any(a => a.Author == "Fred") ?? false));
        Assert.Empty(s.Conflicts);
    }

    [Fact]
    public void Identical_edit_by_both_reviewers_is_consensus_not_conflict()
    {
        var b = Docs.Para("alpha one", "beta two");
        var same = Docs.Para("alpha one EDITED", "beta two");
        var s = MergeOf(b, ("Bob", same), ("Fred", same));
        Assert.Empty(s.Conflicts);
        Assert.Single(s.Operations.Where(o => o.Op.Kind == IrEditOpKind.ModifyBlock));
    }

    [Fact]
    public void Two_reviewers_editing_same_words_differently_conflict()
    {
        var b = Docs.Para("the quick brown fox");
        var r1 = Docs.Para("the SLOW brown fox");
        var r2 = Docs.Para("the FAST brown fox");
        Assert.NotEmpty(MergeOf(b, ("Bob", r1), ("Fred", r2)).Conflicts);
    }

    [Fact]
    public void Two_reviewers_editing_different_words_of_one_paragraph_compose_without_conflict()
    {
        var b = Docs.Para("the quick brown fox jumps");
        var r1 = Docs.Para("the SLOW brown fox jumps");
        var r2 = Docs.Para("the quick brown fox LEAPS");
        var s = MergeOf(b, ("Bob", r1), ("Fred", r2));
        Assert.Empty(s.Conflicts);
        var mod = Assert.Single(s.Operations.Where(o => o.Op.Kind == IrEditOpKind.ModifyBlock));
        Assert.NotNull(mod.AuthoredTokens);
        Assert.Contains(mod.AuthoredTokens!, a => a.Author == "Bob");
        Assert.Contains(mod.AuthoredTokens!, a => a.Author == "Fred");
    }

    [Fact]
    public void Both_reviewers_inserting_paragraphs_both_appear_no_conflict()
    {
        var b = Docs.Para("alpha", "omega");
        var r1 = Docs.Para("alpha", "bob inserted line", "omega");
        var r2 = Docs.Para("alpha", "fred inserted line", "omega");
        var s = MergeOf(b, ("Bob", r1), ("Fred", r2));
        Assert.Equal(2, s.Operations.Count(o => o.Op.Kind == IrEditOpKind.InsertBlock));
        Assert.Empty(s.Conflicts);
    }

    [Fact]
    public void Delete_by_one_modify_by_other_is_block_conflict()
    {
        var b = Docs.Para("alpha", "beta two three", "omega");
        var deleter = Docs.Para("alpha", "omega");
        var editor = Docs.Para("alpha", "beta two three EDITED", "omega");
        Assert.NotEmpty(MergeOf(b, ("Bob", deleter), ("Fred", editor)).Conflicts);
    }

    [Fact]
    public void Both_reviewers_deleting_same_block_is_consensus_not_conflict()
    {
        var b = Docs.Para("alpha", "beta to delete", "omega");
        var r1 = Docs.Para("alpha", "omega");
        var r2 = Docs.Para("alpha", "omega");
        var s = MergeOf(b, ("Bob", r1), ("Fred", r2));
        Assert.Empty(s.Conflicts);
        Assert.Single(s.Operations.Where(o => o.Op.Kind == IrEditOpKind.DeleteBlock));
    }

    [Fact]
    public void Reviewer_changing_both_text_and_pPr_routes_to_conflict_not_silent_compose()
    {
        // Base paragraph (no pPr) plus a second paragraph so the doc has stable structure.
        var b = ParaDoc(("the quick brown fox jumps", ""), ("tail", ""));
        // Bob changes a word AND the paragraph's alignment (center) — a pPr delta.
        var r1 = ParaDoc(("the SLOW brown fox jumps", "<w:jc w:val=\"center\"/>"), ("tail", ""));
        // Fred changes a DIFFERENT word, no pPr change — on its own this would token-compose with Bob.
        var r2 = ParaDoc(("the quick brown fox LEAPS", ""), ("tail", ""));

        var s = MergeOf(b, ("Bob", r1), ("Fred", r2));

        // The reviewer's pPr change must NOT be silently dropped by the token-compose path.
        // It must surface: either as a recorded conflict OR as a non-composed (no AuthoredTokens) op.
        var firstParaMods = s.Operations
            .Where(o => o.Op.Kind == IrEditOpKind.ModifyBlock)
            .ToList();
        bool composedAway = firstParaMods.Any(o => o.AuthoredTokens != null);
        Assert.True(!composedAway || s.Conflicts.Count > 0,
            "A reviewer who changed BOTH text and pPr was silently token-composed, dropping the pPr change.");
        // The block-level conflict path is the expected route for a text+pPr ModifyBlock.
        Assert.NotEmpty(s.Conflicts);
    }

    /// <summary>
    /// Build a one-section DOCX from (text, pPrInnerXml) pairs: each tuple is one paragraph whose
    /// <c>w:pPr</c> inner XML is <c>pPrInnerXml</c> (empty string ⇒ no <c>w:pPr</c>) and whose single
    /// run holds <c>text</c>. Lets a test express a paragraph-level formatting delta (e.g. alignment).
    /// </summary>
    private static WmlDocument ParaDoc(params (string Text, string PPrInnerXml)[] paras)
    {
        var sb = new System.Text.StringBuilder();
        foreach (var (text, pPr) in paras)
        {
            sb.Append("<w:p>");
            if (!string.IsNullOrEmpty(pPr))
                sb.Append("<w:pPr>").Append(pPr).Append("</w:pPr>");
            sb.Append("<w:r><w:t xml:space=\"preserve\">").Append(text).Append("</w:t></w:r></w:p>");
        }
        return Docxodus.Tests.Ir.IrTestDocuments.FromBodyXml(sb.ToString());
    }

    // Helper reused by later tasks: merge base + reviewers into an IrCompositeScript.
    internal static IrCompositeScript MergeOf(WmlDocument baseDoc, params (string Author, WmlDocument Doc)[] reviewers)
        => MergeOf(ConflictResolution.BaseWins, baseDoc, reviewers);

    internal static IrCompositeScript MergeOf(ConflictResolution policy, WmlDocument baseDoc, params (string Author, WmlDocument Doc)[] reviewers)
    {
        var diff = new DocxDiffSettings().ToIrDiffSettings();
        var baseIr = IrReader.Read(baseDoc, ReadOpts);
        var revs = reviewers.Select(r => (r.Author, IrReader.Read(r.Doc, ReadOpts))).ToList();
        return IrCompositeMerger.Merge(baseIr, revs, policy, diff);
    }
}
