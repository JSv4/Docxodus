#nullable enable
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
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

    // ---- Task 2.3: composed multi-author single-paragraph rendering ----

    [Fact]
    public void Two_reviewers_editing_different_words_of_one_paragraph_compose_inline()
    {
        var baseDoc = Docs.Para("the quick brown fox jumps");
        var r1 = Docs.Para("the SLOW brown fox jumps");     // edits word 2
        var r2 = Docs.Para("the quick brown fox LEAPS");    // edits word 5
        var script = IrCompositeMergerTests.MergeOf(baseDoc, ("Bob", r1), ("Fred", r2));
        var merged = IrCompositeMarkupRenderer.Render(script, baseDoc,
            new[] { ("Bob", r1), ("Fred", r2) }, new DocxDiffSettings().ToIrDiffSettings());

        Assert.Equal("the quick brown fox jumps", Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
        var accepted = Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));
        Assert.Contains("SLOW", accepted);
        Assert.Contains("LEAPS", accepted);
        Assert.DoesNotContain("quick", accepted);
        Assert.DoesNotContain("jumps", accepted);
        var xml = Docs.MainPartXml(merged);
        Assert.Contains("w:author=\"Bob\"", xml);
        Assert.Contains("w:author=\"Fred\"", xml);
    }

    [Fact]
    public void Three_reviewers_editing_different_words_of_one_paragraph_compose_inline()
    {
        var baseDoc = Docs.Para("the quick brown fox jumps over");
        var r1 = Docs.Para("the SLOW brown fox jumps over");     // word 2
        var r2 = Docs.Para("the quick GREEN fox jumps over");    // word 3
        var r3 = Docs.Para("the quick brown fox LEAPS over");    // word 5
        var script = IrCompositeMergerTests.MergeOf(baseDoc, ("Bob", r1), ("Fred", r2), ("Gus", r3));
        var merged = IrCompositeMarkupRenderer.Render(script, baseDoc,
            new[] { ("Bob", r1), ("Fred", r2), ("Gus", r3) }, new DocxDiffSettings().ToIrDiffSettings());

        Assert.Equal("the quick brown fox jumps over", Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
        var accepted = Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));
        Assert.Contains("SLOW", accepted);
        Assert.Contains("GREEN", accepted);
        Assert.Contains("LEAPS", accepted);
        Assert.DoesNotContain("quick", accepted);
        Assert.DoesNotContain("brown", accepted);
        Assert.DoesNotContain("jumps", accepted);
        var xml = Docs.MainPartXml(merged);
        Assert.Contains("w:author=\"Bob\"", xml);
        Assert.Contains("w:author=\"Fred\"", xml);
        Assert.Contains("w:author=\"Gus\"", xml);
    }

    [Fact]
    public void One_reviewer_edits_two_words_other_edits_one_word_compose_inline()
    {
        var baseDoc = Docs.Para("the quick brown fox jumps over");
        var r1 = Docs.Para("the SLOW brown fox LANDS over");     // Bob edits words 2 and 5
        var r2 = Docs.Para("the quick GREEN fox jumps over");    // Fred edits word 3
        var script = IrCompositeMergerTests.MergeOf(baseDoc, ("Bob", r1), ("Fred", r2));
        var merged = IrCompositeMarkupRenderer.Render(script, baseDoc,
            new[] { ("Bob", r1), ("Fred", r2) }, new DocxDiffSettings().ToIrDiffSettings());

        Assert.Equal("the quick brown fox jumps over", Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
        var accepted = Docs.PlainText(RevisionAccepter.AcceptRevisions(merged));
        Assert.Contains("SLOW", accepted);
        Assert.Contains("LANDS", accepted);
        Assert.Contains("GREEN", accepted);
        Assert.DoesNotContain("quick", accepted);
        Assert.DoesNotContain("brown", accepted);
        Assert.DoesNotContain("jumps", accepted);
        var xml = Docs.MainPartXml(merged);
        Assert.Contains("w:author=\"Bob\"", xml);
        Assert.Contains("w:author=\"Fred\"", xml);
    }

    // ---- FOLLOW-ON A: native move composition rendering ----

    // Four ≥4-word paragraphs so a reorder is detected as a MoveBlock by the aligner.
    private const string MP1 = "First paragraph alpha bravo charlie";
    private const string MP2 = "Second paragraph delta echo foxtrot";
    private const string MP3 = "Third paragraph golf hotel india";
    private const string MP4 = "Fourth paragraph juliet kilo lima";

    private static XElement BodyOf(WmlDocument d)
    {
        using var ms = new MemoryStream(d.DocumentByteArray);
        using var wd = WordprocessingDocument.Open(ms, false);
        return wd.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }

    [Fact]
    public void Native_composite_move_emits_moveFrom_moveTo_authored_to_mover()
    {
        var baseDoc = Docs.Para(MP1, MP2, MP3, MP4);
        var alice = Docs.Para(MP1, MP3, MP4, MP2);    // Alice relocates P2 to the end
        var script = IrCompositeMergerTests.MergeOf(baseDoc, ("Alice", alice));
        var merged = IrCompositeMarkupRenderer.Render(script, baseDoc,
            new[] { ("Alice", alice) }, new DocxDiffSettings().ToIrDiffSettings());

        var body = BodyOf(merged);
        Assert.NotEmpty(body.Descendants(W.moveFrom));
        Assert.NotEmpty(body.Descendants(W.moveTo));
        Assert.NotEmpty(body.Descendants(W.moveFromRangeStart));
        Assert.NotEmpty(body.Descendants(W.moveToRangeStart));

        // moveFrom/moveTo range names pair (set-equal).
        var fromNames = body.Descendants(W.moveFromRangeStart).Select(e => (string?)e.Attribute(W.name)).ToHashSet();
        var toNames = body.Descendants(W.moveToRangeStart).Select(e => (string?)e.Attribute(W.name)).ToHashSet();
        Assert.True(fromNames.SetEquals(toNames), "moveFrom/moveTo range names must pair");

        // Authored to the mover.
        foreach (var e in body.Descendants(W.moveFrom).Concat(body.Descendants(W.moveTo)))
            Assert.Equal("Alice", (string?)e.Attribute(W.author));

        // reject ≡ base; accept ≡ Alice (the relocated body).
        Assert.Equal(Docs.PlainText(baseDoc), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
        Assert.Equal(Docs.PlainText(alice), Docs.PlainText(RevisionAccepter.AcceptRevisions(merged)));
    }

    [Fact]
    public void Native_composite_move_with_SimplifyMoveMarkup_degrades_to_ins_del()
    {
        var baseDoc = Docs.Para(MP1, MP2, MP3, MP4);
        var alice = Docs.Para(MP1, MP3, MP4, MP2);
        var settings = new DocxDiffSettings().ToIrDiffSettings() with { SimplifyMoveMarkup = true };
        var script = IrCompositeMergerTests.MergeOf(baseDoc, ("Alice", alice));
        var merged = IrCompositeMarkupRenderer.Render(script, baseDoc, new[] { ("Alice", alice) }, settings);

        var body = BodyOf(merged);
        Assert.Empty(body.Descendants(W.moveFrom));
        Assert.Empty(body.Descendants(W.moveTo));
        Assert.True(body.Descendants(W.ins).Any() || body.Descendants(W.del).Any(),
            "simplified move must use ins/del");
        Assert.Equal(Docs.PlainText(baseDoc), Docs.PlainText(RevisionProcessor.RejectRevisions(merged)));
    }
}
