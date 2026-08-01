#nullable enable

using System.Linq;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// <see cref="DocxSession.ListBlocks"/> (the incremental renderer's ordered
/// top-level unit plan) and <see cref="DocxSession.ListNotes"/> (citation-ordered
/// note ids for client-side chrome renumbering).
/// </summary>
public class DocxSessionRenderPlanTests
{
    [Fact]
    public void DS310_ListBlocks_BodyOrderAndTableUnit()
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
        var p0 = session.ListBlocks().Body.Single();
        Assert.Equal("p", p0.Kind);

        session.InsertParagraph(p0.Id, Position.After, "Second paragraph.");
        var afterPara = session.ListBlocks();
        Assert.Equal(2, afterPara.Body.Count);

        session.InsertTable(afterPara.Body[0].Id, Position.After, 2, 2);
        var plan = session.ListBlocks();
        // p, tbl, p — the table is ONE unit; its cell paragraphs are not body units.
        Assert.Equal(new[] { "p", "tbl", "p" }, plan.Body.Select(u => u.Kind).ToArray());
        Assert.StartsWith("tbl:body:", plan.Body[1].Id);
        Assert.Empty(plan.Footnotes);
        Assert.Empty(plan.Endnotes);
    }

    [Fact]
    public void DS311_ListBlocks_NotesInOrder()
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
        var p0 = session.ListBlocks().Body[0];
        session.ReplaceText(p0.Id, "Cite one here and two there.");
        var anchor = session.ListBlocks().Body[0].Id;
        session.InsertFootnote(anchor, 8, "First note.");
        anchor = session.ListBlocks().Body[0].Id;
        session.InsertFootnote(anchor, 20, "Second note.");

        var plan = session.ListBlocks();
        Assert.Equal(2, plan.Footnotes.Count);
        Assert.All(plan.Footnotes, u => Assert.Equal("fn", u.Kind));
        // Reserved separator/continuation notes are excluded.
        Assert.All(plan.Footnotes, u => Assert.StartsWith("fn:fn:", u.Id));
    }

    [Fact]
    public void DS312_ListNotes_CitationOrderAndIds()
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
        var p0 = session.ListBlocks().Body[0].Id;
        session.ReplaceText(p0, "Alpha beta gamma delta.");
        var anchor = session.ListBlocks().Body[0].Id;
        session.InsertFootnote(anchor, 22, "Late note.");   // cited at the end
        anchor = session.ListBlocks().Body[0].Id;
        session.InsertFootnote(anchor, 5, "Early note.");   // cited FIRST — ids shift

        var notes = session.ListNotes();
        Assert.Equal(2, notes.Count);
        Assert.Equal(1, notes[0].Ordinal);
        Assert.Equal(2, notes[1].Ordinal);
        // Reference-order law: ids ascend in citation order.
        Assert.True(int.Parse(notes[0].Id) < int.Parse(notes[1].Id));
        Assert.StartsWith("fn:fn:", notes[0].DefAnchorId);
        Assert.NotEqual(notes[0].DefAnchorId, notes[1].DefAnchorId);

        Assert.Empty(session.ListNotes(endnotes: true));
    }
}
