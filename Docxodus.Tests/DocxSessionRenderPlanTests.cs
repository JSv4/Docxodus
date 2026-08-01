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
        // Plan order = CITATION order (matches the rendered notes section), proven by
        // the citation-ordered ListNotes pointing at the same definitions.
        var notes = session.ListNotes();
        Assert.Equal(notes.Select(n => n.DefAnchorId).ToList(),
            plan.Footnotes.Select(u => u.Id).ToList());
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

    // EVERY unit carries a content signature: in-session an element keeps its unid
    // across edits (ops rebuild children, not the block element), so unid alone
    // cannot tell a diffing renderer that content changed — an undone text edit or a
    // row insert would silently keep a stale node.
    [Fact]
    public void DS313_Units_CarryContentSignature_ThatTracksInnerChanges()
    {
        using var session = new DocxSession(DocxSession.CreateBlankDocxBytes());
        var p0 = session.ListBlocks().Body[0].Id;
        session.InsertTable(p0, Position.After, 2, 2);

        var before = session.ListBlocks();
        var tblBefore = before.Body.Single(u => u.Kind == "tbl");
        Assert.NotNull(tblBefore.Sig);

        // A leaf paragraph's sig moves on a text edit even though its unid may not.
        var pBefore = before.Body.First(u => u.Kind == "p");
        Assert.NotNull(pBefore.Sig);
        Assert.True(session.ReplaceText(pBefore.Id, "Edited paragraph text.").Success);
        var pAfter = session.ListBlocks().Body.First(u => u.Kind == "p");
        Assert.NotEqual(pBefore.Sig, pAfter.Sig);

        // A row insert keeps the table's unid but MUST change its signature.
        var res = session.InsertTableRow(FirstCellParagraphAnchor(session), Position.After);
        Assert.True(res.Success, res.Error?.Message);

        var after = session.ListBlocks();
        var tblAfter = after.Body.Single(u => u.Kind == "tbl");
        Assert.Equal(tblBefore.Id, tblAfter.Id);       // unid stable
        Assert.NotEqual(tblBefore.Sig, tblAfter.Sig);  // content signature moved

        // Same for a footnote: text edit inside the note keeps the fn unid, changes sig.
        var bodyAnchor = after.Body.First(u => u.Kind is "p" or "h" or "li").Id;
        session.ReplaceText(bodyAnchor, "Body text.");
        bodyAnchor = session.ListBlocks().Body.First(u => u.Kind is "p" or "h" or "li").Id;
        var fnRes = session.InsertFootnote(bodyAnchor, 0, "Original note.");
        Assert.True(fnRes.Success, fnRes.Error?.Message);
        var fnBefore = session.ListBlocks().Footnotes.Single();
        var notePara = fnRes.Created.First(a => a.Kind == "p" && a.Scope == "fn");
        Assert.True(session.ReplaceText(notePara.Id, "Edited note text.").Success);
        var fnAfter = session.ListBlocks().Footnotes.Single();
        Assert.Equal(fnBefore.Id, fnAfter.Id);
        Assert.NotEqual(fnBefore.Sig, fnAfter.Sig);
    }

    private static string FirstCellParagraphAnchor(DocxSession session)
    {
        // A cell paragraph is a p:body anchor whose element sits inside w:tc — find one
        // via the raw XML escape hatch: take the table anchor's first cell paragraph unid.
        var tbl = session.ListBlocks().Body.Single(u => u.Kind == "tbl");
        var xml = session.Raw.GetXml(tbl.Id);
        var el = System.Xml.Linq.XElement.Parse(xml);
        System.Xml.Linq.XNamespace pt = "http://powertools.codeplex.com/2011";
        var unid = el.Descendants()
            .Where(d => d.Name.LocalName == "p")
            .Select(d => (string?)d.Attribute(pt + "Unid"))
            .First(u => u is not null);
        return $"p:body:{unid}";
    }
}
