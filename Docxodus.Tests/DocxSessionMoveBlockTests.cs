#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

public class DocxSessionMoveBlockTests
{
    private static readonly XNamespace Xml = XNamespace.Xml;

    private static XElement P(string text, params object[] leading) =>
        new(W.p, leading, new XElement(W.r,
            new XElement(W.t, new XAttribute(Xml + "space", "preserve"), text)));

    private static XElement Table(string text) =>
        new(W.tbl,
            new XElement(W.tblPr),
            new XElement(W.tblGrid, new XElement(W.gridCol, new XAttribute(W.w + "w", 2000))),
            new XElement(W.tr,
                new XElement(W.tc,
                    new XElement(W.tcPr),
                    P(text))));

    private static byte[] Document(params object[] bodyChildren)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body());
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
            main.Document.Save();
            var xdoc = main.GetXDocument();
            xdoc.Root!.Element(W.body)!.ReplaceNodes(bodyChildren);
            main.PutXDocument();
        }
        return stream.ToArray();
    }

    private static string[] ParagraphAnchors(DocxSession session) =>
        session.FindByKind("p", "body").Select(t => t.Anchor.Id).ToArray();

    private static string[] BodyLabels(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var body = document.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        return body.Elements()
            .Where(e => e.Name == W.p || e.Name == W.tbl)
            .Select(e => string.Concat(e.Descendants(W.t).Select(t => t.Value)))
            .ToArray();
    }

    private static byte[] Accept(byte[] bytes) =>
        RevisionProcessor.AcceptRevisions(new WmlDocument("accepted.docx", bytes)).DocumentByteArray;

    private static byte[] Reject(byte[] bytes) =>
        RevisionProcessor.RejectRevisions(new WmlDocument("rejected.docx", bytes)).DocumentByteArray;

    private static void AssertValid(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document));
    }

    [Fact]
    public void MoveBlock_Direct_ReordersSameElementAndUndoRedo()
    {
        using var session = new DocxSession(Document(P("A"), P("B"), P("C")));
        var anchors = ParagraphAnchors(session);
        var sourceUnid = session.FindByText("A")!.Anchor.Unid;

        var result = session.MoveBlock(anchors[0], anchors[2], Position.After);

        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal(sourceUnid, Assert.Single(result.Modified).Unid);
        Assert.Equal(new[] { "B", "C", "A" }, BodyLabels(session.Save()));
        Assert.True(session.Undo());
        Assert.Equal(new[] { "A", "B", "C" }, BodyLabels(session.Save()));
        Assert.True(session.Redo());
        Assert.Equal(new[] { "B", "C", "A" }, BodyLabels(session.Save()));
    }

    [Fact]
    public void MoveBlock_AlreadyAdjacent_IsNoOpAndConsumesNoUndo()
    {
        using var session = new DocxSession(Document(P("A"), P("B")));
        var anchors = ParagraphAnchors(session);

        Assert.True(session.MoveBlock(anchors[0], anchors[1], Position.Before).Success);
        Assert.False(session.Undo());
        Assert.Equal(new[] { "A", "B" }, BodyLabels(session.Save()));
    }

    [Fact]
    public void MoveBlock_Direct_MovesWholeTable()
    {
        using var session = new DocxSession(Document(P("A"), Table("TABLE"), P("B")));
        var table = Assert.Single(session.FindByKind("tbl", "body")).Anchor.Id;
        var target = ParagraphAnchors(session).Last();

        var result = session.MoveBlock(table, target, Position.After);

        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal(new[] { "A", "B", "TABLE" }, BodyLabels(session.Save()));
    }

    [Fact]
    public void MoveBlock_TrackedParagraph_EmitsNamedMoveAndRoundTrips()
    {
        using var session = new DocxSession(
            Document(P("A"), P("B"), P("C")),
            new DocxSessionSettings
            {
                TrackedChanges = TrackedChangeMode.RenderInline,
                RevisionAuthor = "Alice",
            });
        var anchors = ParagraphAnchors(session);

        var result = session.MoveBlock(anchors[0], anchors[2], Position.After);

        Assert.True(result.Success, result.Error?.Message);
        Assert.Single(result.Created);
        var saved = session.Save();
        AssertValid(saved);

        using (var stream = new MemoryStream(saved))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var main = document.MainDocumentPart!.GetXDocument();
            var from = Assert.Single(main.Descendants(W.moveFromRangeStart));
            var to = Assert.Single(main.Descendants(W.moveToRangeStart));
            Assert.Equal((string?)from.Attribute(W.name), (string?)to.Attribute(W.name));
            Assert.Equal("Alice", (string?)from.Attribute(W.author));
            Assert.NotNull(document.MainDocumentPart.DocumentSettingsPart!
                .GetXDocument().Root!.Element(W.trackRevisions));
        }

        Assert.Equal(new[] { "B", "C", "A" }, BodyLabels(Accept(saved)));
        Assert.Equal(new[] { "A", "B", "C" }, BodyLabels(Reject(saved)));
        var listed = Assert.Single(session.ListRevisions().Where(r => r.Type == "move"));
        Assert.Equal("Alice", listed.Author);
    }

    [Fact]
    public void MoveBlock_TrackedTable_UsesDeletedAndInsertedRowsAndRoundTrips()
    {
        using var session = new DocxSession(
            Document(P("A"), Table("TABLE"), P("B")),
            new DocxSessionSettings
            {
                TrackedChanges = TrackedChangeMode.RenderInline,
                RevisionAuthor = "Alice",
            });
        var table = Assert.Single(session.FindByKind("tbl", "body")).Anchor.Id;
        var target = ParagraphAnchors(session).Last();

        var result = session.MoveBlock(table, target, Position.After);

        Assert.True(result.Success, result.Error?.Message);
        var saved = session.Save();
        AssertValid(saved);
        using (var stream = new MemoryStream(saved))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var rows = document.MainDocumentPart!.GetXDocument().Descendants(W.tr).ToList();
            Assert.Contains(rows, r => r.Element(W.trPr)?.Element(W.del) is not null);
            Assert.Contains(rows, r => r.Element(W.trPr)?.Element(W.ins) is not null);
        }
        Assert.Equal(new[] { "A", "B", "TABLE" }, BodyLabels(Accept(saved)));
        Assert.Equal(new[] { "A", "TABLE", "B" }, BodyLabels(Reject(saved)));
    }

    [Fact]
    public void MoveBlock_RejectsCrossBlockRangeMembershipChange()
    {
        var start = new XElement(W.commentRangeStart, new XAttribute(W.id, 7));
        var end = new XElement(W.commentRangeEnd, new XAttribute(W.id, 7));
        using var session = new DocxSession(Document(
            P("A", start), P("B"), P("C", end), P("D")));
        var anchors = ParagraphAnchors(session);

        var result = session.MoveBlock(anchors[1], anchors[3], Position.After);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.InvalidPosition, result.Error!.Code);
        Assert.Contains("cross-block comment range", result.Error.Message);
    }

    // A tracked move clones the source paragraph. Every id-bearing marker in the clone is a
    // SECOND live copy, so the ids must be made unique or the document violates the schema's
    // id-uniqueness constraint while the revision is pending — the exact state a redline is
    // sent out in. Mirrors IrMarkupRenderer.NormalizeBookmarks step (B): both copies keep the
    // NAME (each survives its own resolution); only the ids are renumbered.
    [Fact]
    public void MoveBlock_TrackedParagraph_GivesClonedBookmarksFreshIds()
    {
        using var session = new DocxSession(
            Document(
                P("A",
                    new XElement(W.bookmarkStart, new XAttribute(W.id, 4), new XAttribute(W.name, "_Ref1")),
                    new XElement(W.bookmarkEnd, new XAttribute(W.id, 4))),
                P("B"), P("C")),
            new DocxSessionSettings
            {
                TrackedChanges = TrackedChangeMode.RenderInline,
                RevisionAuthor = "Alice",
            });
        var anchors = ParagraphAnchors(session);

        Assert.True(session.MoveBlock(anchors[0], anchors[2], Position.After).Success);

        var saved = session.Save();
        AssertValid(saved);
        using var stream = new MemoryStream(saved);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!.GetXDocument();
        var starts = main.Descendants(W.bookmarkStart).ToList();
        var ends = main.Descendants(W.bookmarkEnd).ToList();

        Assert.Equal(2, starts.Count);
        Assert.Equal(2, ends.Count);
        // Both copies keep the name; the ids are distinct and each start still pairs with an end.
        Assert.All(starts, s => Assert.Equal("_Ref1", (string?)s.Attribute(W.name)));
        var startIds = starts.Select(s => (string?)s.Attribute(W.id)).ToList();
        Assert.Equal(2, startIds.Distinct().Count());
        Assert.Equal(startIds.OrderBy(x => x), ends.Select(e => (string?)e.Attribute(W.id)).OrderBy(x => x));
    }

    // The drag UI gates its drop indicators on this, so it has to agree with MoveBlock exactly:
    // anything listed must be accepted, and anything omitted must be refused.
    [Fact]
    public void ValidMoveTargets_ExcludesTargetsAcrossASectionBreak()
    {
        var sectionBreak = new XElement(W.pPr, new XElement(W.sectPr));
        using var session = new DocxSession(Document(
            P("A"), P("B"), P("BREAK", sectionBreak), P("C"), P("D")));
        var anchors = ParagraphAnchors(session);

        var valid = session.ValidMoveTargets(anchors[0]);

        // A and B are on one side of the break; C and D on the other.
        Assert.Contains(anchors[1], valid);
        Assert.DoesNotContain(anchors[3], valid);
        Assert.DoesNotContain(anchors[4], valid);

        // Every listed target really is accepted, and an omitted one really is refused.
        foreach (var target in valid)
        {
            using var probe = new DocxSession(session.Save());
            var probeAnchors = ParagraphAnchors(probe);
            Assert.True(probe.MoveBlock(probeAnchors[0], target, Position.After).Success);
        }
        Assert.False(session.MoveBlock(anchors[0], anchors[3], Position.After).Success);
    }

    // A block already carrying revision markup cannot become a TRACKED move (re-wrapping would
    // nest revisions), but a DIRECT move relocates the element untouched and is fine. The two
    // modes must therefore disagree — which is what makes this a test of the source-level guard
    // rather than of the section-break span check.
    [Fact]
    public void ValidMoveTargets_ExcludesABlockWithExistingRevisionsOnlyWhenTracking()
    {
        var alreadyInserted = new XElement(W.p,
            new XElement(W.ins,
                new XAttribute(W.id, 1),
                new XAttribute(W.author, "Bob"),
                new XAttribute(W.date, "2026-01-01T00:00:00Z"),
                new XElement(W.r, new XElement(W.t, "A"))));

        using (var direct = new DocxSession(Document(alreadyInserted, P("B"), P("C"))))
        {
            var anchors = ParagraphAnchors(direct);
            Assert.NotEmpty(direct.ValidMoveTargets(anchors[0]));
        }

        using var tracked = new DocxSession(
            Document(new XElement(alreadyInserted), P("B"), P("C")),
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var trackedAnchors = ParagraphAnchors(tracked);

        Assert.Empty(tracked.ValidMoveTargets(trackedAnchors[0]));
        Assert.False(tracked.MoveBlock(trackedAnchors[0], trackedAnchors[2], Position.After).Success);
    }

    // Deleted content is w:delText, not w:t — that is what Word writes, what the IR renderer's
    // ConvertTextToDelText produces, and what RevisionProcessor's reject path swaps back. A
    // w:moveFrom IS a deletion, so its runs must follow the same rule.
    [Fact]
    public void MoveBlock_TrackedParagraph_MarksMovedFromTextAsDeleted()
    {
        using var session = new DocxSession(
            Document(P("A"), P("B"), P("C")),
            new DocxSessionSettings
            {
                TrackedChanges = TrackedChangeMode.RenderInline,
                RevisionAuthor = "Alice",
            });
        var anchors = ParagraphAnchors(session);

        Assert.True(session.MoveBlock(anchors[0], anchors[2], Position.After).Success);

        var saved = session.Save();
        AssertValid(saved);
        using (var stream = new MemoryStream(saved))
        using (var document = WordprocessingDocument.Open(stream, false))
        {
            var moveFrom = document.MainDocumentPart!.GetXDocument()
                .Descendants(W.moveFrom).Single(e => e.Elements(W.r).Any());
            Assert.NotEmpty(moveFrom.Descendants(W.delText));
            Assert.Empty(moveFrom.Descendants(W.t));
        }

        // …and the text still comes back intact on reject (delText → t).
        Assert.Equal(new[] { "A", "B", "C" }, BodyLabels(Reject(saved)));
        Assert.Equal(new[] { "B", "C", "A" }, BodyLabels(Accept(saved)));
    }

    // The design doc's whole-table lowering is "every source row AND ITS CONTENT is deleted, and
    // every destination row and its content is inserted". Marking only w:trPr leaves the moved-away
    // table's text rendering as ordinary body text inside a row Word believes is deleted.
    [Fact]
    public void MoveBlock_TrackedTable_MarksCellContentNotJustRows()
    {
        using var session = new DocxSession(
            Document(P("A"), Table("CELL"), P("B")),
            new DocxSessionSettings
            {
                TrackedChanges = TrackedChangeMode.RenderInline,
                RevisionAuthor = "Alice",
            });
        var table = Assert.Single(session.FindByKind("tbl", "body")).Anchor.Id;
        var target = ParagraphAnchors(session).Last();

        Assert.True(session.MoveBlock(table, target, Position.After).Success);

        var saved = session.Save();
        AssertValid(saved);
        using var stream2 = new MemoryStream(saved);
        using var document2 = WordprocessingDocument.Open(stream2, false);
        var tables = document2.MainDocumentPart!.GetXDocument().Descendants(W.tbl).ToList();
        Assert.Equal(2, tables.Count);

        var deleted = tables.Single(t => t.Descendants(W.trPr).Elements(W.del).Any());
        var inserted = tables.Single(t => t.Descendants(W.trPr).Elements(W.ins).Any());
        // The deleted table's cell text is inside w:del as w:delText…
        Assert.NotEmpty(deleted.Descendants(W.del).SelectMany(d => d.Descendants(W.delText)));
        Assert.Empty(deleted.Descendants(W.r).Descendants(W.t));
        // …and the inserted table's cell text is inside w:ins.
        Assert.NotEmpty(inserted.Descendants(W.ins).SelectMany(i => i.Descendants(W.t)));

        Assert.Equal(new[] { "A", "B", "CELL" }, BodyLabels(Accept(saved)));
        Assert.Equal(new[] { "A", "CELL", "B" }, BodyLabels(Reject(saved)));
    }

    private static (string[] Ids, int Definitions) CommentShape(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        var ids = main.GetXDocument().Descendants()
            .Where(e => e.Name == W.commentRangeStart)
            .Select(e => (string)e.Attribute(W.id)!)
            .ToArray();
        var defs = main.WordprocessingCommentsPart?.GetXDocument().Root!
            .Elements(W.comment).Count() ?? 0;
        return (ids, defs);
    }

    private static string[] DefinedCommentIds(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return document.MainDocumentPart!.WordprocessingCommentsPart?.GetXDocument().Root!
            .Elements(W.comment).Select(c => (string)c.Attribute(W.id)!).ToArray() ?? [];
    }

    // The tracked-move clone duplicates the source's comment markers. Left alone that is both a
    // schema violation (the id is uniqueness-constrained) and a visible defect — one comment shows
    // twice in Word's Reviewing pane, anchored to the source AND the destination. Mirrors
    // IrMarkupRenderer.NormalizeComments step (B): the DELETED copy (the move source) gets a fresh
    // id + a cloned definition, so accept ≡ the destination's comment and reject ≡ the source's.
    [Fact]
    public void MoveBlock_TrackedParagraph_ClonesCommentForTheMoveSource()
    {
        using var session = new DocxSession(
            Document(P("A"), P("B"), P("C")),
            new DocxSessionSettings { RevisionAuthor = "Alice" });
        var anchors = ParagraphAnchors(session);
        Assert.True(session.AddComment(anchors[0], null, "Alice", "look here").Success);

        session.SetTrackedChanges(TrackedChangeMode.RenderInline);
        Assert.True(session.MoveBlock(anchors[0], anchors[2], Position.After).Success);

        var saved = session.Save();
        AssertValid(saved);

        var (ids, defs) = CommentShape(saved);
        Assert.Equal(2, ids.Length);                 // one range on each live copy
        Assert.Equal(2, ids.Distinct().Count());     // …with DISTINCT ids
        Assert.Equal(2, defs);                       // …each resolving to its own definition

        // Exactly one comment is anchored after each resolution, and it resolves to a definition.
        // (The resolved-away copy's definition stays behind unreferenced — RevisionProcessor does
        // not prune orphaned definitions for any resolved comment, and Word ignores an unanchored
        // one. What matters is that the pane shows the comment once, not twice.)
        foreach (var resolved in new[] { Accept(saved), Reject(saved) })
        {
            AssertValid(resolved);
            var live = Assert.Single(CommentShape(resolved).Ids);
            Assert.Contains(live, DefinedCommentIds(resolved));
        }
    }

    /// <summary>Anchor unids stamped on the rendered top-level body blocks, in document order.</summary>
    private static string[] RenderedBodyAnchors(int handle, bool renderTrackedChanges)
    {
        var html = Docxodus.Internal.DocxSessionOps.RenderHtml(
            handle, "dx-", false, false, 1.0, renderTrackedChanges);
        return System.Text.RegularExpressions.Regex
            .Matches(html, "data-anchor=\"([0-9a-f]{8,})\"")
            .Select(m => m.Groups[1].Value)
            .ToArray();
    }

    private static string[] PlannedBodyAnchors(DocxSession session, bool renderTrackedChanges) =>
        session.ListBlocks(renderTrackedChanges).Body
            .Select(u => u.Id.Split(':')[^1])
            .ToArray();

    // The editor's incremental reconciler diffs the rendered DOM against ListBlocks. After a
    // tracked move the two must still describe the SAME units in BOTH review and accepted views,
    // or every later structural op diffs against a unit the DOM can never contain.
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void MoveBlock_TrackedParagraph_RenderPlanMatchesRenderedBlocks(bool renderTrackedChanges)
    {
        var handle = Docxodus.Internal.DocxSessionOps.OpenSession(
            Document(P("A"), P("B"), P("C"), P("D")),
            new DocxSessionSettings
            {
                TrackedChanges = TrackedChangeMode.RenderInline,
                RevisionAuthor = "Alice",
            });
        try
        {
            var session = Docxodus.Internal.SessionRegistry.Get(handle);
            var anchors = ParagraphAnchors(session);
            Assert.True(session.MoveBlock(anchors[0], anchors[2], Position.After).Success);

            Assert.Equal(
                PlannedBodyAnchors(session, renderTrackedChanges),
                RenderedBodyAnchors(handle, renderTrackedChanges));
        }
        finally
        {
            Docxodus.Internal.DocxSessionOps.CloseSession(handle);
        }
    }
}
