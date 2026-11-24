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

    private static byte[] TextboxDocument(params object[] content)
    {
        XNamespace v = "urn:schemas-microsoft-com:vml";
        return Document(new XElement(W.p, new XElement(W.r, new XElement(W.pict,
            new XElement(v + "shape", new XElement(v + "textbox", new XElement(W.txbxContent, new XElement(W.p, content))))))));
    }

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

    private static void AssertValid(byte[] bytes) =>
        PackageValidation.AssertValid(bytes, FileFormatVersions.Office2019);

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
        Assert.Equal(new[] { "A", "B", "TABLE", "" }, BodyLabels(session.Save()));
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
        Assert.Equal(new[] { "A", "B", "TABLE", "" }, BodyLabels(Accept(saved)));
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

    // A tracked move keeps source and destination copies live simultaneously. A bookmark has
    // document-global name identity, so it cannot be duplicated faithfully across both sides;
    // moving it to only one side would lose it on either accept or reject. Reject explicitly.
    [Fact]
    public void MoveBlock_TrackedParagraph_WithBookmark_IsExplicitlyUnsupported()
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

        var result = session.MoveBlock(anchors[0], anchors[2], Position.After);

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.UnsupportedInlineBoundary, result.Error!.Code);
        Assert.Single(session.ListBookmarks());
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
        Assert.Contains(valid, t => t.AnchorId == anchors[1]);
        Assert.DoesNotContain(valid, t => t.AnchorId == anchors[3]);
        Assert.DoesNotContain(valid, t => t.AnchorId == anchors[4]);

        // Every listed (target, side) pair really is accepted — a target can be reachable on one
        // side and refused on the other, so the side is part of the claim.
        foreach (var target in valid)
        {
            foreach (var (allowed, pos) in new[] { (target.Before, Position.Before), (target.After, Position.After) })
            {
                using var probe = new DocxSession(session.Save());
                var probeAnchors = ParagraphAnchors(probe);
                Assert.Equal(allowed, probe.MoveBlock(probeAnchors[0], target.AnchorId, pos).Success);
            }
        }
        Assert.False(session.MoveBlock(anchors[0], anchors[3], Position.After).Success);
    }

    // The round-trip contract has to hold in RenderInline mode too — that is where MoveBlock carries
    // rejections the direct path does not, and a rejection the shared source-level predicate does not
    // know about is one the drag UI draws a drop indicator over before the drop hard-fails.
    [Fact]
    public void ValidMoveTargets_TrackedMode_RoundTripsWithMoveBlock()
    {
        var settings = new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline };
        using var session = new DocxSession(
            Document(
                P("A",
                    new XElement(W.bookmarkStart, new XAttribute(W.id, 3), new XAttribute(W.name, "_Toc1")),
                    new XElement(W.bookmarkEnd, new XAttribute(W.id, 3))),
                P("B"),
                P("C")),
            settings);
        var anchors = ParagraphAnchors(session);

        // Word puts a _Toc bookmark on every heading, so this is the ordinary case, not a corner one.
        Assert.Empty(session.ValidMoveTargets(anchors[0]));
        Assert.Equal(EditErrorCode.UnsupportedInlineBoundary,
            session.MoveBlock(anchors[0], anchors[1], Position.After).Error!.Code);

        var valid = session.ValidMoveTargets(anchors[1]);
        Assert.NotEmpty(valid);
        foreach (var target in valid)
        {
            foreach (var (allowed, pos) in new[] { (target.Before, Position.Before), (target.After, Position.After) })
            {
                using var probe = new DocxSession(session.Save(), settings);
                var probeAnchors = ParagraphAnchors(probe);
                Assert.Equal(allowed, probe.MoveBlock(probeAnchors[1], target.AnchorId, pos).Success);
            }
        }
    }

    // A target can be legal on one side and refused on the other: moving a block INTO a
    // cross-block range changes that range's membership, while landing outside it does not. A
    // caller told only "this target is reachable" would still pick the refused side.
    [Fact]
    public void ValidMoveTargets_ReportsEachSideOfATargetSeparately()
    {
        // The bookmark spans B..C, so A may land before B but not between B and C.
        using var session = new DocxSession(Document(
            P("A"),
            P("B", new XElement(W.bookmarkStart, new XAttribute(W.id, 3), new XAttribute(W.name, "_span"))),
            P("C", new XElement(W.bookmarkEnd, new XAttribute(W.id, 3))),
            P("D")));
        var anchors = ParagraphAnchors(session);

        var target = Assert.Single(session.ValidMoveTargets(anchors[0]).Where(t => t.AnchorId == anchors[1]));

        Assert.True(target.Before);
        Assert.False(target.After);
        Assert.False(session.MoveBlock(anchors[0], anchors[1], Position.After).Success);
        Assert.True(session.MoveBlock(anchors[0], anchors[1], Position.Before).Success);
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

    // A move source uses ordinary text, even though accepting removes it like a deletion.
    [Fact]
    public void MoveBlock_TrackedParagraph_KeepsOrdinaryTextInMoveSource()
    {
        const string text = "MovedUniquePhrase";
        using var session = new DocxSession(
            Document(P(text), P("B"), P("C")),
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
            Assert.NotEmpty(moveFrom.Descendants(W.t));
            Assert.Empty(moveFrom.Descendants(W.delText));
        }

        // …and the text still comes back intact on reject.
        Assert.Equal(new[] { text, "B", "C" }, BodyLabels(Reject(saved)));
        Assert.Equal(new[] { "B", "C", text }, BodyLabels(Accept(saved)));
        Assert.Single(session.Grep(text));
        Assert.Equal(2, session.Project().Markdown.Split(text).Length);
        Assert.Equal(text + "¶", Assert.Single(session.ListRevisions(), r => r.Type == "move").Text);
    }

    [Fact]
    public void MoveBlock_TrackedField_KeepsItsInstructionAndRoundTrips()
    {
        var field = new XElement(W.p,
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin"))),
            new XElement(W.r, new XElement(W.instrText, new XAttribute(Xml + "space", "preserve"), " PAGE ")),
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate"))),
            new XElement(W.r, new XElement(W.t, "1")),
            new XElement(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, "end"))));
        using var session = new DocxSession(Document(field, P("B"), P("C")),
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var anchors = ParagraphAnchors(session);
        Assert.True(session.MoveBlock(anchors[0], anchors[2], Position.After).Success);
        var saved = session.Save();
        AssertValid(saved);

        string[] FieldParts(byte[] bytes)
        {
            using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
            return document.MainDocumentPart!.GetXDocument().Descendants()
                .Where(e => e.Name == W.fldChar || e.Name == W.instrText)
                .Select(e => e.Name == W.fldChar ? (string?)e.Attribute(W.fldCharType) ?? "" : e.Value).ToArray();
        }
        using var redline = WordprocessingDocument.Open(new MemoryStream(saved), false);
        var moved = redline.MainDocumentPart!.GetXDocument().Descendants(W.moveFrom).ToList();
        Assert.Contains(moved, e => e.Descendants(W.instrText).Any());
        Assert.DoesNotContain(moved, e => e.Descendants(W.delInstrText).Any());
        var expected = new[] { "begin", " PAGE ", "separate", "end" };
        Assert.Equal(expected, FieldParts(Accept(saved)));
        Assert.Equal(expected, FieldParts(Reject(saved)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TextboxDeletionWithOrdinaryText_IsListedAndResolves(bool accept)
    {
        var deletion = new XElement(W.del, new XAttribute(W.id, "1"), new XAttribute(W.author, "Alice"),
            new XElement(W.r, new XElement(W.t, "Textbox unique text")));
        using var session = new DocxSession(TextboxDocument(deletion));
        var revision = Assert.Single(session.ListRevisions());
        Assert.Equal("Textbox unique text", revision.Text);
        Assert.True((accept ? session.AcceptRevision(revision.Id) : session.RejectRevision(revision.Id)).Success);
        Assert.Equal(accept ? "" : "Textbox unique text", string.Concat(BodyLabels(session.Save())));
    }

    [Fact]
    public void TextboxDeletion_DoesNotClaimNestedInsertionText()
    {
        var insertion = new XElement(W.ins, new XAttribute(W.id, "2"), new XAttribute(W.author, "Bob"),
            new XElement(W.r, new XElement(W.t, "nested")));
        var deletion = new XElement(W.del, new XAttribute(W.id, "1"), new XAttribute(W.author, "Alice"),
            new XElement(W.r, new XElement(W.t, "deleted")), insertion);
        using var session = new DocxSession(TextboxDocument(deletion));
        Assert.Equal("deleted", Assert.Single(session.ListRevisions(), r => r.Type == "delete").Text);
    }

    [Fact]
    public void TextboxDeletion_DoesNotClaimAnotherAuthorsNestedDeletion()
    {
        var inner = new XElement(W.del, new XAttribute(W.id, "2"), new XAttribute(W.author, "Bob"),
            new XElement(W.r, new XElement(W.t, "inner")));
        var outer = new XElement(W.del, new XAttribute(W.id, "1"), new XAttribute(W.author, "Alice"),
            new XElement(W.r, new XElement(W.t, "outer")), inner);
        using var session = new DocxSession(TextboxDocument(outer));
        Assert.Equal("outer", Assert.Single(session.ListRevisions(), r => r.Author == "Alice").Text);
        Assert.Equal("inner", Assert.Single(session.ListRevisions(), r => r.Author == "Bob").Text);
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

        Assert.Equal(new[] { "A", "B", "CELL", "" }, BodyLabels(Accept(saved)));
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

    // ─── Differential: the fast guard vs. the literal set-membership definition ───────────

    /// <summary>
    /// The ORIGINAL, deliberately naive move guard: rebuild the reordered block sequence and
    /// compare the actual member SETS of every cross-block range. This is the definition the
    /// index-arithmetic guard in <c>DocxSession.BlockMoveSafetyError</c> has to reproduce, kept
    /// here as an independent oracle so the optimization cannot quietly change which moves the
    /// engine accepts.
    /// </summary>
    private static bool ReferenceMoveAllowed(XElement body, int sourceIndex, int targetIndex, Position pos)
    {
        var blocks = body.Elements().Where(e => e.Name == W.p || e.Name == W.tbl).ToList();
        var source = blocks[sourceIndex];
        var target = blocks[targetIndex];

        // Source-level rejections, which make a block immovable everywhere.
        if (source.Element(W.pPr)?.Element(W.sectPr) is not null) return false;
        if (source.DescendantsAndSelf().Any(e =>
                e.Name == W.moveFromRangeStart || e.Name == W.moveFromRangeEnd ||
                e.Name == W.moveToRangeStart || e.Name == W.moveToRangeEnd))
            return false;

        var reordered = blocks.ToList();
        reordered.RemoveAt(sourceIndex);
        int targetAfterRemoval = reordered.IndexOf(target);
        reordered.Insert(pos == Position.Before ? targetAfterRemoval : targetAfterRemoval + 1, source);

        int lo = System.Math.Min(sourceIndex, targetIndex);
        int hi = System.Math.Max(sourceIndex, targetIndex);
        if (blocks.Skip(lo).Take(hi - lo + 1)
            .Any(e => e.Name == W.p && e.Element(W.pPr)?.Element(W.sectPr) is not null))
            return false;

        System.Collections.Generic.HashSet<XElement>? Members(
            System.Collections.Generic.IReadOnlyList<XElement> order, XElement start, XElement end)
        {
            int a = -1, b = -1;
            for (int i = 0; i < order.Count; i++)
            {
                if (ReferenceEquals(order[i], start)) a = i;
                if (ReferenceEquals(order[i], end)) b = i;
            }
            if (a < 0 || b < a) return null;
            return order.Skip(a).Take(b - a + 1).ToHashSet();
        }

        XElement? Owner(XElement marker) => marker.Ancestors()
            .FirstOrDefault(e => ReferenceEquals(e.Parent, body) && (e.Name == W.p || e.Name == W.tbl));

        var pairs = new[]
        {
            (Start: W.commentRangeStart, End: W.commentRangeEnd),
            (Start: W.bookmarkStart, End: W.bookmarkEnd),
            (Start: W.permStart, End: W.permEnd),
            (Start: W.moveFromRangeStart, End: W.moveFromRangeEnd),
            (Start: W.moveToRangeStart, End: W.moveToRangeEnd),
        };
        foreach (var (startName, endName) in pairs)
        {
            var starts = body.Descendants(startName)
                .GroupBy(e => (string?)e.Attribute(W.id) ?? "")
                .ToDictionary(g => g.Key, g => g.First());
            foreach (var end in body.Descendants(endName))
            {
                if (!starts.TryGetValue((string?)end.Attribute(W.id) ?? "", out var start)) continue;
                var startBlock = Owner(start);
                var endBlock = Owner(end);
                if (startBlock is null || endBlock is null || ReferenceEquals(startBlock, endBlock))
                    continue;
                var before = Members(blocks, startBlock, endBlock);
                var after = Members(reordered, startBlock, endBlock);
                if (before is null || after is null || !before.SetEquals(after)) return false;
            }
        }
        return true;
    }

    private static XElement BookmarkStart(int id) =>
        new(W.bookmarkStart, new XAttribute(W.id, id), new XAttribute(W.name, $"_r{id}"));

    private static XElement BookmarkEnd(int id) => new(W.bookmarkEnd, new XAttribute(W.id, id));

    /// <summary>
    /// Bodies that put the guard through every shape that matters: a range the source sits
    /// inside, ranges the source is an ENDPOINT of, nested and overlapping ranges, ranges
    /// spanning a table, a single-block range (which constrains nothing), a dangling end
    /// marker, and section breaks that partition the document.
    /// </summary>
    public static TheoryData<string, object[]> MoveGuardBodies()
    {
        var sectionBreak = new XElement(W.pPr, new XElement(W.sectPr));
        return new TheoryData<string, object[]>
        {
            {
                "nested and overlapping bookmark ranges",
                new object[]
                {
                    P("A", BookmarkStart(1)),
                    P("B", BookmarkStart(2)),
                    P("C"),
                    P("D", BookmarkEnd(2)),
                    P("E", BookmarkEnd(1)),
                    P("F", BookmarkStart(3)),
                    P("G"),
                    P("H", BookmarkEnd(3)),
                }
            },
            {
                "source is a range endpoint, plus a single-block range",
                new object[]
                {
                    P("A", BookmarkStart(1)),
                    P("B"),
                    P("C", BookmarkEnd(1)),
                    P("D", BookmarkStart(2), BookmarkEnd(2)),
                    P("E"),
                }
            },
            {
                "comment and permission ranges spanning a table",
                new object[]
                {
                    P("A"),
                    P("B", new XElement(W.commentRangeStart, new XAttribute(W.id, 7))),
                    Table("T"),
                    P("D", new XElement(W.commentRangeEnd, new XAttribute(W.id, 7))),
                    P("E", new XElement(W.permStart, new XAttribute(W.id, 9))),
                    P("F"),
                    P("G", new XElement(W.permEnd, new XAttribute(W.id, 9))),
                }
            },
            {
                "section breaks partition a bookmarked body",
                new object[]
                {
                    P("A"),
                    P("B", BookmarkStart(1)),
                    P("C"),
                    P("BREAK", sectionBreak),
                    P("D"),
                    P("E", BookmarkEnd(1)),
                    P("F"),
                }
            },
            {
                "dangling end marker and a duplicated start id",
                new object[]
                {
                    P("A", BookmarkStart(1)),
                    P("B", BookmarkStart(1)),
                    P("C", BookmarkEnd(1)),
                    P("D", BookmarkEnd(5)),
                    P("E"),
                }
            },
        };
    }

    // The optimization that made ValidMoveTargets answer a whole drag in one pass replaced set
    // comparison with index arithmetic. That is only sound if it decides EVERY (source, target,
    // side) triple the same way, so assert exactly that against the literal definition — for the
    // reported targets AND for the ones left out, which is where an over-eager guard would hide.
    [Theory]
    [MemberData(nameof(MoveGuardBodies))]
    public void ValidMoveTargets_MatchesTheSetMembershipDefinitionForEveryPair(
        string scenario, object[] bodyChildren)
    {
        Assert.NotEmpty(scenario);
        var body = new XElement(W.body, bodyChildren);
        using var session = new DocxSession(Document(body.Elements().ToArray()));
        var ordered = session.ListBlocks().Body.Select(u => u.Id).ToArray();
        int n = ordered.Length;
        Assert.Equal(body.Elements().Count(e => e.Name == W.p || e.Name == W.tbl), n);

        for (int s = 0; s < n; s++)
        {
            var reported = session.ValidMoveTargets(ordered[s]).ToDictionary(t => t.AnchorId, t => t);
            for (int t = 0; t < n; t++)
            {
                if (t == s) continue;
                reported.TryGetValue(ordered[t], out var entry);
                foreach (var (pos, actual) in new[]
                         {
                             (Position.Before, entry?.Before ?? false),
                             (Position.After, entry?.After ?? false),
                         })
                {
                    Assert.Equal(ReferenceMoveAllowed(body, s, t, pos), actual);
                }
            }
        }
    }

    // ValidMoveTargets is only useful if MoveBlock agrees with it. Probe the engine itself for
    // every pair on a fresh session, so a listed pair MoveBlock rejects (or an omitted pair it
    // would have accepted) fails here rather than as a drag that mysteriously does nothing.
    [Theory]
    [MemberData(nameof(MoveGuardBodies))]
    public void ValidMoveTargets_AgreesWithMoveBlockForEveryPair(string scenario, object[] bodyChildren)
    {
        Assert.NotEmpty(scenario);
        var body = new XElement(W.body, bodyChildren);
        var bytes = Document(body.Elements().ToArray());
        using var session = new DocxSession(bytes);
        var ordered = session.ListBlocks().Body.Select(u => u.Id).ToArray();

        for (int s = 0; s < ordered.Length; s++)
        {
            var reported = session.ValidMoveTargets(ordered[s]).ToDictionary(t => t.AnchorId, t => t);
            for (int t = 0; t < ordered.Length; t++)
            {
                if (t == s) continue;
                reported.TryGetValue(ordered[t], out var entry);
                foreach (var (pos, expected) in new[]
                         {
                             (Position.Before, entry?.Before ?? false),
                             (Position.After, entry?.After ?? false),
                         })
                {
                    // A source already in the requested slot is a no-op the guard never sees; it
                    // succeeds regardless of whether the pair was offered.
                    if ((pos == Position.Before && t == s + 1) ||
                        (pos == Position.After && t == s - 1)) continue;
                    using var probe = new DocxSession(bytes);
                    var probeAnchors = probe.ListBlocks().Body.Select(u => u.Id).ToArray();
                    Assert.Equal(expected, probe.MoveBlock(probeAnchors[s], probeAnchors[t], pos).Success);
                }
            }
        }
    }

    // ─── Ranges, fields and tables beside the moved block (issue #1861) ───────────────────

    private static XElement Control(params object[] blocks) =>
        new(W.sdt, new XElement(W.sdtPr, new XElement(W.id, new XAttribute(W.val, 11))), new XElement(W.sdtContent, blocks));

    private static XElement FieldChar(string type) =>
        new(W.r, new XElement(W.fldChar, new XAttribute(W.fldCharType, type)));

    /// <summary>A table of contents: its field starts in "E1", ends in "end" and shows "E2" and "E3".</summary>
    private static object[] Contents() => new object[]
    {
        P("E1", FieldChar("begin"), new XElement(W.r, new XElement(W.instrText, " TOC \\o \"1-1\" ")), FieldChar("separate")),
        P("E2", FieldChar("begin"), new XElement(W.r, new XElement(W.instrText, " PAGEREF _Toc1 ")), FieldChar("end")),
        P("E3"), P("end", FieldChar("end")),
    };

    private static XElement BodyOf(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return document.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }

    /// <summary>The body's children: a paragraph's text ("-" for an empty one), a table's in brackets, a
    /// content control's blocks in parentheses, a custom XML element's in braces, and "&lt;" or "&gt;" for a
    /// range marker between blocks.</summary>
    private static string Shape(byte[] bytes)
    {
        static string Of(XElement e) =>
            e.Name == W.tbl ? $"[{Text(e)}]"
            : e.Name == W.sdt ? "(" + string.Join(" ", e.Element(W.sdtContent)!.Elements().Select(Of)) + ")"
            : e.Name == W.customXml ? "{" + string.Join(" ", e.Elements().Where(c => c.Name != W.customXmlPr).Select(Of)) + "}"
            : e.Name == W.p ? (Text(e) is { Length: > 0 } text ? text : "-")
            : e.Name.LocalName.EndsWith("Start") ? "<" : e.Name.LocalName.EndsWith("End") ? ">" : e.Name.LocalName;
        return string.Join(" ", BodyOf(bytes).Elements().Select(Of));
    }

    private static string Text(XElement e) => string.Concat(e.Descendants(W.t).Select(t => t.Value));

    private static string Block(DocxSession session, string text) =>
        session.ListBlocks().Body.Select(u => u.Id).Single(id => Text(XElement.Parse(session.Raw.GetXml(id))) == text);

    private static bool IsOffered(DocxSession session, string source, string target, Position pos) =>
        session.ValidMoveTargets(source).SingleOrDefault(t => t.AnchorId == target) is { } side
        && (pos == Position.Before ? side.Before : side.After);

    /// <summary>Move a block named by its text, and check that <c>ValidMoveTargets</c> said the same.</summary>
    private static EditResult Move(DocxSession session, string source, string target, Position pos)
    {
        var (from, to) = (Block(session, source), Block(session, target));
        bool offered = IsOffered(session, from, to, pos);
        var result = session.MoveBlock(from, to, pos);
        Assert.Equal(offered, result.Success);
        return result;
    }

    private static DocxSession Tracked(byte[] bytes) => new(bytes,
        new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline, RevisionAuthor = "Alice" });

    /// <summary>The same, resolving one revision after another in the session.</summary>
    private static byte[] ResolveEach(byte[] bytes, bool accept)
    {
        using var session = new DocxSession(bytes);
        while (session.ListRevisions().FirstOrDefault() is { } revision)
            Assert.True((accept ? session.AcceptRevision(revision.Id) : session.RejectRevision(revision.Id)).Success);
        return session.Save();
    }

    /// <summary>What each range and field in the body covers: the words between its two ends.</summary>
    private static string[] Coverage(XElement body)
    {
        var elements = body.Descendants().ToList();
        var fields = new System.Collections.Generic.Stack<int>();
        var found = new System.Collections.Generic.List<string>();
        for (int i = 0; i < elements.Count; i++)
        {
            var (name, id) = (elements[i].Name.LocalName, (string?)elements[i].Attribute(W.id));
            var fieldCharacter = (string?)elements[i].Attribute(W.fldCharType);
            if (fieldCharacter == "begin") fields.Push(i);
            if (fieldCharacter != "end" && !name.EndsWith("End")) continue;
            int start = fieldCharacter is not null ? (fields.Count > 0 ? fields.Pop() : -1)
                : elements.FindIndex(e => e.Name.LocalName == name[..^3] + "Start" && (string?)e.Attribute(W.id) == id);
            var words = elements.Take(i).Skip(start).Where(e => e.Name == W.t).Select(t => t.Value).OrderBy(v => v, System.StringComparer.Ordinal);
            found.Add(start < 0 || start > i ? $"{name}{id} broken" : $"{name}{id}: {string.Join(" ", words)}");
        }
        return found.OrderBy(v => v, System.StringComparer.Ordinal).ToArray();
    }

    /// <summary>Whether a table touches another one or ends the body, past markers and the edges of
    /// content controls, neither of which keeps two tables apart.</summary>
    private static bool TableTouches(XElement body)
    {
        static System.Collections.Generic.IEnumerable<XElement> Flat(XElement container) => container.Elements()
            .SelectMany(e => e.Name == W.sdt ? Flat(e.Element(W.sdtContent)!) : new[] { e })
            .Where(e => e.Name == W.p || e.Name == W.tbl);
        var blocks = Flat(body).ToList();
        return blocks[^1].Name == W.tbl || blocks.Zip(blocks.Skip(1)).Any(p => p.First.Name == W.tbl && p.Second.Name == W.tbl);
    }

    public static TheoryData<string, string, string, Position, string> BodyLevelRangeMoves() => new()
    {
        // Word ends a bookmark over a whole paragraph after the paragraph, directly in the body.
        { "end", "A", "C", Position.After, "X B C A >" },
        { "end", "A", "X", Position.Before, "A > X B C" },
        { "start", "A", "C", Position.After, "X B C < A" },
        { "start", "A", "X", Position.Before, "< A X B C" },
        { "both", "A", "C", Position.After, "X B C < A >" },
        { "both", "A", "X", Position.Before, "< A > X B C" },
        // Beside the covered block is outside its range.
        { "end", "C", "A", Position.After, "X A > C B" },
        { "start", "C", "A", Position.Before, "X C < A B" },
        { "both", "X", "A", Position.After, "< A > X B C" },
        { "both", "C", "A", Position.Before, "X C < A > B" },
    };

    [Theory]
    [MemberData(nameof(BodyLevelRangeMoves))]
    public void MoveBlock_KeepsARangeWithBodyLevelMarkersWhole(string markers, string source, string target, Position pos, string expected)
    {
        foreach (var (start, end) in new[]
            { (W.bookmarkStart, W.bookmarkEnd), (W.commentRangeStart, W.commentRangeEnd), (W.permStart, W.permEnd) })
        {
            var first = new XElement(start, new XAttribute(W.id, 3), start == W.bookmarkStart ? new XAttribute(W.name, "Span") : null);
            var last = new XElement(end, new XAttribute(W.id, 3));
            var covered = P("A", markers == "end" ? first : null!);
            if (markers == "start") covered.Add(last);
            var body = new object?[] { P("X"), markers == "end" ? null : first, covered, markers == "start" ? null : last, P("B"), P("C") };
            using var session = new DocxSession(Document(body.OfType<XElement>().ToArray()));
            var before = session.Save();

            var result = Move(session, source, target, pos);

            Assert.True(result.Success, result.Error?.Message);
            Assert.Equal(expected, Shape(session.Save()));
            Assert.Equal(Coverage(BodyOf(before)), Coverage(BodyOf(session.Save())));
            Assert.True(session.Undo());
            Assert.Equal(Shape(before), Shape(session.Save()));
        }
    }

    [Theory]
    [InlineData("A", "D", Position.After)]   // a covered block leaves
    [InlineData("B", "X", Position.Before)]
    [InlineData("D", "A", Position.After)]   // another block lands among the covered ones
    [InlineData("X", "B", Position.Before)]
    public void MoveBlock_RefusesToChangeWhatABodyLevelRangeCovers(string source, string target, Position pos)
    {
        using var session = new DocxSession(Document(P("X"), BookmarkStart(3), P("A"), P("B"), BookmarkEnd(3), P("D")));

        var result = Move(session, source, target, pos);

        Assert.Equal(EditErrorCode.InvalidPosition, result.Error?.Code);
        Assert.Contains("cross-block bookmark range", result.Error!.Message);
        Assert.Equal("X < A B > D", Shape(session.Save()));
        Assert.False(session.Undo());
    }

    [Fact]
    public void MoveBlock_RefusesToCutARangeThatEndsInAContentControl()
    {
        using var session = new DocxSession(Document(P("X"), P("A", BookmarkStart(3)), Control(P("S", BookmarkEnd(3)), P("T")), P("B")));

        Assert.False(Move(session, "A", "B", Position.After).Success);
        Assert.False(Move(session, "T", "S", Position.Before).Success);
        Assert.True(Move(session, "B", "X", Position.Before).Success);
        Assert.Equal("B X A (S T)", Shape(session.Save()));
    }

    [Fact]
    public void MoveBlock_BetweenMarkersOnly_IsInPlaceAndTakesNoUndoStep()
    {
        using var session = new DocxSession(Document(P("X", BookmarkStart(3)), P("A"), BookmarkEnd(3), P("B")));

        Assert.True(Move(session, "A", "B", Position.Before).Success);

        Assert.Equal("X A > B", Shape(session.Save()));
        Assert.False(session.Undo());
    }

    // An empty bookmark between blocks (Word's own _GoBack) marks the start of the block after it.
    [Fact]
    public void MoveBlock_KeepsAnEmptyBodyLevelBookmarkWithTheBlockAfterIt()
    {
        var bytes = Document(P("X"), BookmarkStart(3), BookmarkEnd(3), P("A"), P("B"));
        using (var away = new DocxSession(bytes))
        {
            Assert.True(Move(away, "A", "B", Position.After).Success);
            Assert.Equal("X B < > A", Shape(away.Save()));
        }
        using (var beside = new DocxSession(bytes))
        {
            Assert.True(Move(beside, "B", "A", Position.Before).Success);
            Assert.Equal("X B < > A", Shape(beside.Save()));
            Assert.True(Move(beside, "X", "B", Position.After).Success);
            Assert.Equal("B X < > A", Shape(beside.Save()));
        }
        // It covers nothing, so it never stops a move, not even a tracked one, which carries no markers.
        using var tracked = Tracked(bytes);
        Assert.True(Move(tracked, "A", "B", Position.After).Success);
    }

    private static XElement ProofErr() => new(W.proofErr, new XAttribute(W.type, "spellEnd"));

    private static XElement Moved(XName name) => new(name, new XAttribute(W.id, 2), name.LocalName.EndsWith("End") ? null
        : new object[] { new XAttribute(W.name, "move1"), new XAttribute(W.author, "Bob"), new XAttribute(W.date, "2031-01-01T00:00:00Z") });

    public static TheoryData<object[], string, string, Position, string> CrowdedMarkerMoves() => new()
    {
        // Word starts a new range before it ends another at the same place. An empty one straddling that
        // end stays whole, and goes with the block only when the other range's end has to.
        { new object[] { P("X", BookmarkStart(5)), BookmarkStart(7), BookmarkEnd(5), BookmarkEnd(7), P("A"), P("B"), P("C") }, "A", "C", Position.After, "X < > > B C A" },
        { new object[] { P("X", BookmarkStart(5)), BookmarkStart(7), BookmarkEnd(5), BookmarkEnd(7), P("A"), P("B"), P("C") }, "A", "X", Position.Before, "A X < > > B C" },
        { new object[] { P("X", BookmarkStart(5)), BookmarkStart(7), BookmarkEnd(5), BookmarkEnd(7), P("A"), P("B"), P("C") }, "C", "X", Position.After, "X < > > C A B" },
        { new object[] { P("A", BookmarkStart(5)), BookmarkStart(7), BookmarkEnd(5), BookmarkEnd(7), P("B"), P("C") }, "A", "C", Position.After, "B C A < > >" },
        // A marker without its other end, or a proofing mark, goes only with one beyond it that does.
        { new object[] { P("A", BookmarkStart(5)), ProofErr(), BookmarkEnd(5), P("B"), P("C") }, "A", "C", Position.After, "B C A proofErr >" },
        { new object[] { P("X"), BookmarkEnd(9), P("A"), BookmarkEnd(8), P("B") }, "A", "B", Position.After, "X > > B A" },
        // Nothing beside the target decides: right beside it.
        { new object[] { P("A"), BookmarkEnd(9), P("B"), P("C") }, "C", "A", Position.After, "A C > B" },
        // A native move range never moves.
        { new object[] { P("X"), Moved(W.moveToRangeStart), Moved(W.moveToRangeEnd), P("A"), P("B") }, "A", "B", Position.After, "X < > B A" },
        // An end before its start is no range, as ListBookmarks says: it stops nothing.
        { new object[] { P("X"), BookmarkEnd(3), P("B"), P("C"), P("A", BookmarkStart(3)), P("D"), P("E") }, "E", "D", Position.Before, "X > B C A E D" },
        { new object[] { P("X"), BookmarkEnd(3), P("B"), P("C"), P("A", BookmarkStart(3)), P("D"), P("E") }, "A", "X", Position.After, "X A > B C D E" },
    };

    [Theory]
    [MemberData(nameof(CrowdedMarkerMoves))]
    public void MoveBlock_AmongMarkersBetweenBlocks(object[] body, string source, string target, Position pos, string expected)
    {
        foreach (var mode in new[] { TrackedChangeMode.Accept, TrackedChangeMode.StripDeletions })
        {
            using var session = new DocxSession(Document(body), new DocxSessionSettings { TrackedChanges = mode });
            var before = Coverage(BodyOf(session.Save())).Where(c => !c.EndsWith("broken"));

            var result = Move(session, source, target, pos);

            Assert.True(result.Success, result.Error?.Message);
            Assert.Equal(expected, Shape(session.Save()));
            Assert.Subset(Coverage(BodyOf(session.Save())).ToHashSet(), before.ToHashSet());
        }
    }

    [Fact]
    public void MoveBlock_RefusesToMoveABlockOutOfANativeMoveRange()
    {
        using var session = new DocxSession(Document(
            P("X"), Moved(W.moveFromRangeStart), P("A"), Moved(W.moveFromRangeEnd), P("B")));
        Assert.False(Move(session, "A", "B", Position.After).Success);
        Assert.True(Move(session, "B", "X", Position.Before).Success);
    }

    [Fact]
    public void MoveBlock_RefusesToChangeARangeThatEndsAfterItsContentControl()
    {
        using var session = new DocxSession(Document(P("X"), Control(P("A"), P("B", BookmarkStart(3)), P("E")), P("D", BookmarkEnd(3))));
        Assert.False(Move(session, "A", "E", Position.After).Success);
        Assert.False(Move(session, "E", "A", Position.Before).Success);
    }

    [Theory]
    [InlineData("B", "E2", Position.Before)]   // lands in the field's result
    [InlineData("A", "E1", Position.After)]
    [InlineData("B", "end", Position.Before)]
    [InlineData("E2", "B", Position.After)]    // leaves it
    [InlineData("E1", "B", Position.After)]    // takes the field's start, or its end, with it
    [InlineData("end", "B", Position.After)]
    [InlineData("end", "E2", Position.Before)]
    public void MoveBlock_RefusesToChangeWhatAFieldCovers(string source, string target, Position pos)
    {
        var bytes = Document(Contents().Concat(new object[] { P("A"), P("B") }).ToArray());
        foreach (bool trackChanges in new[] { false, true })
        {
            using var session = trackChanges ? Tracked(bytes) : new DocxSession(bytes);

            var result = Move(session, source, target, pos);

            Assert.Equal(EditErrorCode.InvalidPosition, result.Error?.Code);
            Assert.Contains("cross-block field range", result.Error!.Message);
            Assert.Equal("E1 E2 E3 end A B", Shape(session.Save()));
        }
    }

    [Theory]
    [InlineData("B", "E1", Position.Before, "B E1 E2 E3 end A")]
    [InlineData("B", "end", Position.After, "E1 E2 E3 end B A")]
    [InlineData("E3", "E2", Position.Before, "E1 E3 E2 end A B")]
    public void MoveBlock_AroundOrWithinAField_StillMoves(string source, string target, Position pos, string expected)
    {
        using var session = new DocxSession(Document(Contents().Concat(new object[] { P("A"), P("B") }).ToArray()));

        Assert.True(Move(session, source, target, pos).Success);

        Assert.Equal(expected, Shape(session.Save()));
        Assert.True(session.UpdateFields().Success);
        Assert.Contains("B", Shape(session.Save()).Split(' '));
    }

    public static TheoryData<string, object[], string, string, Position, string> TableMoves()
    {
        var permission = new object[] { new XElement(W.permStart, new XAttribute(W.id, 8)), new XElement(W.permEnd, new XAttribute(W.id, 8)) };
        return new()
        {
            { "a paragraph leaves from between two tables", new object[] { Table("T1"), P("B"), Table("T2"), P("C") }, "B", "C", Position.After, "[T1] - [T2] C B" },
            { "the paragraph after the last table leaves", new object[] { P("A"), Table("T1"), P("B") }, "B", "A", Position.Before, "B A [T1] -" },
            { "the same, with section properties", new object[] { P("A"), Table("T1"), P("B"), new XElement(W.sectPr) }, "B", "A", Position.Before, "B A [T1] - sectPr" },
            { "a table lands at the end", new object[] { P("A"), Table("T1"), P("B"), new XElement(W.sectPr) }, "T1", "B", Position.After, "A B [T1] - sectPr" },
            { "a table lands after a table", new object[] { P("A"), Table("T1"), P("B"), Table("T2"), P("C") }, "T1", "T2", Position.After, "A B [T2] - [T1] C" },
            { "a table lands before a table", new object[] { P("A"), Table("T1"), P("B"), Table("T2"), P("C") }, "T2", "T1", Position.Before, "A [T2] - [T1] B C" },
            { "permission markers between tables", new object[] { Table("T1"), P("B"), permission[0], permission[1], Table("T2"), P("C") }, "B", "C", Position.After, "[T1] - < > [T2] C B" },
            { "a table that ends a content control", new object[] { Control(P("S"), Table("T1")), P("B"), Table("T2"), P("C") }, "B", "C", Position.After, "(S [T1]) - [T2] C B" },
            { "a table that starts one", new object[] { Table("T1"), P("B"), Control(Table("T2"), P("S")), P("C") }, "B", "C", Position.After, "[T1] - ([T2] S) C B" },
            { "to the end of a control that ends the body", new object[] { P("A"), Control(Table("T1"), P("B"), P("C")) }, "T1", "C", Position.After, "A (B C [T1] -)" },
            { "to the end of a control before a table", new object[] { Control(Table("T1"), P("B")), Table("T2"), P("C") }, "T1", "B", Position.After, "(B [T1] -) [T2] C" },
            { "inside a control that something follows", new object[] { P("A"), Control(Table("T1"), P("B"), Table("T2")), P("C") }, "T1", "T2", Position.After, "A (B [T2] - [T1]) C" },
            { "to the start of a control after a table", new object[] { Table("T1"), Control(P("B"), Table("T2"), P("S")), P("C") }, "T2", "B", Position.Before, "[T1] (- [T2] B S) C" },
        };
    }

    [Theory]
    [MemberData(nameof(TableMoves))]
    public void MoveBlock_LeavesAParagraphBetweenTables_AndAfterTheLast(
        string scenario, object[] body, string source, string target, Position pos, string expected)
    {
        Assert.NotEmpty(scenario);
        var bytes = Document(body);
        using var session = new DocxSession(bytes);

        var result = Move(session, source, target, pos);

        Assert.True(result.Success, result.Error?.Message);
        var saved = session.Save();
        Assert.Equal(expected, Shape(saved));
        Assert.Equal("p", Assert.Single(result.Created).Kind);
        Assert.Equal(source, Text(XElement.Parse(session.Raw.GetXml(Assert.Single(result.Modified).Id))));
        AssertValid(saved);
        Assert.True(session.Undo());
        Assert.Equal(Shape(bytes), Shape(session.Save()));
        Assert.False(session.Undo());
    }

    // Accepting a tracked move must leave the same paragraphs, whichever way it is accepted, and rejecting
    // it none of them.
    [Theory]
    [MemberData(nameof(TableMoves))]
    public void MoveBlock_Tracked_LeavesAParagraphBetweenTables_WhenAccepted(
        string scenario, object[] body, string source, string target, Position pos, string expected)
    {
        Assert.NotEmpty(scenario);
        var bytes = Document(body);
        using var session = Tracked(bytes);

        var result = Move(session, source, target, pos);

        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal(source, Text(XElement.Parse(session.Raw.GetXml(result.Created[0].Id))));
        Assert.Equal("p", result.Created[1].Kind);
        Assert.Equal(2, result.Created.Count);
        var saved = session.Save();
        AssertValid(saved);
        Assert.Equal(expected, Shape(Accept(saved)));
        Assert.Equal(expected, Shape(ResolveEach(saved, accept: true)));
        Assert.Equal(Shape(bytes), Shape(Reject(saved)));
        Assert.Equal(Shape(bytes), Shape(ResolveEach(saved, accept: false)));
    }

    [Fact]
    public void MoveBlock_Tracked_OneParagraphAfterAnotherLeavesTwoTablesApart_WhenAccepted()
    {
        var bytes = Document(Table("T1"), P("A"), P("B"), Table("T2"), P("C"));
        using var session = Tracked(bytes);

        Assert.True(Move(session, "A", "C", Position.After).Success);
        Assert.True(Move(session, "B", "C", Position.After).Success);

        var saved = session.Save();
        Assert.Equal("[T1] - [T2] C B A", Shape(Accept(saved)));
        Assert.Equal("[T1] - [T2] C B A", Shape(ResolveEach(saved, accept: true)));
        Assert.Equal(Shape(bytes), Shape(Reject(saved)));
    }

    private static XElement Revised(XName name, params object[] content) => new(name, new XAttribute(W.id, 90),
        new XAttribute(W.author, "Bob"), new XAttribute(W.date, "2031-01-01T00:00:00Z"), content);

    /// <summary>A paragraph whose mark is revised by Bob, and its text too unless it is live.</summary>
    private static XElement RevisedP(string text, XName name, bool liveText = false)
    {
        var run = new XElement(W.r, new XElement(name == W.del && !liveText ? W.delText : W.t, text));
        return new XElement(W.p, new XElement(W.pPr, new XElement(W.rPr, Revised(name))), liveText ? run : Revised(name, run));
    }

    public static TheoryData<string, object[], string, string> PendingNeighbours()
    {
        var gone = RevisedP("gone", W.del);
        var emptied = new XElement(W.p, Revised(W.del, new XElement(W.r, new XElement(W.delText, "text"))));
        XElement Range(XName name, int id) => new(name, new XAttribute(W.id, id), name.LocalName.EndsWith("End") ? null
            : new object[] { new XAttribute(W.author, "Bob"), new XAttribute(W.date, "2031-01-01T00:00:00Z") });
        var inserted = new object[]
        {
            Range(W.customXmlInsRangeStart, 91),
            Control(Range(W.customXmlInsRangeEnd, 91), Table("T2"), P("S"), Range(W.customXmlInsRangeStart, 92)), Range(W.customXmlInsRangeEnd, 92),
        };
        return new()
        {
            // what is left once every change is accepted, and once every one is rejected
            { "a deleted paragraph", new object[] { Table("T1"), gone, P("B"), Table("T2"), P("C") }, "[T1] - [T2] C B", "[T1] gone - [T2] C B" },
            { "one at a control's end", new object[] { Control(Table("T1"), gone), P("B"), Table("T2"), P("C") }, "([T1]) - [T2] C B", "([T1] gone) - [T2] C B" },
            { "one at a control's start", new object[] { Table("T1"), P("B"), Control(gone, Table("T2"), P("S")), P("C") }, "[T1] - ([T2] S) C B", "[T1] - (gone [T2] S) C B" },
            { "an inserted paragraph", new object[] { Table("T1"), RevisedP("new", W.ins), P("B"), Table("T2"), P("C") }, "[T1] new - [T2] C B", "[T1] - [T2] C B" },
            { "an inserted content control", new object[] { Table("T1"), P("B") }.Concat(inserted).Append(P("C")).ToArray(), "[T1] - ([T2] S) C B", "[T1] - [T2] S C B" },
            // these stay whichever way, so they keep the tables apart
            { "an empty paragraph", new object[] { Table("T1"), new XElement(W.p), P("B"), Table("T2"), P("C") }, "[T1] - [T2] C B", "[T1] - [T2] C B" },
            { "a paragraph with its text deleted", new object[] { Table("T1"), emptied, P("B"), Table("T2"), P("C") }, "[T1] - [T2] C B", "[T1] text [T2] C B" },
            { "live text whose mark is deleted", new object[] { Table("T1"), RevisedP("live", W.del, liveText: true), P("B"), Table("T2"), P("C") }, "[T1] live [T2] C B", "[T1] live [T2] C B" },
        };
    }

    // Untracked, the move has to leave the tables apart whichever way the pending changes are resolved.
    [Theory]
    [MemberData(nameof(PendingNeighbours))]
    public void MoveBlock_LeavesTablesApart_HoweverPendingChangesBesideThemAreResolved(
        string scenario, object[] body, string accepted, string rejected)
    {
        Assert.NotEmpty(scenario);
        using var session = new DocxSession(Document(body));

        Assert.True(Move(session, "B", "C", Position.After).Success);

        var saved = session.Save();
        Assert.Equal(accepted, Shape(Accept(saved)));
        Assert.Equal(rejected, Shape(Reject(saved)));
    }

    [Theory]
    [InlineData("a last paragraph of a control that something follows", "B", "T1", Position.Before, "(B [T1]) C")]
    [InlineData("the same, with a blank paragraph after it", "C", "T1", Position.Before, "(C [T1] B)")]
    public void MoveBlock_AddsNoParagraphWhereNoneIsNeeded(string scenario, string source, string target, Position pos, string expected)
    {
        var bytes = scenario.StartsWith("a last") ? Document(Control(Table("T1"), P("B")), P("C")) : Document(Control(Table("T1"), P("B"), P("C")));
        using var untracked = new DocxSession(bytes);
        Assert.Empty(Move(untracked, source, target, pos).Created);
        Assert.Equal(expected, Shape(untracked.Save()));

        using var tracked = Tracked(bytes);
        Assert.Single(Move(tracked, source, target, pos).Created);
        Assert.Equal(expected, Shape(Accept(tracked.Save())));
        Assert.Equal(expected, Shape(ResolveEach(tracked.Save(), accept: true)));
        Assert.Equal(Shape(bytes), Shape(ResolveEach(tracked.Save(), accept: false)));
    }

    [Fact]
    public void MoveBlock_AddsAParagraphOnEachSideOfTheTableThatNeedsOne()
    {
        using var session = new DocxSession(Document(P("A"), Control(Table("T1"), P("B"), Table("T2"))));
        Assert.Equal(2, Move(session, "T1", "T2", Position.After).Created.Count);
        Assert.Equal("A (B [T2] - [T1] -)", Shape(session.Save()));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MoveBlock_PutsTheParagraphAfterTheRangesThatCoverTheTable(bool trackChanges)
    {
        var bytes = Document(P("X"), BookmarkStart(5), P("A"), Table("T1"), P("B"), BookmarkEnd(5), Table("T2"), P("C"));
        using var session = trackChanges ? Tracked(bytes) : new DocxSession(bytes);

        Assert.True(Move(session, "B", "A", Position.Before).Success);

        Assert.Equal("X < B A [T1] > - [T2] C", Shape(trackChanges ? Accept(session.Save()) : session.Save()));
    }

    // The paragraph that keeps the tables apart is part of the move that left it.
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void MoveBlock_Tracked_OneRevisionResolvesTheMoveAndTheParagraphItLeft(bool accept)
    {
        var bytes = Document(Table("T1"), P("B"), Table("T2"), P("C"));
        using var session = Tracked(bytes);
        Assert.True(Move(session, "B", "C", Position.After).Success);

        var revision = Assert.Single(session.ListRevisions());
        Assert.Equal(("move", "B¶"), (revision.Type, revision.Text));
        Assert.True((accept ? session.AcceptRevision(revision.Id) : session.RejectRevision(revision.Id)).Success);

        Assert.Empty(session.ListRevisions());
        Assert.Equal(accept ? "[T1] - [T2] C B" : Shape(bytes), Shape(session.Save()));
    }

    // A blank paragraph someone adds after a moved one is their own insertion.
    [Theory]
    [InlineData("Bob", "2031-01-01T00:00:00Z", true)]
    [InlineData("Alice", "2031-01-01T00:00:01Z", true)]
    [InlineData("Alice", "2031-01-01T00:00:00Z", false)]
    public void ABlankParagraphInsertedAfterAMovedOne_IsItsOwnRevision(string author, string date, bool beforeATable)
    {
        XElement Mark(XName name, int id, string by, string on) =>
            new(name, new XAttribute(W.id, id), new XAttribute(W.author, by), new XAttribute(W.date, on));
        XElement WithRun(XElement mark) { mark.Add(new XElement(W.r, new XElement(W.t, "B"))); return mark; }
        XElement Side(XName content, XName start, XName end, int id) => new(W.p,
            new XElement(W.pPr, new XElement(W.rPr, Mark(content, id, "Alice", "2031-01-01T00:00:00Z"))),
            new XElement(start, new XAttribute(W.id, id + 1), new XAttribute(W.name, "m"), new XAttribute(W.author, "Alice"), new XAttribute(W.date, "2031-01-01T00:00:00Z")),
            WithRun(Mark(content, id + 2, "Alice", "2031-01-01T00:00:00Z")),
            new XElement(end, new XAttribute(W.id, id + 1)));
        var blank = new XElement(W.p, new XElement(W.pPr, new XElement(W.rPr, Mark(W.ins, 80, author, date))));
        using var session = new DocxSession(Document(Table("T1"), Side(W.moveFrom, W.moveFromRangeStart, W.moveFromRangeEnd, 60), blank,
            beforeATable ? Table("T2") : P("D"), Side(W.moveTo, W.moveToRangeStart, W.moveToRangeEnd, 70), P("C")));

        Assert.Equal(new[] { "insert", "move" }, session.ListRevisions().Select(r => r.Type).OrderBy(t => t));
    }

    [Theory]
    [InlineData("T2", "B", Position.Before, "T1", "B", Position.Before)]
    [InlineData("T1", "A", Position.Before, "B", "A", Position.Before)]
    [InlineData("T1", "B", Position.After, "B", "A", Position.Before)]
    [InlineData("B", "A", Position.Before, "T2", "A", Position.After)]
    public void MoveBlock_Tracked_TwoMovesAroundTables_AcceptToTheUntrackedResultAndRejectToTheOriginal(
        string first, string firstTarget, Position firstPos, string second, string secondTarget, Position secondPos)
    {
        var bytes = Document(P("A"), Table("T1"), P("B"), Table("T2"), P("C"));
        using var untracked = new DocxSession(bytes);
        using var tracked = Tracked(bytes);
        foreach (var session in new[] { untracked, tracked })
        {
            var ids = new[] { "A", "T1", "B", "T2", "C" }.ToDictionary(t => t, t => Block(session, t));
            foreach (var (source, target, pos) in new[] { (first, firstTarget, firstPos), (second, secondTarget, secondPos) })
            {
                var result = session.MoveBlock(ids[source], ids[target], pos);
                Assert.True(result.Success, result.Error?.Message);
                if (session == tracked) ids[source] = result.Created[0].Id;
            }
        }
        Assert.Equal(Shape(untracked.Save()), Shape(Accept(tracked.Save())));
        Assert.Equal(Shape(bytes), Shape(Reject(tracked.Save())));
        Assert.Equal(Shape(bytes), Shape(ResolveEach(tracked.Save(), accept: false)));
    }

    [Fact]
    public void MoveBlock_InAControlInCustomXml_SeesTheTableBeforeBoth()
    {
        using var session = new DocxSession(Document(Table("T1"), new XElement(W.customXml, new XAttribute(W.element, "e"),
            new XElement(W.customXmlPr), Control(P("B"), Table("T2"), P("S"))), P("C")));
        var result = session.MoveBlock(session.FindByKind("tbl", "body").Last().Anchor.Id, session.FindByText("B")!.Anchor.Id, Position.Before);
        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal("[T1] {(- [T2] B S)} C", Shape(session.Save()));
    }

    // InsertTable keeps its table apart by the same rule.
    [Fact]
    public void InsertTable_SeesPastPermissionMarkersAndControlEdges()
    {
        using var session = new DocxSession(Document(P("A"), new XElement(W.permStart, new XAttribute(W.id, 8)),
            new XElement(W.permEnd, new XAttribute(W.id, 8)), Table("T1"), P("B"), Control(Table("T2"), P("S")), P("C")));
        foreach (var text in new[] { "A", "B" })
            Assert.True(session.InsertTable(Block(session, text), Position.After, 1, 1).Success);
        Assert.Equal("A [] - < > [T1] B [] - ([T2] S) C", Shape(session.Save()));
    }

    // The paragraph left between the tables is a plain one, not an empty item of the moved one's list.
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MoveBlock_LeavesAPlainParagraphBetweenTables(bool trackChanges)
    {
        var heading = P("A", new XElement(W.pPr, new XElement(W.outlineLvl, new XAttribute(W.val, 0))));
        var bytes = Document(Table("T1"), heading, Table("T2"), P("C"));
        using var session = trackChanges ? Tracked(bytes) : new DocxSession(bytes);

        Assert.True(Move(session, "A", "C", Position.After).Success);

        var saved = trackChanges ? Accept(session.Save()) : session.Save();
        Assert.Equal("[T1] - [T2] C A", Shape(saved));
        Assert.Empty(BodyOf(saved).Elements(W.p).First().Descendants(W.outlineLvl));
        Assert.Single(BodyOf(saved).Descendants(W.outlineLvl));
    }

    // A tracked move leaves a copy on each side, and a marker between blocks could go with only one.
    [Theory]
    [InlineData("end")]
    [InlineData("both")]
    public void MoveBlock_Tracked_RefusesToMoveABlockOutOfItsBodyLevelRange(string markers)
    {
        var start = new XElement(W.commentRangeStart, new XAttribute(W.id, 3));
        var end = new XElement(W.commentRangeEnd, new XAttribute(W.id, 3));
        using var session = Tracked(Document(markers == "both"
            ? new object[] { P("X"), start, P("A"), end, P("B") }
            : new object[] { P("X"), P("A", start), end, P("B") }));

        Assert.False(Move(session, "A", "B", Position.After).Success);
        Assert.True(Move(session, "B", "X", Position.Before).Success);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void MoveBlock_RefusesToCrossASectionBreakInsideATableOrAContentControl(bool inTable)
    {
        var holder = inTable ? Table("T") : Control(P("T"));
        holder.Descendants(W.p).Single().AddFirst(new XElement(W.pPr, new XElement(W.sectPr)));
        using var session = new DocxSession(Document(P("A"), P("B"), holder, P("C")));

        Assert.False(Move(session, "A", "C", Position.After).Success);
        Assert.True(Move(session, "A", "B", Position.After).Success);
        Assert.False(session.MoveBlockRange(Block(session, "B"), Block(session, "T"), Block(session, "C"), Position.After).Success);
    }

    [Fact]
    public void MoveBlockRange_TakesBothMarkersOfTheBookmarkAroundTheRun()
    {
        using var session = new DocxSession(Document(BookmarkStart(3), P("A"), P("B"), BookmarkEnd(3), P("C")));

        var result = session.MoveBlockRange(Block(session, "A"), Block(session, "B"), Block(session, "C"), Position.After);

        Assert.True(result.Success, result.Error?.Message);
        Assert.Equal("C < A B >", Shape(session.Save()));
    }

    public static TheoryData<string, object[]> CrowdedBodies() => new()
    {
        {
            "bookmarks and a permission that end between blocks",
            new object[]
            {
                P("A", BookmarkStart(1)), BookmarkEnd(1), BookmarkStart(2), P("B"), P("C", BookmarkEnd(2)),
                new XElement(W.permStart, new XAttribute(W.id, 7)), P("D"),
                new XElement(W.permEnd, new XAttribute(W.id, 7)), BookmarkStart(4), BookmarkEnd(4), P("E"),
            }
        },
        {
            "nested and overlapping ranges between blocks",
            new object[]
            {
                BookmarkStart(1), P("A"), BookmarkStart(2), P("B"), BookmarkEnd(1), P("C"), BookmarkEnd(2),
                P("D", BookmarkStart(3)), P("E"), BookmarkEnd(3), P("F"),
            }
        },
        {
            "tables, a content control and a permission",
            new object[]
            {
                Table("T1"), P("A"), Table("T2"), new XElement(W.permStart, new XAttribute(W.id, 9)), P("B"), Control(Table("T3"), P("C")),
                new XElement(W.permEnd, new XAttribute(W.id, 9)), P("D"), Table("T4"), BookmarkStart(5), P("E"), BookmarkEnd(5),
            }
        },
        {
            "a field over tables and a bookmark",
            Contents().Concat(new object[] { P("A", BookmarkStart(1)), Table("T1"), BookmarkEnd(1), P("B"), Table("T2"), P("C") }).ToArray()
        },
    };

    /// <summary>Whether some placement beside the target keeps what every range covers: the block alone or
    /// with markers right beside it, on either side of the markers between the target and the next block.</summary>
    private static bool SomePlacementKeepsCoverage(XElement body, int source, int target, Position pos)
    {
        var children = body.Elements().ToList();
        bool Marker(XElement e) => e.Name != W.p && e.Name != W.tbl && e.Name != W.sdt && e.Name != W.sectPr;
        var coverage = Coverage(body);
        for (int lo = source; lo == source || (lo >= 0 && Marker(children[lo])); lo--)
        for (int hi = source; hi == source || (hi < children.Count && Marker(children[hi])); hi++)
        {
            var run = children.GetRange(lo, hi - lo + 1);
            var others = children.Except(run).ToList();
            int slot = others.IndexOf(children[target]) + (pos == Position.Before ? 0 : 1);
            while (true)
            {
                var reordered = others.ToList();
                reordered.InsertRange(slot, run);
                if (Coverage(new XElement(W.body, reordered)).SequenceEqual(coverage)) return true;
                int beside = pos == Position.Before ? slot - 1 : slot;
                if (beside < 0 || beside >= others.Count || !Marker(others[beside])) break;
                slot += pos == Position.Before ? -1 : 1;
            }
        }
        return false;
    }

    // Every move of every block: offered exactly when it is accepted; an accepted one keeps what each
    // range and field covers, leaves the tables apart and undoes to the document it started from; and
    // one is refused only where no placement beside the target would keep what they cover.
    [Theory]
    [MemberData(nameof(CrowdedBodies))]
    public void MoveBlock_EveryMove_KeepsWhatRangesCoverAndTablesApart(string scenario, object[] body)
    {
        Assert.NotEmpty(scenario);
        using var session = new DocxSession(Document(body));
        var ids = session.ListBlocks().Body.Select(u => u.Id).ToArray();
        var original = session.Save();
        var originalBody = BodyOf(original);
        var coverage = Coverage(originalBody);
        // (-1 for a block inside a content control, which the placement check leaves out)
        var position = ids.ToDictionary(id => id, id =>
        {
            var text = Text(XElement.Parse(session.Raw.GetXml(id)));
            return originalBody.Elements().ToList().FindIndex(e => (e.Name == W.p || e.Name == W.tbl) && Text(e) == text);
        });

        foreach (var source in ids)
        foreach (var target in ids.Where(id => id != source))
        foreach (var pos in new[] { Position.Before, Position.After })
        {
            bool offered = IsOffered(session, source, target, pos);
            var result = session.MoveBlock(source, target, pos);
            // A block already in place is a no-op that succeeds even where no move is offered.
            bool inPlace = result.Success && result.Modified.Count == 0;
            if (!inPlace) Assert.Equal(offered, result.Success);
            var saved = session.Save();
            if (result.Success && !inPlace)
            {
                Assert.Equal(coverage, Coverage(BodyOf(saved)));
                Assert.False(TableTouches(BodyOf(saved)), Shape(saved));
                AssertValid(saved);
                Assert.True(session.Undo());
            }
            else if (!inPlace && position[source] >= 0 && position[target] >= 0)
                Assert.False(SomePlacementKeepsCoverage(originalBody, position[source], position[target], pos),
                    $"{source} {pos} {target}: {result.Error!.Message}");
            Assert.Equal(Shape(original), Shape(session.Save()));
        }
    }
}
