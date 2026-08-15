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
}
