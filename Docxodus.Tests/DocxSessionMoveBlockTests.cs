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
}
