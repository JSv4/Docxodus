#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Tests for markup-native revision listing + selective per-revision accept/reject
/// (<see cref="DocxSession.ListRevisions"/>, <see cref="DocxSession.AcceptRevision"/>,
/// <see cref="DocxSession.RejectRevision"/> — issue #318). Test IDs DS370-DS389.
/// </summary>
public class DocxSessionRevisionTests
{
    private static readonly XNamespace Xml = XNamespace.Xml;

    // ─── Fixture builders ─────────────────────────────────────────────────

    /// <summary>Create a minimal package and replace its body with raw revision markup.</summary>
    private static byte[] BuildWithBody(params object[] bodyChildren)
    {
        using var ms = new MemoryStream();
        using (var wDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = wDoc.AddMainDocumentPart();
            main.Document = new Document(new Body());
            var stylesPart = main.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles();
            var settingsPart = main.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings();
            main.Document.Save();

            var xd = main.GetXDocument();
            xd.Root!.Element(W.body)!.ReplaceNodes(bodyChildren);
            main.PutXDocument();
        }
        return ms.ToArray();
    }

    private static XElement Para(params object[] content) => new(W.p, content);

    private static XElement RunT(string text) =>
        new(W.r, new XElement(W.t, new XAttribute(Xml + "space", "preserve"), text));

    private static object[] RevAttrs(int id, string author, string date = "2026-01-02T03:04:05Z") =>
        new object[]
        {
            new XAttribute(W.id, id),
            new XAttribute(W.author, author),
            new XAttribute(W.date, date),
        };

    private static XElement Ins(int id, string author, string text) =>
        new(W.ins, RevAttrs(id, author), RunT(text));

    private static XElement Del(int id, string author, string text) =>
        new(W.del, RevAttrs(id, author),
            new XElement(W.r,
                new XElement(W.delText, new XAttribute(Xml + "space", "preserve"), text)));

    /// <summary>Paragraph-mark revision holder: <c>w:pPr/w:rPr/&lt;markName&gt;</c>.</summary>
    private static XElement MarkPPr(XName markName, int id, string author) =>
        new(W.pPr, new XElement(W.rPr, new XElement(markName, RevAttrs(id, author))));

    /// <summary>The issue's headline shape: one paragraph with an insertion (author A) and
    /// a separate deletion (author B), plus a second paragraph whose whole text is deleted.</summary>
    private static byte[] BuildMixedRevisionsDoc() =>
        BuildWithBody(
            Para(
                RunT("The city of "),
                Ins(101, "Alice", "New York"),
                Del(102, "Bob", "Boston"),
                RunT(" is large.")),
            Para(Del(103, "Bob", "This sentence is gone.")),
            Para(RunT("Untouched tail.")));

    /// <summary>A wholly inserted paragraph (runs + mark, same author) between two clean ones.</summary>
    private static byte[] BuildInsertedParagraphDoc() =>
        BuildWithBody(
            Para(RunT("Before.")),
            Para(
                MarkPPr(W.ins, 201, "Alice"),
                Ins(202, "Alice", "Whole new paragraph.")),
            Para(RunT("After.")));

    /// <summary>Two paragraphs whose join is tracked: the first paragraph's mark is deleted.</summary>
    private static byte[] BuildDeletedParaMarkDoc() =>
        BuildWithBody(
            Para(
                MarkPPr(W.del, 301, "Bob"),
                RunT("Hello ")),
            Para(RunT("world.")));

    /// <summary>A run whose formatting changed (bold added; rPrChange stores the old, empty rPr).</summary>
    private static byte[] BuildFormatChangeDoc() =>
        BuildWithBody(
            Para(
                new XElement(W.r,
                    new XElement(W.rPr,
                        new XElement(W.b),
                        new XElement(W.rPrChange, RevAttrs(401, "Carol"), new XElement(W.rPr))),
                    new XElement(W.t, "Bold now.")),
                RunT(" Plain.")));

    /// <summary>An inline move pair linked by range name "move1".</summary>
    private static byte[] BuildMoveDoc() =>
        BuildWithBody(
            Para(
                RunT("Start "),
                new XElement(W.moveFromRangeStart,
                    new XAttribute(W.id, 500), new XAttribute(W.name, "move1")),
                new XElement(W.moveFrom, RevAttrs(501, "Alice"), RunT("moved bit")),
                new XElement(W.moveFromRangeEnd, new XAttribute(W.id, 500)),
                RunT("end.")),
            Para(
                RunT("Dest "),
                new XElement(W.moveToRangeStart,
                    new XAttribute(W.id, 510), new XAttribute(W.name, "move1")),
                new XElement(W.moveTo, RevAttrs(511, "Alice"), RunT("moved bit")),
                new XElement(W.moveToRangeEnd, new XAttribute(W.id, 510)),
                RunT("here.")));

    /// <summary>A 2x2 table whose second row is row-deleted (trPr/del + del-wrapped cell runs).</summary>
    private static byte[] BuildDeletedRowDoc()
    {
        XElement Cell(params object[] content) =>
            new(W.tc, new XElement(W.tcPr), Para(content));

        return BuildWithBody(
            Para(RunT("Intro.")),
            new XElement(W.tbl,
                new XElement(W.tblPr),
                new XElement(W.tblGrid, new XElement(W.gridCol), new XElement(W.gridCol)),
                new XElement(W.tr, Cell(RunT("A1")), Cell(RunT("B1"))),
                new XElement(W.tr,
                    new XElement(W.trPr, new XElement(W.del, RevAttrs(601, "Bob"))),
                    Cell(MarkPPr(W.del, 602, "Bob"), Del(603, "Bob", "A2")),
                    Cell(MarkPPr(W.del, 604, "Bob"), Del(605, "Bob", "B2")))),
            Para(RunT("Outro.")));
    }

    // ─── Assertion helpers ────────────────────────────────────────────────

    /// <summary>Paragraph texts of the main body, visible (non-deleted) text only.</summary>
    private static string[] ParagraphTexts(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var body = doc.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        return body.Descendants(W.p)
            .Select(p => p.Descendants(W.t).Aggregate("", (acc, t) => acc + t.Value))
            .ToArray();
    }

    private static string VisibleText(byte[] bytes) => string.Join("\n", ParagraphTexts(bytes));

    private static bool HasRevisionMarkup(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var root = doc.MainDocumentPart!.GetXDocument().Root!;
        XName[] markers =
        {
            W.ins, W.del, W.moveFrom, W.moveTo, W.moveFromRangeStart, W.moveFromRangeEnd,
            W.moveToRangeStart, W.moveToRangeEnd, W.rPrChange, W.pPrChange,
        };
        return root.Descendants().Any(d => markers.Contains(d.Name));
    }

    // ─── Listing ──────────────────────────────────────────────────────────

    [Fact]
    public void DS370_ListRevisions_ReadsMarkupIdentityAuthorsAndText()
    {
        using var s = new DocxSession(BuildMixedRevisionsDoc());
        var revs = s.ListRevisions();

        Assert.Equal(3, revs.Count);

        Assert.Equal("rev101", revs[0].Id);
        Assert.Equal("insert", revs[0].Type);
        Assert.Equal("Alice", revs[0].Author);
        Assert.Equal("2026-01-02T03:04:05Z", revs[0].Date);
        Assert.Equal("New York", revs[0].Text);
        Assert.NotNull(revs[0].AnchorId);

        Assert.Equal("rev102", revs[1].Id);
        Assert.Equal("delete", revs[1].Type);
        Assert.Equal("Bob", revs[1].Author);
        Assert.Equal("Boston", revs[1].Text);

        Assert.Equal("rev103", revs[2].Id);
        Assert.Equal("delete", revs[2].Type);
        Assert.Equal("This sentence is gone.", revs[2].Text);
    }

    [Fact]
    public void DS371_ListRevisions_GroupsWhollyInsertedParagraphAsOneRevision()
    {
        using var s = new DocxSession(BuildInsertedParagraphDoc());
        var revs = s.ListRevisions();

        var rev = Assert.Single(revs);
        Assert.Equal("rev201", rev.Id); // min w:id over runs + mark
        Assert.Equal("insert", rev.Type);
        Assert.Equal("Alice", rev.Author);
        Assert.Equal("Whole new paragraph.¶", rev.Text);
    }

    [Fact]
    public void DS372_ListRevisions_MovePairIsOneRevision()
    {
        using var s = new DocxSession(BuildMoveDoc());
        var revs = s.ListRevisions();

        var rev = Assert.Single(revs);
        Assert.Equal("rev500", rev.Id);
        Assert.Equal("move", rev.Type);
        Assert.Equal("moved bit", rev.Text);
    }

    [Fact]
    public void DS373_ListRevisions_DeletedRowIsOneRevision()
    {
        using var s = new DocxSession(BuildDeletedRowDoc());
        var revs = s.ListRevisions();

        var rev = Assert.Single(revs);
        Assert.Equal("rev601", rev.Id);
        Assert.Equal("delete", rev.Type);
        Assert.Equal("Bob", rev.Author);
        Assert.Contains("A2", rev.Text);
        Assert.Contains("B2", rev.Text);
    }

    // ─── Selective resolution (the issue's headline scenario) ─────────────

    [Fact]
    public void DS374_SelectiveAcceptAndReject_MixedResolutionInOneSession()
    {
        using var s = new DocxSession(BuildMixedRevisionsDoc());

        // Accept the insertion, reject the sentence deletion, accept the word deletion.
        Assert.True(s.AcceptRevision("rev101").Success);
        Assert.True(s.RejectRevision("rev103").Success);
        Assert.True(s.AcceptRevision("rev102").Success);

        var bytes = s.Save();
        Assert.Equal(
            new[] { "The city of New York is large.", "This sentence is gone.", "Untouched tail." },
            ParagraphTexts(bytes));
        Assert.False(HasRevisionMarkup(bytes));
    }

    [Fact]
    public void DS375_ResolvingOneRevision_LeavesOtherIdsStable()
    {
        using var s = new DocxSession(BuildMixedRevisionsDoc());
        var before = s.ListRevisions().Select(r => r.Id).ToArray();
        Assert.Equal(new[] { "rev101", "rev102", "rev103" }, before);

        Assert.True(s.AcceptRevision("rev102").Success);

        var after = s.ListRevisions();
        Assert.Equal(new[] { "rev101", "rev103" }, after.Select(r => r.Id).ToArray());
        Assert.Equal("New York", after[0].Text);
    }

    [Fact]
    public void DS376_AcceptInsert_KeepsTextRemovesMarkup_RejectInsert_RemovesText()
    {
        using (var s = new DocxSession(BuildMixedRevisionsDoc()))
        {
            Assert.True(s.AcceptRevision("rev101").Success);
            Assert.StartsWith("The city of New York", ParagraphTexts(s.Save())[0]);
        }

        using (var s = new DocxSession(BuildMixedRevisionsDoc()))
        {
            Assert.True(s.RejectRevision("rev101").Success);
            Assert.Equal("The city of  is large.", ParagraphTexts(s.Save())[0]);
        }
    }

    [Fact]
    public void DS377_RejectDelete_RestoresDeletedTextAsPlainText()
    {
        using var s = new DocxSession(BuildMixedRevisionsDoc());
        Assert.True(s.RejectRevision("rev102").Success);

        var bytes = s.Save();
        // The pending insertion's text still shows; the deletion's text is restored inline.
        Assert.Equal("The city of New YorkBoston is large.", ParagraphTexts(bytes)[0]);
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var p1 = doc.MainDocumentPart!.GetXDocument().Root!.Descendants(W.p).First();
        Assert.Empty(p1.Descendants(W.del));
        Assert.Empty(p1.Descendants(W.delText));
    }

    // ─── Paragraph-mark and structural semantics ──────────────────────────

    [Fact]
    public void DS378_RejectInsertedParagraph_RemovesTheWholeParagraph()
    {
        using var s = new DocxSession(BuildInsertedParagraphDoc());
        var result = s.RejectRevision("rev201");
        Assert.True(result.Success);
        Assert.NotEmpty(result.Removed);

        Assert.Equal(new[] { "Before.", "After." }, ParagraphTexts(s.Save()));
        Assert.False(HasRevisionMarkup(s.Save()));
    }

    [Fact]
    public void DS379_AcceptInsertedParagraph_KeepsItMarkupFree()
    {
        using var s = new DocxSession(BuildInsertedParagraphDoc());
        Assert.True(s.AcceptRevision("rev201").Success);

        Assert.Equal(new[] { "Before.", "Whole new paragraph.", "After." }, ParagraphTexts(s.Save()));
        Assert.False(HasRevisionMarkup(s.Save()));
    }

    [Fact]
    public void DS380_AcceptDeletedParagraphMark_CoalescesIntoFollowingParagraph()
    {
        using var s = new DocxSession(BuildDeletedParaMarkDoc());
        Assert.True(s.AcceptRevision("rev301").Success);

        Assert.Equal(new[] { "Hello world." }, ParagraphTexts(s.Save()));
        Assert.False(HasRevisionMarkup(s.Save()));
    }

    [Fact]
    public void DS381_RejectDeletedParagraphMark_KeepsBothParagraphs()
    {
        using var s = new DocxSession(BuildDeletedParaMarkDoc());
        Assert.True(s.RejectRevision("rev301").Success);

        Assert.Equal(new[] { "Hello ", "world." }, ParagraphTexts(s.Save()));
        Assert.False(HasRevisionMarkup(s.Save()));
    }

    [Fact]
    public void DS382_AcceptMove_MaterializesAtDestination_RejectMove_StaysAtSource()
    {
        using (var s = new DocxSession(BuildMoveDoc()))
        {
            Assert.True(s.AcceptRevision("rev500").Success);
            Assert.Equal(new[] { "Start end.", "Dest moved bithere." }, ParagraphTexts(s.Save()));
            Assert.False(HasRevisionMarkup(s.Save()));
        }

        using (var s = new DocxSession(BuildMoveDoc()))
        {
            Assert.True(s.RejectRevision("rev500").Success);
            Assert.Equal(new[] { "Start moved bitend.", "Dest here." }, ParagraphTexts(s.Save()));
            Assert.False(HasRevisionMarkup(s.Save()));
        }
    }

    [Fact]
    public void DS383_AcceptDeletedRow_RemovesRow_RejectDeletedRow_RestoresIt()
    {
        using (var s = new DocxSession(BuildDeletedRowDoc()))
        {
            var result = s.AcceptRevision("rev601");
            Assert.True(result.Success);
            using var ms = new MemoryStream(s.Save());
            using var doc = WordprocessingDocument.Open(ms, false);
            var tbl = doc.MainDocumentPart!.GetXDocument().Root!.Descendants(W.tbl).Single();
            Assert.Single(tbl.Elements(W.tr));
        }

        using (var s = new DocxSession(BuildDeletedRowDoc()))
        {
            Assert.True(s.RejectRevision("rev601").Success);
            var bytes = s.Save();
            using var ms = new MemoryStream(bytes);
            using var doc = WordprocessingDocument.Open(ms, false);
            var tbl = doc.MainDocumentPart!.GetXDocument().Root!.Descendants(W.tbl).Single();
            Assert.Equal(2, tbl.Elements(W.tr).Count());
            Assert.Contains("A2", VisibleText(bytes));
            Assert.False(HasRevisionMarkup(bytes));
        }
    }

    // ─── Format changes ───────────────────────────────────────────────────

    [Fact]
    public void DS384_FormatChange_ListedAndResolvedBothWays()
    {
        using (var s = new DocxSession(BuildFormatChangeDoc()))
        {
            var rev = Assert.Single(s.ListRevisions());
            Assert.Equal("rev401", rev.Id);
            Assert.Equal("format", rev.Type);
            Assert.Equal("Carol", rev.Author);
            Assert.Equal("Bold now.", rev.Text);

            Assert.True(s.AcceptRevision("rev401").Success);
            using var ms = new MemoryStream(s.Save());
            using var doc = WordprocessingDocument.Open(ms, false);
            var run = doc.MainDocumentPart!.GetXDocument().Root!.Descendants(W.r).First();
            Assert.NotNull(run.Element(W.rPr)?.Element(W.b));      // new formatting kept
            Assert.Empty(run.Descendants(W.rPrChange));
        }

        using (var s = new DocxSession(BuildFormatChangeDoc()))
        {
            Assert.True(s.RejectRevision("rev401").Success);
            using var ms = new MemoryStream(s.Save());
            using var doc = WordprocessingDocument.Open(ms, false);
            var run = doc.MainDocumentPart!.GetXDocument().Root!.Descendants(W.r).First();
            Assert.Null(run.Element(W.rPr)?.Element(W.b));         // old (empty) formatting restored
            Assert.Empty(run.Descendants(W.rPrChange));
        }
    }

    // ─── Parity with whole-document accept/reject ─────────────────────────

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DS385_ResolvingEveryRevisionOneByOne_MatchesRevisionProcessor(bool accept)
    {
        var bytes = BuildMixedRevisionsDoc();
        var oracle = accept
            ? RevisionProcessor.AcceptRevisions(new WmlDocument("t.docx", bytes))
            : RevisionProcessor.RejectRevisions(new WmlDocument("t.docx", bytes));

        using var s = new DocxSession(bytes);
        for (int guard = 0; guard < 20; guard++)
        {
            var revs = s.ListRevisions();
            if (revs.Count == 0) break;
            var result = accept ? s.AcceptRevision(revs[0].Id) : s.RejectRevision(revs[0].Id);
            Assert.True(result.Success);
        }

        Assert.Empty(s.ListRevisions());
        Assert.Equal(VisibleText(oracle.DocumentByteArray), VisibleText(s.Save()));
    }

    // ─── Undo / errors / session-authored markup ──────────────────────────

    [Fact]
    public void DS386_AcceptRevision_IsUndoable()
    {
        using var s = new DocxSession(BuildMixedRevisionsDoc());
        Assert.Equal(3, s.ListRevisions().Count);

        Assert.True(s.AcceptRevision("rev103").Success);
        Assert.Equal(2, s.ListRevisions().Count);

        Assert.True(s.Undo());
        var restored = s.ListRevisions();
        Assert.Equal(new[] { "rev101", "rev102", "rev103" }, restored.Select(r => r.Id).ToArray());

        Assert.True(s.Redo());
        Assert.Equal(2, s.ListRevisions().Count);
    }

    [Fact]
    public void DS387_UnknownRevisionId_FailsWithRevisionNotFound()
    {
        using var s = new DocxSession(BuildMixedRevisionsDoc());
        var result = s.AcceptRevision("rev999");
        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.RevisionNotFound, result.Error!.Code);

        // A failed lookup takes no snapshot — nothing to undo beyond prior state.
        Assert.False(s.Undo());
    }

    [Fact]
    public void DS388_SessionAuthoredTrackedEdit_ListsAndRejectsBackToOriginal()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });

        var anchors = s.Project().AnchorIndex;
        var first = anchors.Values.First(a => a.TextPreview.StartsWith("First"));
        Assert.True(s.ReplaceText(first.Anchor.Id, "Rewritten opening.").Success);

        var revs = s.ListRevisions();
        Assert.Equal(2, revs.Count);
        Assert.Equal("delete", revs[0].Type);
        Assert.Equal("First paragraph.", revs[0].Text);
        Assert.Equal("insert", revs[1].Type);
        Assert.Equal("Rewritten opening.", revs[1].Text);

        // Reject both — the paragraph is back to its original text, markup-free.
        Assert.True(s.RejectRevision(revs[1].Id).Success);
        Assert.True(s.RejectRevision(revs[0].Id).Success);
        Assert.Equal(new[] { "First paragraph.", "Second paragraph." }, ParagraphTexts(s.Save()));
        Assert.False(HasRevisionMarkup(s.Save()));
    }

    [Fact]
    public void DS389_EditResult_ReportsModifiedBlockAnchor()
    {
        using var s = new DocxSession(BuildMixedRevisionsDoc());
        var listed = s.ListRevisions();
        var result = s.AcceptRevision("rev101");
        Assert.True(result.Success);
        var modified = Assert.Single(result.Modified);
        Assert.Equal(listed[0].AnchorId, modified.Id);
    }
}
