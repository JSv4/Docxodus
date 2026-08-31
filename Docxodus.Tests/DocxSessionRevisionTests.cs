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
/// <see cref="DocxSession.RejectRevision"/> — issues #318, #319, and #341). Test IDs DS370-DS417.
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

    /// <summary>Two adjacent runs with distinct original formatting, used to prove
    /// that one tracked ApplyFormat call snapshots each run independently.</summary>
    private static byte[] BuildTwoRunFormatTargetDoc() =>
        BuildWithBody(
            Para(
                new XElement(W.r,
                    new XElement(W.rPr, new XElement(W.i)),
                    new XElement(W.t, "Alpha ")),
                RunT("Beta")));

    private static byte[] BuildBoldRunDoc() =>
        BuildWithBody(
            Para(
                new XElement(W.r,
                    // Lexically different from the bare w:b ApplyFormat writes, but
                    // semantically the same — this must remain an untracked no-op.
                    new XElement(W.rPr,
                        new XElement(W.b, new XAttribute(W.val, "true"))),
                    new XElement(W.t, "Already bold."))));

    private static byte[] BuildSemanticNoOpRunDoc() =>
        BuildWithBody(
            Para(
                new XElement(W.r,
                    // Deliberately noncanonical child order and lexical spellings.
                    // Applying the same semantic values must preserve this exact input
                    // instead of manufacturing a review-pane format change.
                    new XElement(W.rPr,
                        new XElement(W.u),
                        new XElement(W.b, new XAttribute(W.val, "true")),
                        new XElement(W.color, new XAttribute(W.val, "ff00aa")),
                        new XElement(W.sz, new XAttribute(W.val, "022")),
                        new XElement(W.szCs, new XAttribute(W.val, "022"))),
                    new XElement(W.t, "Semantically unchanged."))));

    private static byte[] BuildSemanticOffNoOpRunDoc() =>
        BuildWithBody(
            Para(
                new XElement(W.r,
                    new XElement(W.rPr,
                        new XElement(W.b, new XAttribute(W.val, "0")),
                        new XElement(W.i, new XAttribute(W.val, "false")),
                        new XElement(W.strike, new XAttribute(W.val, "off")),
                        new XElement(W.u, new XAttribute(W.val, "none")),
                        new XElement(W.vertAlign, new XAttribute(W.val, "baseline"))),
                    new XElement(W.t, "Already off."))));

    /// <summary>An inline move pair linked by range name "move1".</summary>
    private static byte[] BuildMoveDoc() =>
        BuildWithBody(
            Para(
                RunT("Start "),
                new XElement(W.moveFromRangeStart,
                    RevAttrs(500, "Alice"), new XAttribute(W.name, "move1")),
                new XElement(W.moveFrom, RevAttrs(501, "Alice"), RunT("moved bit")),
                new XElement(W.moveFromRangeEnd, new XAttribute(W.id, 500)),
                RunT("end.")),
            Para(
                RunT("Dest "),
                new XElement(W.moveToRangeStart,
                    RevAttrs(510, "Alice"), new XAttribute(W.name, "move1")),
                new XElement(W.moveTo, RevAttrs(511, "Alice"), RunT("moved bit")),
                new XElement(W.moveToRangeEnd, new XAttribute(W.id, 510)),
                RunT("here.")));

    private static byte[] BuildParagraphFormatChangeDoc() =>
        BuildWithBody(
            Para(
                new XElement(W.pPr,
                    new XElement(W.pPrChange, RevAttrs(701, "Alice"), new XElement(W.pPr))),
                RunT("Formatted paragraph.")));

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

    private static XElement MainDocumentRoot(byte[] bytes)
    {
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        return new XElement(doc.MainDocumentPart!.GetXDocument().Root!);
    }

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

        Assert.StartsWith("rev2-", revs[0].Id);
        Assert.Equal(new[] { "101" }, revs[0].ConstituentIds);
        Assert.Equal("insert", revs[0].Type);
        Assert.Equal("Alice", revs[0].Author);
        Assert.Equal("2026-01-02T03:04:05Z", revs[0].Date);
        Assert.Equal("New York", revs[0].Text);
        Assert.NotNull(revs[0].AnchorId);

        Assert.StartsWith("rev2-", revs[1].Id);
        Assert.Equal(new[] { "102" }, revs[1].ConstituentIds);
        Assert.Equal("delete", revs[1].Type);
        Assert.Equal("Bob", revs[1].Author);
        Assert.Equal("Boston", revs[1].Text);

        Assert.StartsWith("rev2-", revs[2].Id);
        Assert.Equal(new[] { "103" }, revs[2].ConstituentIds);
        Assert.Equal("delete", revs[2].Type);
        Assert.Equal("This sentence is gone.", revs[2].Text);
    }

    [Fact]
    public void DS371_ListRevisions_GroupsWhollyInsertedParagraphAsOneRevision()
    {
        using var s = new DocxSession(BuildInsertedParagraphDoc());
        var revs = s.ListRevisions();

        var rev = Assert.Single(revs);
        Assert.StartsWith("rev2-", rev.Id);
        Assert.Contains("201", rev.ConstituentIds);
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
        Assert.StartsWith("rev2-", rev.Id);
        Assert.Contains("500", rev.ConstituentIds);
        Assert.Equal("move", rev.Type);
        Assert.Equal("moved bit", rev.Text);
    }

    [Fact]
    public void DS373_ListRevisions_DeletedRowIsOneRevision()
    {
        using var s = new DocxSession(BuildDeletedRowDoc());
        var revs = s.ListRevisions();

        var rev = Assert.Single(revs);
        Assert.StartsWith("rev2-", rev.Id);
        Assert.Contains("601", rev.ConstituentIds);
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
        var before = s.ListRevisions().ToArray();
        Assert.Equal(new[] { "101", "102", "103" },
            before.Select(r => Assert.Single(r.ConstituentIds)).ToArray());

        Assert.True(s.AcceptRevision("rev102").Success);

        var after = s.ListRevisions();
        Assert.Equal(new[] { before[0].Id, before[2].Id }, after.Select(r => r.Id).ToArray());
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
            Assert.StartsWith("rev2-", rev.Id);
            Assert.Equal(new[] { "401" }, rev.ConstituentIds);
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

    // ─── Comments targeted by revision id (issue #341) ──────────────────

    [Theory]
    [InlineData("rev101", "ins")]
    [InlineData("rev102", "del")]
    public void DS410_AddCommentByRevisionId_BracketsExactContentExtent(
        string revisionId, string wrapperLocalName)
    {
        using var s = new DocxSession(BuildMixedRevisionsDoc());

        var result = s.AddCommentToRevision(revisionId, "Reviewer", "Discuss this revision.");
        Assert.True(result.Success, result.Error?.Message);
        Assert.Single(s.ListComments());
        Assert.Equal(new[] { "101", "102", "103" },
            s.ListRevisions().Select(r => Assert.Single(r.ConstituentIds)).ToArray());

        var bytes = s.Save();
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var paragraph = doc.MainDocumentPart!.GetXDocument().Root!.Descendants(W.p).First();
        var children = paragraph.Elements().ToList();
        var revisionIndex = children.FindIndex(e => e.Name.LocalName == wrapperLocalName);
        Assert.True(revisionIndex > 0);
        Assert.Equal(W.commentRangeStart, children[revisionIndex - 1].Name);
        Assert.Equal(W.commentRangeEnd, children[revisionIndex + 1].Name);
        Assert.NotNull(children[revisionIndex + 2].Element(W.commentReference));

        var id = (string?)children[revisionIndex - 1].Attribute(W.id);
        Assert.Equal(id, (string?)children[revisionIndex + 1].Attribute(W.id));
        Assert.Equal(id, (string?)children[revisionIndex + 2].Element(W.commentReference)!.Attribute(W.id));

        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(
                FileFormatVersions.Office2019)
            .Validate(doc)
            .ToList();
        Assert.Empty(errors);
    }

    [Fact]
    public void DS411_AddCommentByRevisionId_FormatRevisionBracketsAffectedRun()
    {
        using var s = new DocxSession(BuildFormatChangeDoc());

        var result = s.AddCommentToRevision("rev401", "Reviewer", "Check this formatting.");
        Assert.True(result.Success, result.Error?.Message);

        var root = MainDocumentRoot(s.Save());
        var paragraph = root.Descendants(W.p).Single();
        var children = paragraph.Elements().ToList();
        var changedRunIndex = children.FindIndex(e => e.Descendants(W.rPrChange).Any());
        Assert.True(changedRunIndex > 0);
        Assert.Equal(W.commentRangeStart, children[changedRunIndex - 1].Name);
        Assert.Equal(W.commentRangeEnd, children[changedRunIndex + 1].Name);
        Assert.NotNull(children[changedRunIndex + 2].Element(W.commentReference));
    }

    [Theory]
    [InlineData("rev101", true, true, "New York")]
    [InlineData("rev101", false, false, "New York")]
    [InlineData("rev102", true, false, "Boston")]
    [InlineData("rev102", false, true, "Boston")]
    public void DS412_ResolvingCommentedContent_PreservesOrCollapsesCommentRange(
        string revisionId, bool accept, bool contentSurvives, string text)
    {
        using var s = new DocxSession(BuildMixedRevisionsDoc());
        Assert.True(s.AddCommentToRevision(revisionId, "Reviewer", "Discuss this.").Success);

        var resolved = accept ? s.AcceptRevision(revisionId) : s.RejectRevision(revisionId);
        Assert.True(resolved.Success, resolved.Error?.Message);
        Assert.Single(s.ListComments());

        var root = MainDocumentRoot(s.Save());
        var paragraph = root.Descendants(W.p).First();
        var children = paragraph.Elements().ToList();
        var startIndex = children.FindIndex(e => e.Name == W.commentRangeStart);
        var endIndex = children.FindIndex(e => e.Name == W.commentRangeEnd);
        Assert.True(startIndex >= 0);
        Assert.True(endIndex > startIndex);
        Assert.NotNull(children[endIndex + 1].Element(W.commentReference));
        if (contentSurvives)
        {
            Assert.Contains(children.Skip(startIndex + 1).Take(endIndex - startIndex - 1),
                e => e.Descendants(W.t).Any(t => t.Value == text));
        }
        else
        {
            Assert.Equal(startIndex + 1, endIndex); // collapsed point at the old insertion site
        }
    }

    [Fact]
    public void DS413_RejectCommentedInsertedParagraph_MovesCollapsedAnchorToSurvivor()
    {
        using var s = new DocxSession(BuildInsertedParagraphDoc());
        Assert.True(s.AddCommentToRevision("rev201", "Reviewer", "Do we need this paragraph?").Success);
        Assert.Contains("201", Assert.Single(s.ListRevisions()).ConstituentIds);
        Assert.True(s.RejectRevision("rev201").Success);
        Assert.Single(s.ListComments());

        var root = MainDocumentRoot(s.Save());
        var paragraph = root.Descendants(W.p).Single(p => p.Descendants(W.t).Any(t => t.Value == "After."));
        var children = paragraph.Elements().ToList();
        var startIndex = children.FindIndex(e => e.Name == W.commentRangeStart);
        var endIndex = children.FindIndex(e => e.Name == W.commentRangeEnd);
        Assert.True(startIndex >= 0);
        Assert.Equal(startIndex + 1, endIndex);
        Assert.NotNull(children[endIndex + 1].Element(W.commentReference));
    }

    [Fact]
    public void DS414_AddCommentByUnknownRevisionId_UsesRevisionNotFoundEnvelope()
    {
        using var s = new DocxSession(BuildMixedRevisionsDoc());
        var result = s.AddCommentToRevision("rev999999", "Reviewer", "Cannot attach.");

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.RevisionNotFound, result.Error!.Code);
        Assert.Empty(s.ListComments());
        Assert.False(s.Undo());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DS415_ResolvingCommentedRow_PreservesCommentAndValidMarkup(bool accept)
    {
        using var s = new DocxSession(BuildDeletedRowDoc());
        Assert.True(s.AddCommentToRevision("rev601", "Reviewer", "Discuss this row.").Success);

        var resolved = accept ? s.AcceptRevision("rev601") : s.RejectRevision("rev601");
        Assert.True(resolved.Success, resolved.Error?.Message);
        Assert.Single(s.ListComments());

        var bytes = s.Save();
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var root = doc.MainDocumentPart!.GetXDocument().Root!;
        var start = Assert.Single(root.Descendants(W.commentRangeStart));
        var end = Assert.Single(root.Descendants(W.commentRangeEnd));
        var reference = Assert.Single(root.Descendants(W.commentReference));
        Assert.Equal((string?)start.Attribute(W.id), (string?)end.Attribute(W.id));
        Assert.Equal((string?)start.Attribute(W.id), (string?)reference.Attribute(W.id));

        if (accept)
        {
            Assert.Same(start.Parent, end.Parent);
            Assert.Same(end.Parent, reference.Parent?.Parent);
            Assert.Empty(start.ElementsAfterSelf().TakeWhile(e => e != end));
        }
        else
        {
            Assert.Contains("A2", VisibleText(bytes));
            Assert.Contains("B2", VisibleText(bytes));
        }

        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(
                FileFormatVersions.Office2019)
            .Validate(doc)
            .ToList();
        Assert.Empty(errors);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DS416_CommentedMove_TargetsDestinationAndSurvivesResolution(bool accept)
    {
        using var s = new DocxSession(BuildMoveDoc());
        Assert.True(s.AddCommentToRevision("rev500", "Reviewer", "Discuss this move.").Success);

        var before = MainDocumentRoot(s.Save());
        var destination = before.Descendants(W.p).Single(p => p.Descendants(W.moveTo).Any());
        var destinationChildren = destination.Elements().ToList();
        var moveIndex = destinationChildren.FindIndex(e => e.Name == W.moveTo);
        Assert.Equal(W.commentRangeStart, destinationChildren[moveIndex - 1].Name);
        Assert.Equal(W.commentRangeEnd, destinationChildren[moveIndex + 1].Name);
        Assert.NotNull(destinationChildren[moveIndex + 2].Element(W.commentReference));

        var resolved = accept ? s.AcceptRevision("rev500") : s.RejectRevision("rev500");
        Assert.True(resolved.Success, resolved.Error?.Message);
        Assert.Single(s.ListComments());

        var bytes = s.Save();
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var root = doc.MainDocumentPart!.GetXDocument().Root!;
        var paragraph = root.Descendants(W.p)
            .Single(p => p.Descendants(W.commentReference).Any());
        var children = paragraph.Elements().ToList();
        var startIndex = children.FindIndex(e => e.Name == W.commentRangeStart);
        var endIndex = children.FindIndex(e => e.Name == W.commentRangeEnd);
        Assert.True(startIndex >= 0);
        Assert.True(endIndex > startIndex);
        Assert.NotNull(children[endIndex + 1].Element(W.commentReference));
        if (accept)
            Assert.Contains(children.Skip(startIndex + 1).Take(endIndex - startIndex - 1),
                e => e.Descendants(W.t).Any(t => t.Value == "moved bit"));
        else
            Assert.Equal(startIndex + 1, endIndex);

        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(
                FileFormatVersions.Office2019)
            .Validate(doc)
            .ToList();
        Assert.Empty(errors);
    }

    [Fact]
    public void DS417_ParagraphPropertyRevision_UsesItsTextBearingParagraph()
    {
        using var s = new DocxSession(BuildParagraphFormatChangeDoc());
        Assert.True(s.AddCommentToRevision("rev701", "Reviewer", "Discuss this format.").Success);
        Assert.True(s.RejectRevision("rev701").Success);
        Assert.Single(s.ListComments());

        var bytes = s.Save();
        using var ms = new MemoryStream(bytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var paragraph = doc.MainDocumentPart!.GetXDocument().Root!.Descendants(W.p).Single();
        var children = paragraph.Elements().ToList();
        var startIndex = children.FindIndex(e => e.Name == W.commentRangeStart);
        var endIndex = children.FindIndex(e => e.Name == W.commentRangeEnd);
        Assert.True(startIndex >= 0);
        Assert.True(endIndex > startIndex);
        Assert.Contains(children.Skip(startIndex + 1).Take(endIndex - startIndex - 1),
            e => e.Descendants(W.t).Any(t => t.Value == "Formatted paragraph."));
        Assert.NotNull(children[endIndex + 1].Element(W.commentReference));

        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(
                FileFormatVersions.Office2019)
            .Validate(doc)
            .ToList();
        Assert.Empty(errors);
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
        Assert.Equal(new[] { "101", "102", "103" },
            restored.Select(r => Assert.Single(r.ConstituentIds)).ToArray());

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

    // ─── Session-authored run-format changes (issue #319) ────────────────

    [Fact]
    public void DS390_ApplyFormat_Tracked_EmitsPerRunSnapshotsAndSessionStamp()
    {
        using var s = new DocxSession(BuildTwoRunFormatTargetDoc(),
            new DocxSessionSettings
            {
                TrackedChanges = TrackedChangeMode.RenderInline,
                RevisionAuthor = "Format Reviewer",
            });
        var anchor = s.Project().AnchorIndex.Values.Single().Anchor.Id;

        var result = s.ApplyFormat(anchor, span: null, new FormatOp { Bold = true });
        Assert.True(result.Success, result.Error?.Message);

        // Adjacent per-run markers surface as one user-visible format revision.
        var listed = Assert.Single(s.ListRevisions());
        Assert.StartsWith("rev2-", listed.Id);
        Assert.Equal(new[] { "1001", "1002" }, listed.ConstituentIds);
        Assert.Equal("format", listed.Type);
        Assert.Equal("Format Reviewer", listed.Author);
        Assert.Equal("Alpha Beta", listed.Text);
        Assert.Matches(@"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{7}Z$", listed.Date!);

        var bytes = s.Save();
        using (var ms = new MemoryStream(bytes))
        using (var doc = WordprocessingDocument.Open(ms, false))
        {
            var runs = doc.MainDocumentPart!.GetXDocument().Root!
                .Descendants(W.r).Where(r => r.Element(W.t) is not null).ToArray();
            Assert.Equal(2, runs.Length);
            var changes = runs.Select(r => Assert.Single(r.Element(W.rPr)!.Elements(W.rPrChange)))
                .ToArray();

            Assert.Equal(new[] { "1001", "1002" },
                changes.Select(c => (string?)c.Attribute(W.id)).ToArray());
            Assert.All(changes, c => Assert.Equal("Format Reviewer", (string?)c.Attribute(W.author)));
            Assert.Single(changes.Select(c => (string?)c.Attribute(W.date)).Distinct());
            Assert.All(runs, r => Assert.Equal(W.rPrChange, r.Element(W.rPr)!.Elements().Last().Name));
            Assert.All(runs, r => Assert.NotNull(r.Element(W.rPr)!.Element(W.b)));

            // Each marker archives that run's own old property set.
            Assert.NotNull(changes[0].Element(W.rPr)!.Element(W.i));
            Assert.Null(changes[0].Element(W.rPr)!.Element(W.b));
            Assert.Empty(changes[1].Element(W.rPr)!.Elements());
            Assert.Empty(changes.SelectMany(c => c.Element(W.rPr)!.Descendants(W.rPrChange)));

            var schemaErrors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(
                    FileFormatVersions.Office2019)
                .Validate(doc)
                .ToList();
            Assert.Empty(schemaErrors);
        }

        Assert.True(s.AcceptRevision(listed.Id).Success);
        using var acceptedMs = new MemoryStream(s.Save());
        using var acceptedDoc = WordprocessingDocument.Open(acceptedMs, false);
        var acceptedRuns = acceptedDoc.MainDocumentPart!.GetXDocument().Root!
            .Descendants(W.r).Where(r => r.Element(W.t) is not null).ToArray();
        Assert.All(acceptedRuns, r => Assert.NotNull(r.Element(W.rPr)?.Element(W.b)));
        Assert.NotNull(acceptedRuns[0].Element(W.rPr)?.Element(W.i));
        Assert.Empty(acceptedRuns.SelectMany(r => r.Descendants(W.rPrChange)));
    }

    [Fact]
    public void DS391_ApplyFormat_Tracked_RejectRestoresEachRunsOriginalProperties()
    {
        using var s = new DocxSession(BuildTwoRunFormatTargetDoc(),
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var anchor = s.Project().AnchorIndex.Values.Single().Anchor.Id;
        Assert.True(s.ApplyFormat(anchor, span: null, new FormatOp { Bold = true }).Success);

        var revision = Assert.Single(s.ListRevisions());
        Assert.True(s.RejectRevision(revision.Id).Success);

        using var ms = new MemoryStream(s.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        var runs = doc.MainDocumentPart!.GetXDocument().Root!
            .Descendants(W.r).Where(r => r.Element(W.t) is not null).ToArray();
        Assert.Equal(2, runs.Length);
        Assert.Null(runs[0].Element(W.rPr)?.Element(W.b));
        Assert.NotNull(runs[0].Element(W.rPr)?.Element(W.i));
        Assert.Null(runs[1].Element(W.rPr)?.Element(W.b));
        Assert.Empty(runs.SelectMany(r => r.Descendants(W.rPrChange)));
    }

    [Fact]
    public void DS392_ApplyFormatToSubstring_Tracked_RoundTripsThroughRevisionProcessor()
    {
        using var s = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs(),
            new DocxSessionSettings
            {
                TrackedChanges = TrackedChangeMode.RenderInline,
                RevisionAuthor = "Substring Reviewer",
            });
        var anchor = s.Project().AnchorIndex.Values.First().Anchor.Id;

        var result = s.ApplyFormatToSubstring(anchor, "paragraph", new FormatOp { Bold = true });
        Assert.True(result.Success, result.Error?.Message);
        var listed = Assert.Single(s.ListRevisions());
        Assert.Equal("format", listed.Type);
        Assert.Equal("paragraph", listed.Text);
        Assert.Equal("Substring Reviewer", listed.Author);

        var tracked = s.Save();
        using (var trackedMs = new MemoryStream(tracked))
        using (var trackedDoc = WordprocessingDocument.Open(trackedMs, false))
        {
            var changedRun = trackedDoc.MainDocumentPart!.GetXDocument().Root!
                .Descendants(W.r).Single(r => (string?)r.Element(W.t) == "paragraph");
            Assert.NotNull(changedRun.Element(W.rPr)?.Element(W.b));
            Assert.NotNull(changedRun.Element(W.rPr)?.Element(W.rPrChange));
            Assert.Empty(trackedDoc.MainDocumentPart.GetXDocument().Root!
                .Descendants(W.rPrChange)
                .Where(c => (string?)c.Parent?.Parent?.Element(W.t) != "paragraph"));
        }

        var accepted = RevisionProcessor.AcceptRevisions(new WmlDocument("accepted.docx", tracked));
        var rejected = RevisionProcessor.RejectRevisions(new WmlDocument("rejected.docx", tracked));

        using (var acceptedMs = new MemoryStream(accepted.DocumentByteArray))
        using (var acceptedDoc = WordprocessingDocument.Open(acceptedMs, false))
        {
            var run = acceptedDoc.MainDocumentPart!.GetXDocument().Root!
                .Descendants(W.r).Single(r => (string?)r.Element(W.t) == "paragraph");
            Assert.NotNull(run.Element(W.rPr)?.Element(W.b));
            Assert.Empty(run.Descendants(W.rPrChange));
        }
        using (var rejectedMs = new MemoryStream(rejected.DocumentByteArray))
        using (var rejectedDoc = WordprocessingDocument.Open(rejectedMs, false))
        {
            var root = rejectedDoc.MainDocumentPart!.GetXDocument().Root!;
            var run = root.Descendants(W.r).Single(r => (string?)r.Element(W.t) == "paragraph");
            Assert.Null(run.Element(W.rPr)?.Element(W.b));
            Assert.Empty(root.Descendants(W.rPrChange));
            Assert.Contains("First paragraph.", string.Concat(root.Descendants(W.t).Select(t => t.Value)));
        }
    }

    [Fact]
    public void DS393_ApplyFormat_Tracked_PreservesExistingRevisionBaselineAndMetadata()
    {
        using var s = new DocxSession(BuildFormatChangeDoc(),
            new DocxSessionSettings
            {
                TrackedChanges = TrackedChangeMode.RenderInline,
                RevisionAuthor = "Later Reviewer",
            });
        var anchor = s.Project().AnchorIndex.Values.Single().Anchor.Id;

        var result = s.ApplyFormat(anchor, new CharSpan(0, "Bold now.".Length),
            new FormatOp { Italic = true });
        Assert.True(result.Success, result.Error?.Message);

        // OOXML allows one rPrChange only. The later edit folds into Carol's pending
        // change so reject still reaches the original empty formatting baseline.
        var listed = Assert.Single(s.ListRevisions());
        Assert.StartsWith("rev2-", listed.Id);
        Assert.Equal(new[] { "401" }, listed.ConstituentIds);
        Assert.Equal("Carol", listed.Author);

        using (var trackedMs = new MemoryStream(s.Save()))
        using (var trackedDoc = WordprocessingDocument.Open(trackedMs, false))
        {
            var run = trackedDoc.MainDocumentPart!.GetXDocument().Root!
                .Descendants(W.r).First();
            var rPr = run.Element(W.rPr)!;
            Assert.NotNull(rPr.Element(W.b));
            Assert.NotNull(rPr.Element(W.i));
            var change = Assert.Single(rPr.Elements(W.rPrChange));
            Assert.Equal("401", (string?)change.Attribute(W.id));
            Assert.Equal("Carol", (string?)change.Attribute(W.author));
            Assert.Empty(change.Element(W.rPr)!.Elements());
            Assert.Empty(change.Element(W.rPr)!.Descendants(W.rPrChange));
            Assert.Equal(W.rPrChange, rPr.Elements().Last().Name);
        }

        Assert.True(s.RejectRevision("rev401").Success);
        using var rejectedMs = new MemoryStream(s.Save());
        using var rejectedDoc = WordprocessingDocument.Open(rejectedMs, false);
        var rejectedRun = rejectedDoc.MainDocumentPart!.GetXDocument().Root!
            .Descendants(W.r).First();
        Assert.Null(rejectedRun.Element(W.rPr)?.Element(W.b));
        Assert.Null(rejectedRun.Element(W.rPr)?.Element(W.i));
        Assert.Empty(rejectedRun.Descendants(W.rPrChange));
    }

    [Fact]
    public void DS394_ApplyFormat_Tracked_NoOpDoesNotCreateRevision()
    {
        using var s = new DocxSession(BuildBoldRunDoc(),
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var anchor = s.Project().AnchorIndex.Values.Single().Anchor.Id;

        var result = s.ApplyFormat(anchor, span: null, new FormatOp { Bold = true });
        Assert.True(result.Success, result.Error?.Message);
        Assert.Empty(s.ListRevisions());

        using var ms = new MemoryStream(s.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        var run = doc.MainDocumentPart!.GetXDocument().Root!.Descendants(W.r).Single();
        Assert.NotNull(run.Element(W.rPr)?.Element(W.b));
        Assert.Empty(run.Descendants(W.rPrChange));
    }

    [Fact]
    public void DS395_ApplyFormat_Tracked_FailureRestoresWholeOperationSnapshot()
    {
        using var s = new DocxSession(BuildFormatChangeDoc(),
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var anchor = s.Project().AnchorIndex.Values.Single().Anchor.Id;
        var before = MainDocumentRoot(s.Save());

        // The partial span first splits a run carrying an existing rPrChange. The
        // invalid value then fails after another property was tentatively changed.
        // Run-local restoration is insufficient: the operation snapshot must also
        // undo the split and restore the original marker exactly.
        var result = s.ApplyFormat(anchor, new CharSpan(1, 3),
            new FormatOp { Italic = true, VertAlign = "sideways" });

        Assert.False(result.Success);
        Assert.Equal(EditErrorCode.InternalError, result.Error!.Code);
        Assert.True(XNode.DeepEquals(before, MainDocumentRoot(s.Save())));
        var revision = Assert.Single(s.ListRevisions());
        Assert.StartsWith("rev2-", revision.Id);
        Assert.Equal(new[] { "401" }, revision.ConstituentIds);
        Assert.Equal("Carol", revision.Author);
        Assert.False(s.Undo()); // failed operations do not remain on the history stack
    }

    [Fact]
    public void DS396_ApplyFormat_Tracked_SeparateAdjacentCallsResolveIndependently()
    {
        using var s = new DocxSession(BuildTwoRunFormatTargetDoc(),
            new DocxSessionSettings
            {
                TrackedChanges = TrackedChangeMode.RenderInline,
                RevisionAuthor = "One Reviewer",
            });
        var anchor = s.Project().AnchorIndex.Values.Single().Anchor.Id;

        Assert.True(s.ApplyFormat(anchor, new CharSpan(0, 6),
            new FormatOp { Bold = true }).Success);
        Assert.True(s.ApplyFormat(anchor, new CharSpan(6, 4),
            new FormatOp { Italic = true }).Success);

        var revisions = s.ListRevisions();
        Assert.Equal(2, revisions.Count);
        Assert.All(revisions, r => Assert.StartsWith("rev2-", r.Id));
        Assert.Equal(new[] { "1001", "1002" },
            revisions.Select(r => Assert.Single(r.ConstituentIds)).ToArray());
        Assert.Equal(new[] { "Alpha ", "Beta" }, revisions.Select(r => r.Text).ToArray());
        Assert.All(revisions, r => Assert.Equal("One Reviewer", r.Author));
        Assert.NotEqual(revisions[0].Date, revisions[1].Date);

        Assert.True(s.RejectRevision(revisions[0].Id).Success);
        var remaining = Assert.Single(s.ListRevisions());
        Assert.Equal(new[] { "1002" }, remaining.ConstituentIds);
        Assert.Equal("Beta", remaining.Text);

        using (var rejectedFirstMs = new MemoryStream(s.Save()))
        using (var rejectedFirstDoc = WordprocessingDocument.Open(rejectedFirstMs, false))
        {
            var runs = rejectedFirstDoc.MainDocumentPart!.GetXDocument().Root!
                .Descendants(W.r).Where(r => r.Element(W.t) is not null).ToArray();
            Assert.Null(runs[0].Element(W.rPr)?.Element(W.b));
            Assert.NotNull(runs[0].Element(W.rPr)?.Element(W.i));
            Assert.NotNull(runs[1].Element(W.rPr)?.Element(W.i));
            Assert.Empty(runs[0].Descendants(W.rPrChange));
            Assert.Single(runs[1].Descendants(W.rPrChange));
        }

        Assert.True(s.AcceptRevision(remaining.Id).Success);
        Assert.Empty(s.ListRevisions());
    }

    [Fact]
    public void DS397_ApplyFormat_Tracked_RevertingWholePendingChangeDropsMarker()
    {
        using var s = new DocxSession(BuildFormatChangeDoc(),
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var anchor = s.Project().AnchorIndex.Values.Single().Anchor.Id;

        var result = s.ApplyFormat(anchor, new CharSpan(0, "Bold now.".Length),
            new FormatOp { Bold = false });
        Assert.True(result.Success, result.Error?.Message);
        Assert.Empty(s.ListRevisions());

        using (var ms = new MemoryStream(s.Save()))
        using (var doc = WordprocessingDocument.Open(ms, false))
        {
            var run = doc.MainDocumentPart!.GetXDocument().Root!
                .Descendants(W.r).First(r => (string?)r.Element(W.t) == "Bold now.");
            Assert.Null(run.Element(W.rPr)?.Element(W.b));
            Assert.Empty(run.Descendants(W.rPrChange));
        }

        Assert.True(s.Undo());
        Assert.Equal(new[] { "401" }, Assert.Single(s.ListRevisions()).ConstituentIds);
        Assert.True(s.Redo());
        Assert.Empty(s.ListRevisions());
    }

    [Fact]
    public void DS398_ApplyFormat_Tracked_RevertingPartialPendingChangeKeepsOnlyRemainder()
    {
        using var s = new DocxSession(BuildFormatChangeDoc(),
            new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
        var anchor = s.Project().AnchorIndex.Values.Single().Anchor.Id;

        var result = s.ApplyFormat(anchor, new CharSpan(0, 4),
            new FormatOp { Bold = false });
        Assert.True(result.Success, result.Error?.Message);
        var remaining = Assert.Single(s.ListRevisions());
        Assert.Equal(new[] { "401" }, remaining.ConstituentIds);
        Assert.Equal(" now.", remaining.Text);

        using (var ms = new MemoryStream(s.Save()))
        using (var doc = WordprocessingDocument.Open(ms, false))
        {
            var runs = doc.MainDocumentPart!.GetXDocument().Root!
                .Descendants(W.r).Where(r => r.Element(W.t) is not null).ToArray();
            var reverted = Assert.Single(runs, r => (string?)r.Element(W.t) == "Bold");
            var pending = Assert.Single(runs, r => (string?)r.Element(W.t) == " now.");
            Assert.Null(reverted.Element(W.rPr)?.Element(W.b));
            Assert.Empty(reverted.Descendants(W.rPrChange));
            Assert.NotNull(pending.Element(W.rPr)?.Element(W.b));
            Assert.Single(pending.Descendants(W.rPrChange));
        }

        Assert.True(s.RejectRevision(remaining.Id).Success);
        Assert.Empty(s.ListRevisions());
        using var rejectedMs = new MemoryStream(s.Save());
        using var rejectedDoc = WordprocessingDocument.Open(rejectedMs, false);
        var rejectedRuns = rejectedDoc.MainDocumentPart!.GetXDocument().Root!
            .Descendants(W.r).Where(r => r.Element(W.t) is not null).ToArray();
        Assert.Empty(rejectedRuns.SelectMany(r => r.Descendants(W.rPrChange)));
        Assert.DoesNotContain(rejectedRuns, r => r.Element(W.rPr)?.Element(W.b) is not null);
    }

    [Fact]
    public void DS399_ApplyFormat_Tracked_SemanticNoOpPreservesRawPropertyXml()
    {
        {
            using var s = new DocxSession(BuildSemanticNoOpRunDoc(),
                new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
            var anchor = s.Project().AnchorIndex.Values.Single().Anchor.Id;
            var before = MainDocumentRoot(s.Save());

            var result = s.ApplyFormat(anchor, span: null, new FormatOp
            {
                Bold = true,
                Underline = true,
                Color = "FF00AA",
                FontSizePts = 11,
            });

            Assert.True(result.Success, result.Error?.Message);
            Assert.Empty(s.ListRevisions());
            Assert.True(XNode.DeepEquals(before, MainDocumentRoot(s.Save())));
        }

        {
            using var s = new DocxSession(BuildSemanticOffNoOpRunDoc(),
                new DocxSessionSettings { TrackedChanges = TrackedChangeMode.RenderInline });
            var anchor = s.Project().AnchorIndex.Values.Single().Anchor.Id;
            var before = MainDocumentRoot(s.Save());

            var result = s.ApplyFormat(anchor, span: null, new FormatOp
            {
                Bold = false,
                Italic = false,
                Strike = false,
                Underline = false,
                VertAlign = "baseline",
            });

            Assert.True(result.Success, result.Error?.Message);
            Assert.Empty(s.ListRevisions());
            Assert.True(XNode.DeepEquals(before, MainDocumentRoot(s.Save())));
        }
    }

    // ─── Note cleanup on resolution (issue #516) ─────────────────────────

    /// <summary>
    /// Issue #516: resolving away a document's first-and-only footnote/endnote must delete
    /// the note itself, not leave it as an empty reference-less husk in the notes part.
    /// Word removes a note when the resolution that carries its reference away lands. The
    /// PART deliberately stays — Word never prunes a notes part (the RP050 oracle keeps its
    /// separator-only footnotes.xml after accepting the only note's deletion) — so the
    /// resolved part must hold nothing but separator definitions.
    /// </summary>
    [Theory]
    [InlineData("WC035-Footnote-Before.docx", "WC035-Footnote-After.docx", "footnote")]
    [InlineData("WC035-Footnote-After.docx", "WC035-Footnote-Before.docx", "footnote")]
    [InlineData("WC035-Endnote-Before.docx", "WC035-Endnote-After.docx", "endnote")]
    [InlineData("WC035-Endnote-After.docx", "WC035-Endnote-Before.docx", "endnote")]
    public void DS418_ResolvingAwayTheOnlyNote_DeletesTheNoteNotJustItsContent(
        string leftName, string rightName, string kind)
    {
        var left = File.ReadAllBytes(Path.Combine("../../../../TestFiles/WC", leftName));
        var right = File.ReadAllBytes(Path.Combine("../../../../TestFiles/WC", rightName));
        var redline = DocxDiff.Compare(
            new WmlDocument(leftName, left), new WmlDocument(rightName, right)).DocumentByteArray;

        // Resolve toward the endpoint WITHOUT the note: reject lands on left, accept on right.
        bool accept = HasNotesPart(left, kind);

        using var session = new DocxSession(redline);
        for (var revisions = session.ListRevisions(); revisions.Count > 0; revisions = session.ListRevisions())
        {
            var edit = accept
                ? session.AcceptRevision(revisions[0].Id)
                : session.RejectRevision(revisions[0].Id);
            Assert.True(edit.Success, edit.Error?.Message);
        }

        var resolved = session.Save();
        using var stream = new MemoryStream(resolved, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        Assert.Empty(main.Document.Descendants<FootnoteReference>());
        Assert.Empty(main.Document.Descendants<EndnoteReference>());
        if (kind == "footnote")
        {
            var husks = main.FootnotesPart!.Footnotes!.Elements<Footnote>()
                .Where(note => note.Type is null
                    || note.Type == FootnoteEndnoteValues.Normal).ToArray();
            Assert.Empty(husks);
        }
        else
        {
            var husks = main.EndnotesPart!.Endnotes!.Elements<Endnote>()
                .Where(note => note.Type is null
                    || note.Type == FootnoteEndnoteValues.Normal).ToArray();
            Assert.Empty(husks);
        }
    }

    /// <summary>
    /// The cleanup in DS418 is scoped to notes whose reference the resolution itself removed:
    /// a note that was ALREADY dangling before any resolution (its id referenced nowhere) is
    /// pre-existing document state, and resolving an unrelated revision must not garbage-collect
    /// it — that would be silent content loss.
    /// </summary>
    [Fact]
    public void DS419_ResolvingUnrelatedRevision_LeavesPreExistingDanglingNoteAlone()
    {
        var bytes = BuildWithBody(
            new XElement(W.p,
                new XElement(W.r, new XElement(W.t, "Kept. ")),
                new XElement(W.ins,
                    new XAttribute(W.id, "101"),
                    new XAttribute(W.author, "Alice"),
                    new XAttribute(W.date, "2026-01-01T00:00:00Z"),
                    new XElement(W.r, new XElement(W.t, "Inserted.")))));
        bytes = AddDanglingFootnote(bytes, noteId: 7, text: "Orphan note body.");

        using var session = new DocxSession(bytes);
        var revision = Assert.Single(session.ListRevisions());
        Assert.True(session.RejectRevision(revision.Id).Success);

        var resolved = session.Save();
        Assert.True(HasNotesPart(resolved, "footnote"), "pre-existing notes part was deleted");
        using var stream = new MemoryStream(resolved, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);
        var texts = document.MainDocumentPart!.FootnotesPart!.Footnotes!
            .Elements<Footnote>().Select(note => note.InnerText).ToArray();
        Assert.Contains("Orphan note body.", string.Concat(texts));
    }

    /// <summary>
    /// The same rule through the STATELESS resolver — the path every non-.NET transport reaches
    /// through <c>DocxDiffOps</c>. Rejecting a redline reproduces the baseline, so a note the
    /// redline introduced has to go with the citation it introduced; before #614 the citation
    /// vanished and the definition stayed, and the "rejected" package still shipped the note.
    /// </summary>
    [Fact]
    public void DS429_StatelessReject_TakesAnIntroducedNoteWithItsCitation()
    {
        var left = File.ReadAllBytes(Path.Combine("../../../../TestFiles/WC", "WC035-Footnote-Before.docx"));
        var right = File.ReadAllBytes(Path.Combine("../../../../TestFiles/WC", "WC035-Footnote-After.docx"));

        // Compare in the direction that INSERTS the note, so rejecting must take it back out.
        var inserting = UserNoteCount(left, "footnote") < UserNoteCount(right, "footnote")
            ? (Baseline: left, Counterpart: right)
            : (Baseline: right, Counterpart: left);
        Assert.True(UserNoteCount(inserting.Counterpart, "footnote")
            > UserNoteCount(inserting.Baseline, "footnote"), "fixture no longer inserts a note");

        var redline = DocxDiff.Compare(
            new WmlDocument("baseline.docx", inserting.Baseline),
            new WmlDocument("counterpart.docx", inserting.Counterpart)).DocumentByteArray;

        var rejected = Docxodus.Internal.DocxDiffOps.RejectRevisions(redline);

        Assert.True(HasNotesPart(rejected, "footnote"), "the notes part itself must survive");
        Assert.Equal(UserNoteCount(inserting.Baseline, "footnote"), UserNoteCount(rejected, "footnote"));
    }

    /// <summary>
    /// …and the accept side deliberately does NOT apply it. Accept reproduces the counterpart
    /// document as the comparison saw it — <c>Accept(Compare(l, r)) == r</c> — so a counterpart that
    /// carries a reference-less note definition keeps it. <see cref="DocxSession"/>'s own resolve
    /// paths (DS418) apply the rule in both directions because an editor is authoring a document
    /// rather than inverting a comparison; this test pins that the two contracts stay distinct.
    /// </summary>
    [Fact]
    public void DS430_StatelessAccept_PreservesACounterpartsOwnOrphanedNote()
    {
        var baseline = BuildWithBody(
            new XElement(W.p,
                new XElement(W.r,
                    new XElement(W.t, "Cited here."),
                    new XElement(W.footnoteReference, new XAttribute(W.id, "7")))),
            new XElement(W.p, new XElement(W.r, new XElement(W.t, "Sentinel."))));
        baseline = AddDanglingFootnote(baseline, noteId: 7, text: "Kept as a husk.");

        // The counterpart drops the citing paragraph but keeps the definition — the husk shape a
        // naive edit leaves behind, and exactly what accept has to reproduce.
        var counterpart = BuildWithBody(
            new XElement(W.p, new XElement(W.r, new XElement(W.t, "Sentinel."))));
        counterpart = AddDanglingFootnote(counterpart, noteId: 7, text: "Kept as a husk.");

        var redline = DocxDiff.Compare(
            new WmlDocument("baseline.docx", baseline),
            new WmlDocument("counterpart.docx", counterpart)).DocumentByteArray;

        var accepted = Docxodus.Internal.DocxDiffOps.AcceptRevisions(redline);

        Assert.Equal(1, UserNoteCount(counterpart, "footnote"));
        Assert.Equal(1, UserNoteCount(accepted, "footnote"));
    }

    /// <summary>
    /// The prune asks the whole package who still cites a note, not just the body. A note cited
    /// from the body AND a running header outlives the body citation; a body-only scan would read
    /// "referenced before, unreferenced after" and delete a note the header still points at.
    /// </summary>
    [Fact]
    public void DS431_NoteStillCitedFromAHeader_SurvivesLosingItsBodyCitation()
    {
        var bytes = BuildWithBody(
            new XElement(W.p,
                new XElement(W.r,
                    new XElement(W.t, "Cited here."),
                    new XElement(W.footnoteReference, new XAttribute(W.id, "7")))),
            new XElement(W.p, new XElement(W.r, new XElement(W.t, "Sentinel."))));
        bytes = AddDanglingFootnote(bytes, noteId: 7, text: "Cited twice.");
        bytes = AddHeaderCiting(bytes, noteId: 7);

        using var session = new DocxSession(bytes);
        var citing = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && (t.TextPreview ?? string.Empty).StartsWith("Cited"))
            .Anchor.Id;
        Assert.True(session.DeleteBlock(citing).Success);

        var saved = session.Save();
        Assert.Equal(1, UserNoteCount(saved, "footnote"));
    }

    private static int UserNoteCount(byte[] bytes, string kind)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);
        var main = document.MainDocumentPart!;
        return kind == "footnote"
            ? main.FootnotesPart?.Footnotes?.Elements<Footnote>().Count(n => n.Type is null) ?? 0
            : main.EndnotesPart?.Endnotes?.Elements<Endnote>().Count(n => n.Type is null) ?? 0;
    }

    /// <summary>Add a default running header that cites <paramref name="noteId"/>, so the note has a
    /// second citation outside the body.</summary>
    private static byte[] AddHeaderCiting(byte[] source, int noteId)
    {
        using var ms = new MemoryStream();
        ms.Write(source);
        ms.Position = 0;
        using (var wDoc = WordprocessingDocument.Open(ms, true))
        {
            var main = wDoc.MainDocumentPart!;
            var header = main.AddNewPart<HeaderPart>();
            header.Header = new Header(
                new Paragraph(new Run(
                    new Text("Running head."),
                    new FootnoteReference { Id = noteId })));
            header.Header.Save();

            var body = main.Document!.Body!;
            var sectPr = body.Elements<SectionProperties>().FirstOrDefault();
            if (sectPr is null)
            {
                sectPr = new SectionProperties();
                body.Append(sectPr);
            }

            sectPr.PrependChild(new HeaderReference
            {
                Type = HeaderFooterValues.Default,
                Id = main.GetIdOfPart(header),
            });
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static bool HasNotesPart(byte[] bytes, string kind)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = WordprocessingDocument.Open(stream, false);
        return kind == "footnote"
            ? document.MainDocumentPart!.FootnotesPart is not null
            : document.MainDocumentPart!.EndnotesPart is not null;
    }

    /// <summary>Add a footnotes part holding separator stubs plus one real note that nothing
    /// in the body references — the pre-existing dangling-note shape DS419 protects.</summary>
    private static byte[] AddDanglingFootnote(byte[] source, int noteId, string text)
    {
        using var ms = new MemoryStream();
        ms.Write(source);
        ms.Position = 0;
        using (var wDoc = WordprocessingDocument.Open(ms, true))
        {
            var part = wDoc.MainDocumentPart!.AddNewPart<FootnotesPart>();
            part.Footnotes = new Footnotes(
                new Footnote(new Paragraph(new Run(new SeparatorMark())))
                {
                    Type = FootnoteEndnoteValues.Separator,
                    Id = -1,
                },
                new Footnote(new Paragraph(new Run(new SeparatorMark())))
                {
                    Type = FootnoteEndnoteValues.ContinuationSeparator,
                    Id = 0,
                },
                new Footnote(new Paragraph(new Run(new Text(text)))) { Id = noteId });
            part.Footnotes.Save();
        }
        return ms.ToArray();
    }
}
