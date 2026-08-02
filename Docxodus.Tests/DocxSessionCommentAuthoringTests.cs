#nullable enable

using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Native Word comment <em>authoring and threads</em> on <see cref="DocxSession"/>
/// (issues #300 and #317):
/// <see cref="DocxSession.AddComment"/> creates the <c>WordprocessingCommentsPart</c> (plus the
/// <c>CommentText</c>/<c>CommentReference</c> styles) when absent, writes the
/// <c>w:commentRangeStart</c>/<c>w:commentRangeEnd</c> pair around a character span, appends the
/// run-level <c>w:commentReference</c>, and adds the <c>w:comment</c> definition. Editing a comment
/// body goes through <see cref="DocxSession.UpdateComment"/> (or <see cref="DocxSession.ReplaceText"/>
/// on a <c>p:cmt</c> paragraph); removal through <see cref="DocxSession.RemoveComment"/> /
/// <see cref="DocxSession.DeleteBlock"/>. Replies use Word's reference-only child shape plus
/// <c>commentsExtended.xml</c>; resolve/reopen state is carried by <c>w15:done</c>. Test IDs use
/// DS34x/DS35x/DS36x (DS346+) for base comments and DS400–DS404 for threading/state.
/// </summary>
public class DocxSessionCommentAuthoringTests
{
    private static readonly XNamespace W =
        "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string FirstBodyParagraph(DocxSession session) =>
        session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Scope == "body" && t.Anchor.Kind is "p" or "h").Anchor.Id;

    private static XElement PartXml(byte[] docxBytes, Func<MainDocumentPart, OpenXmlPart?> pick)
    {
        using var ms = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var part = pick(doc.MainDocumentPart!);
        Assert.NotNull(part);
        return part!.GetXDocument().Root!;
    }

    private static XElement BodyXml(byte[] docxBytes) => PartXml(docxBytes, m => m);

    private static string[] IgnorableTokens(XElement root, XNamespace mc) =>
        ((string?)root.Attribute(mc + "Ignorable") ?? string.Empty)
            .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

    private static bool HasPart(byte[] docxBytes, Func<MainDocumentPart, OpenXmlPart?> pick)
    {
        using var ms = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        return pick(doc.MainDocumentPart!) is not null;
    }

    private static string PartRelationshipId(
        byte[] docxBytes, Func<MainDocumentPart, OpenXmlPart?> pick)
    {
        using var ms = new MemoryStream(docxBytes);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        var part = pick(main);
        Assert.NotNull(part);
        return main.GetIdOfPart(part!);
    }

    /// <summary>One body paragraph with the given text in a single run.</summary>
    private static byte[] BuildSingleParagraphDoc(string text)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                new DocumentFormat.OpenXml.Wordprocessing.Body(
                    new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                        new DocumentFormat.OpenXml.Wordprocessing.Run(
                            new DocumentFormat.OpenXml.Wordprocessing.Text(text)))));
        }
        return ms.ToArray();
    }

    /// <summary>Text of the inline content between the rangeStart/rangeEnd pair for <paramref name="id"/>.</summary>
    private static string TextBetweenRangeMarkers(XElement para, string id)
    {
        bool inRange = false;
        var sb = new StringBuilder();
        foreach (var node in para.Descendants())
        {
            if (node.Name == W + "commentRangeStart" && (string?)node.Attribute(W + "id") == id) { inRange = true; continue; }
            if (node.Name == W + "commentRangeEnd" && (string?)node.Attribute(W + "id") == id) break;
            if (inRange && node.Name == W + "t") sb.Append((string)node);
        }
        return sb.ToString();
    }

    // ─── Creation: part, definition, body plumbing, styles ──────────────

    [Fact]
    public void DS346_AddComment_CreatesPartDefinitionAndBodyPlumbing()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var result = session.AddComment(anchor, null, "Alice", "Needs review.");
        Assert.True(result.Success, result.Error?.Message);

        var saved = session.Save();

        // Definition: one w:comment with author, an id, and NO date (deterministic default).
        var comments = PartXml(saved, m => m.WordprocessingCommentsPart);
        Assert.Equal(W + "comments", comments.Name);
        var comment = Assert.Single(comments.Elements(W + "comment"));
        Assert.Equal("Alice", (string?)comment.Attribute(W + "author"));
        var id = (string?)comment.Attribute(W + "id");
        Assert.False(string.IsNullOrEmpty(id));
        Assert.Null(comment.Attribute(W + "date"));
        Assert.Contains("Needs review.", comment.Descendants(W + "t").Select(t => (string)t));

        // Definition paragraph: CommentText style + leading annotationRef mark run.
        var firstPara = comment.Elements(W + "p").First();
        Assert.Equal("CommentText",
            (string?)firstPara.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val"));
        var markRun = firstPara.Elements(W + "r").First();
        Assert.Single(markRun.Elements(W + "annotationRef"));
        Assert.Equal("CommentReference",
            (string?)markRun.Element(W + "rPr")?.Element(W + "rStyle")?.Attribute(W + "val"));

        // Body plumbing: rangeStart + rangeEnd with the same id, reference run directly after rangeEnd.
        var body = BodyXml(saved);
        var para = body.Descendants(W + "p").First(p => p.Descendants(W + "commentRangeStart").Any());
        Assert.Single(para.Descendants(W + "commentRangeStart").Where(e => (string?)e.Attribute(W + "id") == id));
        var rangeEnd = Assert.Single(para.Descendants(W + "commentRangeEnd").Where(e => (string?)e.Attribute(W + "id") == id));
        var next = rangeEnd.ElementsAfterSelf().First();
        Assert.Equal(W + "r", next.Name);
        Assert.Single(next.Elements(W + "commentReference").Where(e => (string?)e.Attribute(W + "id") == id));
        Assert.Equal("CommentReference",
            (string?)next.Element(W + "rPr")?.Element(W + "rStyle")?.Attribute(W + "val"));

        // Both referenced styles are actually defined.
        var styles = PartXml(saved, m => m.StyleDefinitionsPart);
        var ids = styles.Elements(W + "style").Select(s => (string?)s.Attribute(W + "styleId")).ToList();
        Assert.Contains("CommentText", ids);
        Assert.Contains("CommentReference", ids);
    }

    [Fact]
    public void DS347_AddComment_SpanBracketsExactRange()
    {
        using var session = new DocxSession(BuildSingleParagraphDoc("Hello brave new world"));
        var anchor = FirstBodyParagraph(session);

        var result = session.AddComment(anchor, new CharSpan(6, 5), "Bob", "On one word.");
        Assert.True(result.Success, result.Error?.Message);

        var body = BodyXml(session.Save());
        var para = body.Descendants(W + "p").First(p => p.Descendants(W + "commentRangeStart").Any());
        var id = (string?)para.Descendants(W + "commentRangeStart").First().Attribute(W + "id");

        Assert.Equal("brave", TextBetweenRangeMarkers(para, id!));

        // Paragraph text is preserved verbatim across the mid-run splits.
        Assert.Equal("Hello brave new world",
            string.Concat(para.Descendants(W + "t").Select(t => (string)t)));
    }

    [Fact]
    public void DS348_AddComment_MetadataAttributes()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var date = new DateTime(2026, 8, 1, 12, 30, 0, DateTimeKind.Utc);
        var result = session.AddComment(anchor, null, "Alice", "Dated.", initials: "AL", date: date);
        Assert.True(result.Success, result.Error?.Message);

        var comments = PartXml(session.Save(), m => m.WordprocessingCommentsPart);
        var comment = Assert.Single(comments.Elements(W + "comment"));
        Assert.Equal("AL", (string?)comment.Attribute(W + "initials"));
        Assert.Equal("2026-08-01T12:30:00Z", (string?)comment.Attribute(W + "date"));
    }

    [Fact]
    public void DS349_AddComment_SecondReusesPartAndIncrementsId()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        Assert.True(session.AddComment(anchor, null, "Alice", "First comment.").Success);
        Assert.True(session.AddComment(anchor, null, "Bob", "Second comment.").Success);

        var comments = PartXml(session.Save(), m => m.WordprocessingCommentsPart);
        var ids = comments.Elements(W + "comment")
            .Select(c => int.Parse((string)c.Attribute(W + "id")!))
            .ToList();
        Assert.Equal(2, ids.Count);
        Assert.Equal(ids.Count, ids.Distinct().Count());

        var markdown = session.Project().Markdown;
        Assert.Contains("# Comments", markdown);
        Assert.Contains("First comment.", markdown);
        Assert.Contains("Second comment.", markdown);
    }

    [Fact]
    public void DS350_AddComment_EnvelopeAndProjection()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var result = session.AddComment(anchor, null, "Alice", "First para.\n\nSecond **para**.");
        Assert.True(result.Success, result.Error?.Message);

        // The definition anchor and its paragraph anchors come back Created, so a caller can
        // immediately address the comment for a follow-up edit.
        Assert.Contains(result.Created, a => a.Kind == "cmt" && a.Scope == "cmt");
        Assert.Equal(2, result.Created.Count(a => a.Kind == "p" && a.Scope == "cmt"));
        Assert.Contains(result.Modified, a => a.Id == anchor);
        Assert.NotNull(result.Patch);

        var markdown = session.Project().Markdown;
        Assert.Contains("# Comments", markdown);
        Assert.Contains("**Alice**", markdown);
        Assert.Contains("First para.", markdown);
    }

    [Fact]
    public void DS351_AddComment_Errors()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var missing = session.AddComment("p:body:ffffffff", null, "A", "x");
        Assert.False(missing.Success);
        Assert.Equal(EditErrorCode.AnchorNotFound, missing.Error!.Code);

        var seeded = session.AddComment(anchor, null, "A", "Comment for wrong-kind checks.");
        Assert.True(seeded.Success, seeded.Error?.Message);
        var cmtAnchor = seeded.Created.First(a => a.Kind == "cmt").Id;
        var cmtPara = seeded.Created.First(a => a.Kind == "p" && a.Scope == "cmt").Id;

        // The definition itself is not a legal host…
        var onDef = session.AddComment(cmtAnchor, null, "A", "Nested.");
        Assert.False(onDef.Success);
        Assert.Equal(EditErrorCode.AnchorWrongKind, onDef.Error!.Code);

        // …and neither is a comment-scope paragraph (Word forbids comments-on-comments).
        var onCmtPara = session.AddComment(cmtPara, null, "A", "Nested.");
        Assert.False(onCmtPara.Success);
        Assert.Equal(EditErrorCode.AnchorWrongKind, onCmtPara.Error!.Code);

        var empty = session.AddComment(anchor, new CharSpan(0, 0), "A", "x");
        Assert.False(empty.Success);
        Assert.Equal(EditErrorCode.EmptyCommentSpan, empty.Error!.Code);

        var outOfRange = session.AddComment(anchor, new CharSpan(0, 10_000), "A", "x");
        Assert.False(outOfRange.Success);
        Assert.Equal(EditErrorCode.OffsetOutOfRange, outOfRange.Error!.Code);
    }

    // ─── Undo / redo across the part create ─────────────────────────────

    [Fact]
    public void DS353_UndoFirstAddComment_RemovesThePart_RedoRestoresIt()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        Assert.True(session.AddComment(anchor, null, "Alice", "Undo me.").Success);

        Assert.True(session.Undo());
        using (var ms = new MemoryStream(session.Save()))
        using (var doc = WordprocessingDocument.Open(ms, false))
        {
            Assert.Null(doc.MainDocumentPart!.WordprocessingCommentsPart);
            var body = doc.MainDocumentPart.GetXDocument().Root!;
            Assert.Empty(body.Descendants(W + "commentReference"));
            Assert.Empty(body.Descendants(W + "commentRangeStart"));
            Assert.Empty(body.Descendants(W + "commentRangeEnd"));
        }

        Assert.True(session.Redo());
        using (var ms = new MemoryStream(session.Save()))
        using (var doc = WordprocessingDocument.Open(ms, false))
        {
            var part = doc.MainDocumentPart!.WordprocessingCommentsPart;
            Assert.NotNull(part);
            Assert.Single(part!.GetXDocument().Root!.Elements(W + "comment"));
            Assert.Single(doc.MainDocumentPart.GetXDocument().Root!.Descendants(W + "commentReference"));
        }
    }

    [Fact]
    public void DS354_UndoSecondAddComment_KeepsThePartAndRollsBackOnlyThatDefinition()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        Assert.True(session.AddComment(anchor, null, "Alice", "Keep me.").Success);
        Assert.True(session.AddComment(anchor, null, "Bob", "Roll me back.").Success);

        Assert.True(session.Undo());

        using var ms = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        var part = doc.MainDocumentPart!.WordprocessingCommentsPart;
        Assert.NotNull(part);
        var surviving = Assert.Single(part!.GetXDocument().Root!.Elements(W + "comment"));
        Assert.Contains("Keep me.", surviving.Descendants(W + "t").Select(t => (string)t));
        Assert.Single(doc.MainDocumentPart.GetXDocument().Root!.Descendants(W + "commentReference"));
    }

    // ─── UpdateComment / ListComments ───────────────────────────────────

    [Fact]
    public void DS355_UpdateComment_ReplacesBodyAndPreservesAttributes()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var added = session.AddComment(anchor, null, "Alice", "Original body.",
            initials: "AL", date: new DateTime(2026, 8, 1, 0, 0, 0, DateTimeKind.Utc));
        Assert.True(added.Success, added.Error?.Message);
        var cmtAnchor = added.Created.First(a => a.Kind == "cmt").Id;
        var oldParaUnids = added.Created.Where(a => a.Kind == "p" && a.Scope == "cmt")
            .Select(a => a.Unid).ToHashSet();

        var updated = session.UpdateComment(cmtAnchor, "Revised **text**.\n\nSecond para.");
        Assert.True(updated.Success, updated.Error?.Message);

        var comments = PartXml(session.Save(), m => m.WordprocessingCommentsPart);
        var comment = Assert.Single(comments.Elements(W + "comment"));
        Assert.Equal("Alice", (string?)comment.Attribute(W + "author"));
        Assert.Equal("AL", (string?)comment.Attribute(W + "initials"));
        Assert.Equal("2026-08-01T00:00:00Z", (string?)comment.Attribute(W + "date"));
        Assert.Equal(2, comment.Elements(W + "p").Count());
        Assert.Contains("Revised", comment.Descendants(W + "t").Select(t => (string)t).SelectMany(s => s.Split(' ')));
        Assert.DoesNotContain("Original body.", comment.Descendants(W + "t").Select(t => (string)t));

        // New body keeps the Word shape: CommentText style + leading annotationRef mark.
        var firstPara = comment.Elements(W + "p").First();
        Assert.Equal("CommentText",
            (string?)firstPara.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val"));
        Assert.Single(firstPara.Descendants(W + "annotationRef"));

        // Envelope: the definition is Modified; old paragraphs Removed, new ones Created.
        Assert.Contains(updated.Modified, a => a.Id == cmtAnchor);
        Assert.Equal(2, updated.Created.Count(a => a.Kind == "p" && a.Scope == "cmt"));
        Assert.Contains(updated.Removed, a => oldParaUnids.Contains(a.Unid));
    }

    [Fact]
    public void DS356_UpdateComment_PreservesLastParagraphParaIdForThreading()
    {
        using var session = new DocxSession(BuildDocWithThreadedComments());
        var cmtAnchor = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Kind == "cmt" && (t.TextPreview ?? "").Contains("Root")).Anchor.Id;

        var updated = session.UpdateComment(cmtAnchor, "Edited root body.");
        Assert.True(updated.Success, updated.Error?.Message);

        var comments = PartXml(session.Save(), m => m.WordprocessingCommentsPart);
        XNamespace w14 = "http://schemas.microsoft.com/office/word/2010/wordml";
        var root = comments.Elements(W + "comment")
            .First(c => c.Descendants(W + "t").Any(t => ((string)t).Contains("Edited root body.")));

        // commentsExtended entries key on the LAST paragraph's w14:paraId — a body edit must
        // not orphan the thread metadata.
        Assert.Equal("11111111", (string?)root.Elements(W + "p").Last().Attribute(w14 + "paraId"));
    }

    [Fact]
    public void DS357_UpdateComment_RequiresACommentAnchor()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var wrongKind = session.UpdateComment(anchor, "Nope.");
        Assert.False(wrongKind.Success);
        Assert.Equal(EditErrorCode.AnchorWrongKind, wrongKind.Error!.Code);

        var missing = session.UpdateComment("cmt:cmt:ffffffff", "Nope.");
        Assert.False(missing.Success);
        Assert.Equal(EditErrorCode.AnchorNotFound, missing.Error!.Code);
    }

    [Fact]
    public void DS358_ListComments_ReturnsEntriesInPartOrder()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var first = session.AddComment(anchor, null, "Alice", "First body.\n\nMore text.",
            initials: "AL", date: new DateTime(2026, 8, 1, 0, 0, 0, DateTimeKind.Utc));
        Assert.True(first.Success);
        var second = session.AddComment(anchor, null, "Bob", "Second body.");
        Assert.True(second.Success);

        var entries = session.ListComments();
        Assert.Equal(2, entries.Count);

        Assert.Equal(first.Created.First(a => a.Kind == "cmt").Id, entries[0].DefAnchorId);
        Assert.Equal("Alice", entries[0].Author);
        Assert.Equal("AL", entries[0].Initials);
        Assert.Equal("2026-08-01T00:00:00Z", entries[0].Date);
        Assert.Equal("First body. More text.", entries[0].Text);

        Assert.Equal("Bob", entries[1].Author);
        Assert.Null(entries[1].Initials);
        Assert.Null(entries[1].Date);
        Assert.Equal("Second body.", entries[1].Text);
    }

    [Fact]
    public void DS363_OpsWireRoundTrip()
    {
        var handle = Docxodus.Internal.DocxSessionOps.OpenSession(
            DocxSessionTests.BuildDS001_SimpleTwoParagraphs(), null);
        try
        {
            using var probe = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
            var anchor = FirstBodyParagraph(probe);

            var addJson = Docxodus.Internal.DocxSessionOps.AddComment(
                handle, anchor, null, "Alice", null, "2026-08-01T00:00:00Z", "Wire test.");
            Assert.Contains("\"success\":true", addJson);
            Assert.Contains("cmt:cmt:", addJson);

            var listJson = Docxodus.Internal.DocxSessionOps.ListComments(handle);
            Assert.Contains("\"author\":\"Alice\"", listJson);
            Assert.Contains("\"date\":\"2026-08-01T00:00:00Z\"", listJson);
            Assert.Contains("\"text\":\"Wire test.\"", listJson);
            Assert.DoesNotContain("\"resolved\"", listJson); // flat: additive fields omitted

            using var addDoc = System.Text.Json.JsonDocument.Parse(addJson);
            var parentAnchor = addDoc.RootElement.GetProperty("created").EnumerateArray()
                .First(a => a.GetProperty("kind").GetString() == "cmt")
                .GetProperty("id").GetString()!;
            var replyJson = Docxodus.Internal.DocxSessionOps.AddCommentReply(
                handle, parentAnchor, "Bob", "B", null, "Wire reply.");
            Assert.Contains("\"success\":true", replyJson);
            using var replyDoc = System.Text.Json.JsonDocument.Parse(replyJson);
            var replyAnchor = replyDoc.RootElement.GetProperty("created").EnumerateArray()
                .First(a => a.GetProperty("kind").GetString() == "cmt")
                .GetProperty("id").GetString()!;

            Assert.Contains("\"success\":true",
                Docxodus.Internal.DocxSessionOps.SetCommentResolved(handle, replyAnchor, true));
            var threadedListJson = Docxodus.Internal.DocxSessionOps.ListComments(handle);
            Assert.Contains($"\"parentAnchorId\":\"{parentAnchor}\"", threadedListJson);
            Assert.Contains("\"resolved\":true", threadedListJson);

            // Bad date string throws at the transport layer, never a silent drop.
            Assert.ThrowsAny<FormatException>(() =>
                Docxodus.Internal.DocxSessionOps.AddComment(
                    handle, anchor, null, "Alice", null, "not-a-date", "x"));
        }
        finally
        {
            Docxodus.Internal.DocxSessionOps.CloseSession(handle);
        }
    }

    [Fact]
    public void DS364_AuthoredComment_IsEditableThroughReplaceTextOnItsParagraph()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var added = session.AddComment(anchor, null, "Alice", "Original.");
        Assert.True(added.Success);
        var cmtPara = added.Created.First(a => a.Kind == "p" && a.Scope == "cmt").Id;

        var edited = session.ReplaceText(cmtPara, "Rewritten through ReplaceText.");
        Assert.True(edited.Success, edited.Error?.Message);
        Assert.Contains("Rewritten through ReplaceText.", session.Project().Markdown);
    }

    /// <summary>
    /// A Word-threaded fixture: two comments whose paragraphs carry <c>w14:paraId</c>, a
    /// <c>commentsExtended.xml</c> whose second entry replies to the first
    /// (<c>w15:paraIdParent</c>), and body markers for both. The shape Word 2013+ writes.
    /// </summary>
    private static byte[] BuildDocWithThreadedComments()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var s = main.GetStream(FileMode.Create))
            using (var w = new StreamWriter(s))
                w.Write("""
                    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                      <w:body>
                        <w:p>
                          <w:r><w:t xml:space="preserve">Alpha text with </w:t></w:r>
                          <w:commentRangeStart w:id="1"/>
                          <w:commentRangeStart w:id="2"/>
                          <w:r><w:t>target</w:t></w:r>
                          <w:commentRangeEnd w:id="1"/>
                          <w:commentRangeEnd w:id="2"/>
                          <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="1"/></w:r>
                          <w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="2"/></w:r>
                          <w:r><w:t xml:space="preserve"> and more.</w:t></w:r>
                        </w:p>
                      </w:body>
                    </w:document>
                    """);

            var commentsPart = main.AddNewPart<WordprocessingCommentsPart>();
            using (var s = commentsPart.GetStream(FileMode.Create))
            using (var w = new StreamWriter(s))
                w.Write("""
                    <w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                                xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
                                xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
                                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                                mc:Ignorable="w15">
                      <w:comment w:id="1" w:author="Alice" w:initials="A">
                        <w:p w14:paraId="11111111"><w:r><w:t>Root comment.</w:t></w:r></w:p>
                      </w:comment>
                      <w:comment w:id="2" w:author="Bob" w:initials="B">
                        <w:p w14:paraId="22222222"><w:r><w:t>Reply comment.</w:t></w:r></w:p>
                      </w:comment>
                    </w:comments>
                    """);

            var exPart = main.AddNewPart<WordprocessingCommentsExPart>();
            using (var s = exPart.GetStream(FileMode.Create))
            using (var w = new StreamWriter(s))
                w.Write("""
                    <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                      <w15:commentEx w15:paraId="11111111" w15:done="0"/>
                      <w15:commentEx w15:paraId="22222222" w15:paraIdParent="11111111" w15:done="0"/>
                    </w15:commentsEx>
                    """);

            var idsPart = main.AddNewPart<WordprocessingCommentsIdsPart>();
            using (var s = idsPart.GetStream(FileMode.Create))
            using (var w = new StreamWriter(s))
                w.Write("""
                    <w16cid:commentsIds xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid">
                      <w16cid:commentId w16cid:paraId="11111111" w16cid:durableId="1AAA1111"/>
                      <w16cid:commentId w16cid:paraId="22222222" w16cid:durableId="2BBB2222"/>
                    </w16cid:commentsIds>
                    """);
        }
        return ms.ToArray();
    }

    /// <summary>The canonical Word thread shape: only the root owns range markers; its reply
    /// contributes an adjacent reference and is linked by <c>w15:paraIdParent</c>.</summary>
    private static byte[] BuildDocWithReferenceOnlyReply()
    {
        using var ms = new MemoryStream();
        var source = BuildDocWithThreadedComments();
        ms.Write(source, 0, source.Length);
        ms.Position = 0;
        using (var doc = WordprocessingDocument.Open(ms, true))
        {
            var root = doc.MainDocumentPart!.GetXDocument().Root!;
            root.Descendants()
                .Where(e => e.Name is var name
                    && (name == W + "commentRangeStart" || name == W + "commentRangeEnd")
                    && (string?)e.Attribute(W + "id") == "2")
                .Remove();
            doc.MainDocumentPart.PutXDocument();
        }
        return ms.ToArray();
    }

    // ─── RemoveComment + threading pruning ──────────────────────────────

    [Fact]
    public void DS359_RemoveComment_StripsTripleAndDefinition_LeavesSiblingIntact()
    {
        using var session = new DocxSession(BuildDocWithThreadedComments());
        var cmtAnchor = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Kind == "cmt" && (t.TextPreview ?? "").Contains("Root")).Anchor.Id;

        var result = session.RemoveComment(cmtAnchor);
        Assert.True(result.Success, result.Error?.Message);
        Assert.Contains(result.Removed, a => a.Kind == "cmt");
        Assert.Contains(result.Removed, a => a.Kind == "p" && a.Scope == "cmt");

        var saved = session.Save();

        // Definition gone, sibling comment untouched.
        var comments = PartXml(saved, m => m.WordprocessingCommentsPart);
        var surviving = Assert.Single(comments.Elements(W + "comment"));
        Assert.Contains("Reply comment.", surviving.Descendants(W + "t").Select(t => (string)t));

        // The body triple for the removed comment is gone; the sibling's survives.
        var body = BodyXml(saved);
        var survivingId = (string?)surviving.Attribute(W + "id");
        Assert.Single(body.Descendants(W + "commentReference"));
        Assert.Equal(survivingId, (string?)body.Descendants(W + "commentReference").Single().Attribute(W + "id"));
        Assert.Single(body.Descendants(W + "commentRangeStart"));
        Assert.Single(body.Descendants(W + "commentRangeEnd"));

        // No empty wrapper run left behind (a w:r whose only child is w:rPr).
        Assert.DoesNotContain(body.Descendants(W + "r"),
            r => r.Elements().Any() && r.Elements().All(e => e.Name == W + "rPr"));

        // Body text is untouched.
        var para = body.Descendants(W + "p").First();
        Assert.Equal("Alpha text with target and more.",
            string.Concat(para.Descendants(W + "t").Select(t => (string)t)));
    }

    [Fact]
    public void DS360_RemoveComment_PrunesThreadingMetadata_AndUndoRestoresIt()
    {
        using var session = new DocxSession(BuildDocWithThreadedComments());
        XNamespace w15 = "http://schemas.microsoft.com/office/word/2012/wordml";
        XNamespace w16cid = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
        var cmtAnchor = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Kind == "cmt" && (t.TextPreview ?? "").Contains("Root")).Anchor.Id;

        Assert.True(session.RemoveComment(cmtAnchor).Success);

        var saved = session.Save();
        var ex = PartXml(saved, m => m.WordprocessingCommentsExPart);
        var entry = Assert.Single(ex.Elements(w15 + "commentEx"));
        Assert.Equal("22222222", (string?)entry.Attribute(w15 + "paraId"));
        // The reply no longer points at a removed parent — it became top-level, not dangling.
        Assert.Null(entry.Attribute(w15 + "paraIdParent"));

        var ids = PartXml(saved, m => m.WordprocessingCommentsIdsPart);
        var idEntry = Assert.Single(ids.Elements(w16cid + "commentId"));
        Assert.Equal("22222222", (string?)idEntry.Attribute(w16cid + "paraId"));

        // The pruning is undoable: the threading parts are snapshot-scoped.
        Assert.True(session.Undo());
        var exRestored = PartXml(session.Save(), m => m.WordprocessingCommentsExPart);
        Assert.Equal(2, exRestored.Elements(w15 + "commentEx").Count());
        Assert.Contains(exRestored.Elements(w15 + "commentEx"),
            e => (string?)e.Attribute(w15 + "paraIdParent") == "11111111");
    }

    [Fact]
    public void DS361_RemoveComment_RequiresACommentAnchor()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        var wrongKind = session.RemoveComment(anchor);
        Assert.False(wrongKind.Success);
        Assert.Equal(EditErrorCode.AnchorWrongKind, wrongKind.Error!.Code);
        Assert.Contains("RemoveComment", wrongKind.Error.Message);

        var missing = session.RemoveComment("cmt:cmt:ffffffff");
        Assert.False(missing.Success);
        Assert.Equal(EditErrorCode.AnchorNotFound, missing.Error!.Code);
    }

    [Fact]
    public void DS362_DeleteBlock_OnACommentAnchor_AlsoPrunesThreadingMetadata()
    {
        // Single-owner proof: the pruning lives in DeleteBlock's cmt teardown, so the generic
        // path gets it too — not just the typed RemoveComment wrapper.
        using var session = new DocxSession(BuildDocWithThreadedComments());
        XNamespace w15 = "http://schemas.microsoft.com/office/word/2012/wordml";
        var cmtAnchor = session.Project().AnchorIndex.Values
            .First(t => t.Anchor.Kind == "cmt" && (t.TextPreview ?? "").Contains("Root")).Anchor.Id;

        Assert.True(session.DeleteBlock(cmtAnchor).Success);

        var ex = PartXml(session.Save(), m => m.WordprocessingCommentsExPart);
        var entry = Assert.Single(ex.Elements(w15 + "commentEx"));
        Assert.Equal("22222222", (string?)entry.Attribute(w15 + "paraId"));
        Assert.Null(entry.Attribute(w15 + "paraIdParent"));
    }

    // ─── Reply threading + resolve/reopen (issue #317) ─────────────────

    [Fact]
    public void DS400_AddCommentReply_AuthorsNativeThread_AndSharesParentRange()
    {
        // The five-value constructor/deconstructor are an existing CLR contract. Thread fields
        // are init properties so extending the entry remains binary-compatible.
        var legacyEntry = new CommentListEntry("cmt:cmt:1", "A", null, null, "body");
        var (legacyAnchor, legacyAuthor, legacyInitials, legacyDate, legacyText) = legacyEntry;
        Assert.Equal("cmt:cmt:1", legacyAnchor);
        Assert.Equal("A", legacyAuthor);
        Assert.Null(legacyInitials);
        Assert.Null(legacyDate);
        Assert.Equal("body", legacyText);

        using var session = new DocxSession(BuildSingleParagraphDoc("Hello brave new world"));
        var host = FirstBodyParagraph(session);
        var parentResult = session.AddComment(host, new CharSpan(6, 5), "Alice", "Root comment.");
        Assert.True(parentResult.Success, parentResult.Error?.Message);
        var parentAnchor = parentResult.Created.First(a => a.Kind == "cmt").Id;

        // Flat comments remain distinguishable until a threading/resolve operation upgrades them.
        var flat = Assert.Single(session.ListComments());
        Assert.Null(flat.ParentAnchorId);
        Assert.Null(flat.Resolved);

        var replyResult = session.AddCommentReply(parentAnchor, "Bob", "Reply body.", initials: "B",
            date: new DateTime(2026, 8, 2, 12, 0, 0, DateTimeKind.Utc));
        Assert.True(replyResult.Success, replyResult.Error?.Message);
        Assert.Contains(replyResult.Modified, a => a.Id == parentAnchor);
        Assert.Contains(replyResult.Modified, a => a.Id == host);
        Assert.Equal(host, replyResult.Patch?.ScopeAnchorId);
        var replyAnchor = replyResult.Created.First(a => a.Kind == "cmt").Id;

        var entries = session.ListComments();
        Assert.Equal(2, entries.Count);
        Assert.Equal(parentAnchor, entries[0].DefAnchorId);
        Assert.Null(entries[0].ParentAnchorId);
        Assert.False(entries[0].Resolved);
        Assert.Equal(replyAnchor, entries[1].DefAnchorId);
        Assert.Equal(parentAnchor, entries[1].ParentAnchorId);
        Assert.False(entries[1].Resolved);
        Assert.Equal("Bob", entries[1].Author);
        Assert.Equal("B", entries[1].Initials);
        Assert.Equal("2026-08-02T12:00:00Z", entries[1].Date);

        var saved = session.Save();
        XNamespace w14 = "http://schemas.microsoft.com/office/word/2010/wordml";
        XNamespace w15 = "http://schemas.microsoft.com/office/word/2012/wordml";
        XNamespace w16cid = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
        XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        var comments = PartXml(saved, m => m.WordprocessingCommentsPart);
        Assert.Equal(mc, comments.GetNamespaceOfPrefix("mc"));
        Assert.Contains("w14", IgnorableTokens(comments, mc));
        var defs = comments.Elements(W + "comment").ToList();
        Assert.Equal(2, defs.Count);
        var parentId = (string)defs[0].Attribute(W + "id")!;
        var replyId = (string)defs[1].Attribute(W + "id")!;
        var parentParaId = (string)defs[0].Elements(W + "p").Last().Attribute(w14 + "paraId")!;
        var replyParaId = (string)defs[1].Elements(W + "p").Last().Attribute(w14 + "paraId")!;
        Assert.Equal("00000001", parentParaId);
        Assert.Equal("00000002", replyParaId);

        var ex = PartXml(saved, m => m.WordprocessingCommentsExPart);
        var exEntries = ex.Elements(w15 + "commentEx").ToList();
        Assert.Equal(2, exEntries.Count);
        Assert.Equal("0", (string?)exEntries[0].Attribute(w15 + "done"));
        Assert.Equal(parentParaId, (string?)exEntries[1].Attribute(w15 + "paraIdParent"));

        var ids = PartXml(saved, m => m.WordprocessingCommentsIdsPart);
        Assert.Equal(new[] { "00000001", "00000002" }, ids.Elements(w16cid + "commentId")
            .Select(e => (string)e.Attribute(w16cid + "durableId")!).ToArray());

        // Word's native thread shape keeps the range on the root and gives the reply only an
        // adjacent reference; commentsExtended parentage makes it share the root's exact range.
        var body = BodyXml(saved);
        var para = body.Descendants(W + "p").Single(p => p.Descendants(W + "commentReference").Any());
        Assert.Equal("brave", TextBetweenRangeMarkers(para, parentId));
        Assert.Equal(new[] { parentId }, para.Descendants(W + "commentRangeStart")
            .Select(e => (string)e.Attribute(W + "id")!).ToArray());
        Assert.Equal(new[] { parentId }, para.Descendants(W + "commentRangeEnd")
            .Select(e => (string)e.Attribute(W + "id")!).ToArray());
        Assert.Equal(new[] { parentId, replyId }, para.Descendants(W + "commentReference")
            .Select(e => (string)e.Attribute(W + "id")!).ToArray());

        using var ms = new MemoryStream(saved);
        using var doc = WordprocessingDocument.Open(ms, false);
        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator(
                DocumentFormat.OpenXml.FileFormatVersions.Office2019)
            .Validate(doc).Select(e => $"{e.Part?.Uri}: {e.Description}").ToList();
        Assert.Empty(errors);
    }

    [Fact]
    public void DS401_SetCommentResolved_FlatCommentCreatesParts_UndoRedoReconcilesThem()
    {
        using var session = new DocxSession(BuildSingleParagraphDoc("Resolve this"));
        var host = FirstBodyParagraph(session);
        var made = session.AddComment(host, null, "Alice", "Please resolve.");
        Assert.True(made.Success, made.Error?.Message);
        var anchor = made.Created.First(a => a.Kind == "cmt").Id;

        Assert.Null(Assert.Single(session.ListComments()).Resolved);
        Assert.False(HasPart(session.Save(), m => m.WordprocessingCommentsExPart));
        Assert.False(HasPart(session.Save(), m => m.WordprocessingCommentsIdsPart));

        var resolved = session.SetCommentResolved(anchor, true);
        Assert.True(resolved.Success, resolved.Error?.Message);
        Assert.True(Assert.Single(session.ListComments()).Resolved);
        var resolvedBytes = session.Save();
        Assert.True(HasPart(resolvedBytes, m => m.WordprocessingCommentsExPart));
        Assert.True(HasPart(resolvedBytes, m => m.WordprocessingCommentsIdsPart));
        var exRelationshipId = PartRelationshipId(resolvedBytes, m => m.WordprocessingCommentsExPart);
        var idsRelationshipId = PartRelationshipId(resolvedBytes, m => m.WordprocessingCommentsIdsPart);

        // Undo restores both content AND package topology: no empty/orphan extension parts remain.
        Assert.True(session.Undo());
        Assert.Null(Assert.Single(session.ListComments()).Resolved);
        Assert.False(HasPart(session.Save(), m => m.WordprocessingCommentsExPart));
        Assert.False(HasPart(session.Save(), m => m.WordprocessingCommentsIdsPart));

        Assert.True(session.Redo());
        Assert.True(Assert.Single(session.ListComments()).Resolved);
        var redoneBytes = session.Save();
        Assert.True(HasPart(redoneBytes, m => m.WordprocessingCommentsExPart));
        Assert.True(HasPart(redoneBytes, m => m.WordprocessingCommentsIdsPart));
        Assert.Equal(exRelationshipId,
            PartRelationshipId(redoneBytes, m => m.WordprocessingCommentsExPart));
        Assert.Equal(idsRelationshipId,
            PartRelationshipId(redoneBytes, m => m.WordprocessingCommentsIdsPart));

        var reopened = session.SetCommentResolved(anchor, false);
        Assert.True(reopened.Success, reopened.Error?.Message);
        Assert.False(Assert.Single(session.ListComments()).Resolved);
    }

    [Fact]
    public void DS401b_SetCommentResolved_PropagatesThroughReplySubtree()
    {
        using var session = new DocxSession(BuildSingleParagraphDoc("Resolve this thread"));
        var host = FirstBodyParagraph(session);
        var made = session.AddComment(host, null, "Alice", "Root comment.");
        Assert.True(made.Success, made.Error?.Message);
        var rootAnchor = made.Created.First(a => a.Kind == "cmt").Id;
        var reply = session.AddCommentReply(rootAnchor, "Bob", "Reply comment.");
        Assert.True(reply.Success, reply.Error?.Message);

        var resolved = session.SetCommentResolved(rootAnchor, true);

        Assert.True(resolved.Success, resolved.Error?.Message);
        Assert.Equal(2, resolved.Modified.Count);
        Assert.All(session.ListComments(), comment => Assert.True(comment.Resolved));

        var reopened = session.SetCommentResolved(rootAnchor, false);
        Assert.True(reopened.Success, reopened.Error?.Message);
        Assert.All(session.ListComments(), comment => Assert.False(comment.Resolved));
    }

    [Fact]
    public void DS402_ExistingThread_ListsParentAndResolveState_WithoutLosingParentage()
    {
        using var session = new DocxSession(BuildDocWithReferenceOnlyReply());
        var host = FirstBodyParagraph(session);
        var entries = session.ListComments();
        Assert.Equal(2, entries.Count);
        Assert.Null(entries[0].ParentAnchorId);
        Assert.False(entries[0].Resolved);
        Assert.Equal(entries[0].DefAnchorId, entries[1].ParentAnchorId);
        Assert.False(entries[1].Resolved);

        Assert.True(session.SetCommentResolved(entries[1].DefAnchorId, true).Success);
        var resolvedReply = session.ListComments()[1];
        Assert.True(resolvedReply.Resolved);
        Assert.Equal(entries[0].DefAnchorId, resolvedReply.ParentAnchorId);

        Assert.True(session.SetCommentResolved(entries[1].DefAnchorId, false).Success);
        var reopenedReply = session.ListComments()[1];
        Assert.False(reopenedReply.Resolved);
        Assert.Equal(entries[0].DefAnchorId, reopenedReply.ParentAnchorId);

        // Regress the canonical nested case: the immediate parent has only a reference, while
        // the thread root owns the range. A child-of-reply stays on that range and links to the
        // immediate parent rather than being downgraded to an unrelated point comment.
        var nested = session.AddCommentReply(entries[1].DefAnchorId, "Carol", "Nested reply.");
        Assert.True(nested.Success, nested.Error?.Message);
        Assert.Contains(nested.Modified, a => a.Id == entries[1].DefAnchorId);
        Assert.Contains(nested.Modified, a => a.Id == host);
        Assert.Equal(host, nested.Patch?.ScopeAnchorId);
        var nestedAnchor = nested.Created.First(a => a.Kind == "cmt").Id;
        var nestedEntry = session.ListComments().Single(e => e.DefAnchorId == nestedAnchor);
        Assert.Equal(entries[1].DefAnchorId, nestedEntry.ParentAnchorId);

        var saved = session.Save();
        XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        var commentsRoot = PartXml(saved, m => m.WordprocessingCommentsPart);
        Assert.Equal(new[] { "w15", "w14" }, IgnorableTokens(commentsRoot, mc));
        var comments = commentsRoot.Elements(W + "comment").ToList();
        var ids = comments.Select(c => (string)c.Attribute(W + "id")!).ToArray();
        var body = BodyXml(saved);
        Assert.Equal(new[] { ids[0] }, body.Descendants(W + "commentRangeStart")
            .Select(e => (string)e.Attribute(W + "id")!).ToArray());
        Assert.Equal(new[] { ids[0] }, body.Descendants(W + "commentRangeEnd")
            .Select(e => (string)e.Attribute(W + "id")!).ToArray());
        Assert.Equal(ids, body.Descendants(W + "commentReference")
            .Select(e => (string)e.Attribute(W + "id")!).ToArray());
    }

    [Fact]
    public void DS403_ReplyMetadataIds_AreDeterministicMaxPlusOne()
    {
        using var session = new DocxSession(BuildDocWithThreadedComments());
        var parent = session.ListComments()[0].DefAnchorId;
        Assert.True(session.AddCommentReply(parent, "Carol", "Another reply.").Success);

        XNamespace w14 = "http://schemas.microsoft.com/office/word/2010/wordml";
        XNamespace w16cid = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
        var saved = session.Save();
        var comments = PartXml(saved, m => m.WordprocessingCommentsPart);
        Assert.Equal("22222223", (string?)comments.Elements(W + "comment").Last()
            .Elements(W + "p").Last().Attribute(w14 + "paraId"));
        var ids = PartXml(saved, m => m.WordprocessingCommentsIdsPart);
        Assert.Equal("2BBB2223", (string?)ids.Elements(w16cid + "commentId").Last()
            .Attribute(w16cid + "durableId"));
    }

    [Fact]
    public void DS404_ReplyAndResolve_RequireCommentAnchors()
    {
        using var session = new DocxSession(BuildSingleParagraphDoc("Wrong kind"));
        var host = FirstBodyParagraph(session);

        var reply = session.AddCommentReply(host, "A", "Nope.");
        Assert.False(reply.Success);
        Assert.Equal(EditErrorCode.AnchorWrongKind, reply.Error!.Code);

        var resolve = session.SetCommentResolved(host, true);
        Assert.False(resolve.Success);
        Assert.Equal(EditErrorCode.AnchorWrongKind, resolve.Error!.Code);
    }

    [Fact]
    public void DS365_AuthoredComment_RendersThroughTheHtmlConverter()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);
        Assert.True(session.AddComment(anchor, null, "Alice", "Rendered in HTML.").Success);

        var wml = new WmlDocument("commented.docx", session.Save());
        var settings = new WmlToHtmlConverterSettings { RenderComments = true };
        var html = WmlToHtmlConverter.ConvertToHtml(wml, settings).ToString(SaveOptions.DisableFormatting);

        Assert.Contains("Rendered in HTML.", html);
        Assert.Contains("comments-section", html);
    }

    [Fact]
    public void DS352_AddComment_ProducesASchemaValidDocument()
    {
        using var session = new DocxSession(DocxSessionTests.BuildDS001_SimpleTwoParagraphs());
        var anchor = FirstBodyParagraph(session);

        Assert.True(session.AddComment(anchor, new CharSpan(0, 3), "Alice",
            "A **bold** claim.\n\nSecond paragraph.", initials: "AL",
            date: new DateTime(2026, 8, 1, 0, 0, 0, DateTimeKind.Utc)).Success);

        using var ms = new MemoryStream(session.Save());
        using var doc = WordprocessingDocument.Open(ms, false);
        var errors = new DocumentFormat.OpenXml.Validation.OpenXmlValidator()
            .Validate(doc)
            .Select(e => $"{e.Part?.Uri}: {e.Description}")
            .ToList();
        Assert.Empty(errors);
    }
}
