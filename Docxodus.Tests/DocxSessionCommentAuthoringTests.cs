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
/// Native Word comment <em>authoring</em> on <see cref="DocxSession"/> (issue #300):
/// <see cref="DocxSession.AddComment"/> creates the <c>WordprocessingCommentsPart</c> (plus the
/// <c>CommentText</c>/<c>CommentReference</c> styles) when absent, writes the
/// <c>w:commentRangeStart</c>/<c>w:commentRangeEnd</c> pair around a character span, appends the
/// run-level <c>w:commentReference</c>, and adds the <c>w:comment</c> definition. Editing a comment
/// body goes through <see cref="DocxSession.UpdateComment"/> (or <see cref="DocxSession.ReplaceText"/>
/// on a <c>p:cmt</c> paragraph); removal through <see cref="DocxSession.RemoveComment"/> /
/// <see cref="DocxSession.DeleteBlock"/>. Test IDs use the DS34x/DS35x/DS36x range (DS346+).
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
                                xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
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
        }
        return ms.ToArray();
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
