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
