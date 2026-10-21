using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

/// <summary>
/// In <see cref="CommentRenderMode.Margin"/> a reply's note nests inside its parent's note and
/// also wears the base note class. The note's box (fill, border, shadow, padding) therefore has to
/// be undone for replies, or every nesting level draws another box and narrows the reply text.
/// The browser geometry is measured by <c>npm/tests/margin-reply-boxes.spec.ts</c>.
/// </summary>
public class MarginCommentReplyBoxTests
{
    [Theory]
    [InlineData("comment-")]
    [InlineData("review-")]
    public void Reply_note_draws_no_box_of_its_own(string prefix)
    {
        var css = MarginCss(prefix);
        var reply = Rule(css, $"div.{prefix}margin-reply");

        Assert.Equal("none", reply.GetValueOrDefault("background"));
        Assert.Equal("none", reply.GetValueOrDefault("box-shadow"));
        Assert.Equal("0", reply.GetValueOrDefault("border-radius"));
        Assert.Equal("0", reply.GetValueOrDefault("border"));
        // Vertical space separates the replies; no horizontal inset narrows the text.
        Assert.Equal("0.5em 0 0", reply.GetValueOrDefault("padding"));
        Assert.Equal("0", reply.GetValueOrDefault("margin"));
    }

    [Theory]
    [InlineData("comment-")]
    [InlineData("review-")]
    public void Replies_container_does_not_indent(string prefix)
    {
        var replies = Rule(MarginCss(prefix), $"div.{prefix}margin-replies");

        Assert.DoesNotContain("padding-left", replies.Keys);
        Assert.DoesNotContain("border-left", replies.Keys);
        Assert.DoesNotContain("margin-left", replies.Keys);
    }

    [Fact]
    public void Top_level_note_keeps_its_box()
    {
        // Control: only replies lose the box; the thread's root note still draws it.
        var note = Rule(MarginCss("comment-"), "div.comment-margin-note");

        Assert.Equal("#fff9c4", note.GetValueOrDefault("background-color"));
        Assert.Equal("3px solid #fbc02d", note.GetValueOrDefault("border-left"));
        Assert.Equal("0 1px 3px rgba(0,0,0,0.1)", note.GetValueOrDefault("box-shadow"));
        Assert.Equal("0.5em", note.GetValueOrDefault("padding"));
    }

    [Fact]
    public void Reply_rule_follows_the_note_rule_it_overrides()
    {
        // Both selectors have the same specificity, so the later rule wins the cascade.
        var css = MarginCss("comment-");
        Assert.True(css.IndexOf("div.comment-margin-reply {") > css.IndexOf("div.comment-margin-note {"));
    }

    [Fact]
    public void Three_level_thread_nests_every_reply_as_a_note()
    {
        // The markup the CSS relies on: each reply is a note with the reply class, inside its
        // parent's replies container.
        var html = Convert(BuildThreadDoc(), "comment-");
        var root = NoteWithText(html, "Root note.");
        var reply = NoteWithText(html, "First reply.");
        var nested = NoteWithText(html, "Nested reply.");

        Assert.Equal("comment-margin-note", (string?)root.Attribute("class"));
        Assert.Equal("comment-margin-note comment-margin-reply", (string?)reply.Attribute("class"));
        Assert.Equal("comment-margin-note comment-margin-reply", (string?)nested.Attribute("class"));
        Assert.Same(root, ParentNote(reply));
        Assert.Same(reply, ParentNote(nested));
    }

    private static XElement ParentNote(XElement note) =>
        note.Ancestors().First(e => ((string?)e.Attribute("class") ?? "").Contains("margin-note"));

    private static XElement NoteWithText(XElement html, string text) =>
        html.Descendants().Single(e => e.Name.LocalName == "p" && e.Value == text)
            .Ancestors().First(e => e.Attribute("data-comment-id") != null);

    private static string MarginCss(string prefix) =>
        Convert(BuildThreadDoc(), prefix).Descendants().First(e => e.Name.LocalName == "style").Value;

    private static Dictionary<string, string> Rule(string css, string selector)
    {
        var match = Regex.Match(css, Regex.Escape(selector) + @"\s*\{([^}]*)\}");
        Assert.True(match.Success, $"no rule for {selector}");
        return match.Groups[1].Value
            .Split(';')
            .Select(d => d.Split(':', 2))
            .Where(p => p.Length == 2)
            .ToDictionary(p => p[0].Trim(), p => p[1].Trim());
    }

    private static XElement Convert(byte[] docx, string prefix)
    {
        using var ms = new MemoryStream();
        ms.Write(docx, 0, docx.Length);
        using var doc = WordprocessingDocument.Open(ms, true);
        return WmlToHtmlConverter.ConvertToHtml(doc, new WmlToHtmlConverterSettings
        {
            RenderComments = true,
            CommentRenderMode = CommentRenderMode.Margin,
            CommentCssClassPrefix = prefix,
        });
    }

    /// <summary>A root comment, a reply to it, and a reply to that reply.</summary>
    private static byte[] BuildThreadDoc()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(new Wp.Body(new Wp.Paragraph(new Wp.Run(new Wp.Text("Review this clause.")))));
        }

        using var session = new DocxSession(ms.ToArray());
        var paragraph = session.Project().AnchorIndex.Values.First(t => t.Anchor.Kind == "p").Anchor.Id;
        var root = session.AddComment(paragraph, new CharSpan(0, 6), "Ann", "Root note.");
        var reply = session.AddCommentReply(root.Created.First(a => a.Kind == "cmt").Id, "Bo", "First reply.");
        var nested = session.AddCommentReply(reply.Created.First(a => a.Kind == "cmt").Id, "Cy", "Nested reply.");
        Assert.True(nested.Success, nested.Error?.Message);
        return session.Save();
    }
}
