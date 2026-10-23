using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

/// <summary>
/// A comment anchored in a header, footer, footnote or endnote renders like a body comment: its
/// range is highlighted, it has a reference marker, and it has an entry (issue #1310). Before the
/// fix the converter read those comments but emitted nothing for them, because the comment state
/// was attached to the body's part only and every other story's markup found none.
/// </summary>
public class StoryCommentRenderingTests
{
    private static readonly XNamespace H = "http://www.w3.org/1999/xhtml";

    /// <summary>Comment id, the commented text in its story, and the comment's own text.</summary>
    public static readonly (int Id, string Text, string Comment)[] Stories =
    {
        (0, "BODYTEXT", "BODYCOMMENT"),
        (1, "HEADERTEXT", "HEADERCOMMENT"),
        (2, "FOOTERTEXT", "FOOTERCOMMENT"),
        (3, "NOTETEXT", "NOTECOMMENT"),
        (4, "ENDTEXT", "ENDCOMMENT"),
    };

    public static IEnumerable<object[]> Modes() =>
        from mode in new[] { CommentRenderMode.EndnoteStyle, CommentRenderMode.Inline, CommentRenderMode.Margin }
        from paginated in new[] { false, true }
        select new object[] { mode, paginated };

    [Theory]
    [MemberData(nameof(Modes))]
    public void Comment_in_every_story_gets_a_highlighted_range_and_a_marker(CommentRenderMode mode, bool paginated)
    {
        var html = Convert(StoryCommentsDocx(), mode, paginated);

        foreach (var (id, text, _) in Stories)
        {
            Assert.Contains(Highlights(html, id), span => span.Value.Contains(text));
            Assert.Single(Markers(html, id));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Endnote_style_lists_every_story_comment_once_body_first(bool paginated)
    {
        var html = Convert(StoryCommentsDocx(), CommentRenderMode.EndnoteStyle, paginated);
        var items = html.Descendants(H + "li").Where(li => ((string?)li.Attribute("id"))?.StartsWith("comment-") == true)
            .ToList();

        // Word's story order: the main document, then headers and footers, footnotes, endnotes.
        Assert.Equal(new[] { "comment-0", "comment-1", "comment-2", "comment-3", "comment-4" },
            items.Select(li => (string?)li.Attribute("id")));
        foreach (var (id, _, comment) in Stories)
            Assert.Contains(comment, items.Single(li => (string?)li.Attribute("id") == $"comment-{id}").Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Inline_mode_carries_each_story_comment_on_its_highlight(bool paginated)
    {
        var html = Convert(StoryCommentsDocx(), CommentRenderMode.Inline, paginated);

        foreach (var (id, _, comment) in Stories)
            Assert.Contains(Highlights(html, id), span => (string?)span.Attribute("data-comment") == comment);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Margin_mode_has_a_note_for_every_story_comment(bool paginated)
    {
        var html = Convert(StoryCommentsDocx(), CommentRenderMode.Margin, paginated);
        var notes = html.Descendants(H + "div")
            .Where(div => ((string?)div.Attribute("class"))?.Split(' ').Contains("comment-margin-note") == true)
            .ToList();

        Assert.Equal(Stories.Select(s => s.Id.ToString()), notes.Select(n => (string?)n.Attribute("data-comment-id")));
        foreach (var (id, _, comment) in Stories)
            Assert.Contains(comment, notes.Single(n => (string?)n.Attribute("data-comment-id") == id.ToString()).Value);
    }

    [Theory]
    [MemberData(nameof(Modes))]
    public void Body_only_comments_render_as_before(CommentRenderMode mode, bool paginated)
    {
        // Control: a document whose only comment is in the body, with the same running stories
        // and notes, renders that one comment and nothing else comment-shaped.
        var html = Convert(StoryCommentsDocx(bodyOnly: true), mode, paginated);

        Assert.Single(Markers(html, 0));
        Assert.Contains(Highlights(html, 0), span => span.Value.Contains("BODYTEXT"));
        Assert.All(html.Descendants().Where(e => e.Attribute("data-comment-id") != null),
            e => Assert.Equal("0", (string?)e.Attribute("data-comment-id")));
        Assert.Contains("FOOTERTEXT", html.Value);
        Assert.Contains("NOTETEXT", html.Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Range_left_open_in_the_body_does_not_highlight_another_story(bool paginated)
    {
        // Each story keeps its own open ranges: a body range with no end must not run on into
        // the header (converted after the body when paginated) or the notes and footer
        // (converted after it in a continuous render).
        var html = Convert(StoryCommentsDocx(unclosedBody: true), CommentRenderMode.EndnoteStyle, paginated);

        foreach (var (id, text, _) in Stories.Skip(1))
        {
            var span = Assert.Single(html.Descendants(H + "span"), s =>
                ((string?)s.Attribute("class"))?.Split(' ').Contains("comment-highlight") == true
                && s.Value.Contains(text));
            Assert.Equal(id.ToString(), (string?)span.Attribute("data-comment-id"));
        }
    }

    private static IEnumerable<XElement> Highlights(XElement html, int id) =>
        html.Descendants(H + "span").Where(span =>
            ((string?)span.Attribute("class"))?.Split(' ').Contains("comment-highlight") == true
            && ((string?)span.Attribute("data-comment-id"))?.Split(',').Contains(id.ToString()) == true);

    private static IEnumerable<XElement> Markers(XElement html, int id) =>
        html.Descendants(H + "a").Where(a => (string?)a.Attribute("id") == $"comment-ref-{id}");

    private static XElement Convert(byte[] docx, CommentRenderMode mode, bool paginated) =>
        WmlToHtmlConverter.ConvertToHtml(new WmlDocument("story-comments.docx", docx), new WmlToHtmlConverterSettings
        {
            RenderComments = true,
            CommentRenderMode = mode,
            RenderFootnotesAndEndnotes = true,
            RenderHeadersAndFooters = true,
            RenderPagination = paginated ? PaginationMode.Paginated : PaginationMode.None,
        });

    /// <summary>
    /// The issue's package: a body paragraph citing a footnote and an endnote, a default header
    /// and footer, and a comment wrapping the text of each of those five stories. With
    /// <paramref name="bodyOnly"/> only the body comment remains (the control); with
    /// <paramref name="unclosedBody"/> the body comment's range has no end.
    /// </summary>
    internal static byte[] StoryCommentsDocx(bool bodyOnly = false, bool unclosedBody = false)
    {
        bool commented(int id) => !bodyOnly || id == 0;
        OpenXmlElement[] Commented(int id, string text) => commented(id)
            ? new OpenXmlElement?[]
            {
                new Wp.CommentRangeStart { Id = id.ToString() },
                new Wp.Run(new Wp.Text(text)),
                unclosedBody && id == 0 ? null : new Wp.CommentRangeEnd { Id = id.ToString() },
                new Wp.Run(new Wp.CommentReference { Id = id.ToString() }),
            }.OfType<OpenXmlElement>().ToArray()
            : new OpenXmlElement[] { new Wp.Run(new Wp.Text(text)) };

        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Wp.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Wp.Settings();

            var header = main.AddNewPart<HeaderPart>();
            header.Header = new Wp.Header(new Wp.Paragraph(Commented(1, "HEADERTEXT")));
            var footer = main.AddNewPart<FooterPart>();
            footer.Footer = new Wp.Footer(new Wp.Paragraph(Commented(2, "FOOTERTEXT")));

            main.AddNewPart<FootnotesPart>().Footnotes = new Wp.Footnotes(
                new Wp.Footnote(new Wp.Paragraph(new Wp.Run(new Wp.SeparatorMark())))
                    { Type = Wp.FootnoteEndnoteValues.Separator, Id = 0 },
                new Wp.Footnote(new Wp.Paragraph(new Wp.Run(new Wp.ContinuationSeparatorMark())))
                    { Type = Wp.FootnoteEndnoteValues.ContinuationSeparator, Id = 1 },
                new Wp.Footnote(new Wp.Paragraph(
                    new OpenXmlElement[] { new Wp.Run(new Wp.FootnoteReferenceMark()) }
                        .Concat(Commented(3, "NOTETEXT")))) { Id = 2 });
            main.AddNewPart<EndnotesPart>().Endnotes = new Wp.Endnotes(
                new Wp.Endnote(new Wp.Paragraph(new Wp.Run(new Wp.SeparatorMark())))
                    { Type = Wp.FootnoteEndnoteValues.Separator, Id = 0 },
                new Wp.Endnote(new Wp.Paragraph(new Wp.Run(new Wp.ContinuationSeparatorMark())))
                    { Type = Wp.FootnoteEndnoteValues.ContinuationSeparator, Id = 1 },
                new Wp.Endnote(new Wp.Paragraph(
                    new OpenXmlElement[] { new Wp.Run(new Wp.EndnoteReferenceMark()) }
                        .Concat(Commented(4, "ENDTEXT")))) { Id = 2 });

            main.AddNewPart<WordprocessingCommentsPart>().Comments = new Wp.Comments(
                Stories.Where(s => commented(s.Id)).Select(s => new Wp.Comment(
                    new Wp.Paragraph(new Wp.Run(new Wp.Text(s.Comment))))
                    { Id = s.Id.ToString(), Author = "Reviewer", Initials = "R" }));

            main.Document = new Wp.Document(new Wp.Body(
                new Wp.Paragraph(Commented(0, "BODYTEXT")
                    .Append(new Wp.Run(new Wp.FootnoteReference { Id = 2 }))
                    .Append(new Wp.Run(new Wp.EndnoteReference { Id = 2 }))),
                new Wp.SectionProperties(
                    new Wp.HeaderReference { Type = Wp.HeaderFooterValues.Default, Id = main.GetIdOfPart(header) },
                    new Wp.FooterReference { Type = Wp.HeaderFooterValues.Default, Id = main.GetIdOfPart(footer) },
                    new Wp.PageSize { Width = 12240u, Height = 15840u },
                    new Wp.PageMargin
                    {
                        Top = 1440, Bottom = 1440, Left = 1440, Right = 1440, Header = 720, Footer = 720, Gutter = 0,
                    })));
        }
        return stream.ToArray();
    }
}
