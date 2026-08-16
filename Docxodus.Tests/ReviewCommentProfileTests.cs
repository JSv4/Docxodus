// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#nullable enable
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;

namespace Docxodus.Tests;

public class ReviewCommentProfileTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    [Fact]
    public void Endnote_profile_renders_cross_story_references_and_ordered_thread_metadata()
    {
        var html = Render(CommentRenderMode.EndnoteStyle, renderComments: true);

        for (var id = 1; id <= 6; id++)
            Assert.Single(html.Descendants(), e => (string?)e.Attribute("id") == $"comment-ref-{id}");

        var section = Assert.Single(html.Descendants(), e => HasClass(e, "comments-section"));
        var list = Assert.Single(section.Elements(), e => e.Name.LocalName == "ol");
        var topLevel = list.Elements().Where(e => e.Name.LocalName == "li").ToList();
        Assert.DoesNotContain(topLevel, item => (string?)item.Attribute("id") == "comment-2");

        var root = ById(section, "comment-1");
        var reply = ById(root, "comment-2");
        Assert.Equal("1", (string?)root.Attribute("data-comment-node-id"));
        Assert.Equal("2", (string?)reply.Attribute("data-comment-node-id"));
        Assert.Contains("Root review body.", root.Value);
        Assert.Contains("Reply review body.", reply.Value);
        Assert.Equal("resolved", (string?)root.Attribute("data-comment-status"));
        Assert.Equal("open", (string?)reply.Attribute("data-comment-status"));
        Assert.Equal("1", (string?)reply.Attribute("data-comment-parent-id"));
        Assert.Contains("Alice", root.Value);
        Assert.Contains("2024-01-02T03:04:05Z", root.Value);
        Assert.Contains("Resolved", root.Value);
        Assert.Contains("Reply", reply.Value);
        Assert.Contains("Open", reply.Value);
        Assert.Equal("#comment-ref-1", (string?)Assert.Single(reply.Descendants(),
            e => HasClass(e, "comment-backref")).Attribute("href"));

        var unknown = ById(section, "comment-3");
        Assert.Equal("unknown", (string?)unknown.Attribute("data-comment-status"));
        Assert.Contains("Status unknown", unknown.Value);
        Assert.Contains("Header review body.", unknown.Value);
        Assert.Contains("Footer review body.", section.Value);
        Assert.Contains("Footnote review body.", section.Value);
        Assert.Contains("Endnote review body.", section.Value);
    }

    [Fact]
    public void Inline_profile_emits_each_visible_thread_once_with_body_author_date_and_status()
    {
        var html = Render(CommentRenderMode.Inline, renderComments: true);

        Assert.DoesNotContain(html.Descendants(), e => HasClass(e, "comments-section"));
        var threads = html.Descendants().Where(e => HasClass(e, "comment-inline-thread")).ToList();
        Assert.Equal(5, threads.Count); // root+reply is one thread; four other roots are independent.

        var rootThread = Assert.Single(threads, e =>
            e.DescendantsAndSelf().Any(n => (string?)n.Attribute("id") == "comment-1"));
        Assert.Single(rootThread.DescendantsAndSelf(),
            e => (string?)e.Attribute("id") == "comment-2");
        Assert.Contains("Alice", rootThread.Value);
        Assert.Contains("2024-01-02T03:04:05Z", rootThread.Value);
        Assert.Contains("Root review body.", rootThread.Value);
        Assert.Contains("Bob", rootThread.Value);
        Assert.Contains("Reply review body.", rootThread.Value);
        Assert.Contains("Resolved", rootThread.Value);
        Assert.Contains("Open", rootThread.Value);
        Assert.Equal("#comment-ref-1", (string?)Assert.Single(ById(rootThread, "comment-2")
            .Descendants(), e => HasClass(e, "comment-backref")).Attribute("href"));

        for (var id = 1; id <= 6; id++)
            Assert.Single(html.Descendants(), e => (string?)e.Attribute("id") == $"comment-ref-{id}");
    }

    [Fact]
    public void Margin_profile_uses_one_page_selectable_note_per_thread_and_nests_replies()
    {
        var html = Render(CommentRenderMode.Margin, renderComments: true);
        var column = Assert.Single(html.Descendants(), e => HasClass(e, "comment-margin-column"));
        var selectableRoots = column.Descendants()
            .Where(e => e.Attribute("data-comment-id") != null)
            .ToList();

        Assert.Equal(new[] { "3", "1", "4", "5", "6" },
            selectableRoots.Select(e => (string)e.Attribute("data-comment-id")!).ToArray());
        var root = ById(column, "comment-1");
        var reply = ById(root, "comment-2");
        Assert.Null(reply.Attribute("data-comment-id"));
        Assert.Equal("2", (string?)reply.Attribute("data-comment-node-id"));
        Assert.True(HasClass(reply, "comment-margin-reply"));
        Assert.Contains("Reply review body.", reply.Value);
        Assert.Equal("#comment-ref-1", (string?)Assert.Single(reply.Descendants(),
            e => HasClass(e, "comment-margin-backref")).Attribute("href"));
        Assert.Contains("Status unknown", ById(column, "comment-3").Value);

        // A reply range/reference selects its root thread for paginated margin placement.
        var replyRangeOrMarker = html.Descendants()
            .Where(e => (string?)e.Attribute("data-comment-id") == "1")
            .Where(e => !e.AncestorsAndSelf().Any(a => HasClass(a, "comment-margin-column")))
            .ToList();
        Assert.NotEmpty(replyRangeOrMarker);
    }

    [Fact]
    public void Hidden_profile_removes_comment_presentation_but_preserves_story_text()
    {
        var html = Render(CommentRenderMode.EndnoteStyle, renderComments: false);
        var text = html.Value;

        Assert.DoesNotContain("comment-marker", html.ToString(SaveOptions.DisableFormatting));
        Assert.DoesNotContain("data-comment-status", html.ToString(SaveOptions.DisableFormatting));
        Assert.DoesNotContain("Root review body.", text);
        Assert.DoesNotContain("Reply review body.", text);
        Assert.Contains("BODY TOKEN", text);
        Assert.Contains("HEADER TOKEN", text);
        Assert.Contains("FOOTER TOKEN", text);
        Assert.Contains("FOOTNOTE TOKEN", text);
        Assert.Contains("ENDNOTE TOKEN", text);
    }

    [Fact]
    public void Markup_format_description_honors_explicit_false_toggle()
    {
        var html = Render(CommentRenderMode.EndnoteStyle, renderComments: false);
        var formatChange = Assert.Single(html.Descendants().Where(e => HasClass(e, "rev-format-change")));

        Assert.Contains("Bold added", (string?)formatChange.Attribute("title"));
    }

    [Fact]
    public void Markup_comment_body_preserves_visible_revision_authorship_and_date()
    {
        var html = Render(CommentRenderMode.EndnoteStyle, renderComments: true);
        var root = ById(html, "comment-1");
        var deleted = Assert.Single(root.Descendants(), e => e.Name.LocalName == "del"
            && e.Value.Contains("COMMENT BODY ORIGINAL", StringComparison.Ordinal));
        var inserted = Assert.Single(root.Descendants(), e => e.Name.LocalName == "ins"
            && e.Value.Contains("COMMENT BODY FINAL", StringComparison.Ordinal));

        Assert.Equal("Comment Editor", (string?)deleted.Attribute("data-author"));
        Assert.Equal("2024-01-08T09:10:11Z", (string?)deleted.Attribute("data-date"));
        Assert.Equal("Comment Editor", (string?)inserted.Attribute("data-author"));
        Assert.Equal("2024-01-08T09:10:11Z", (string?)inserted.Attribute("data-date"));
    }

    [Fact]
    public void Ambiguous_comment_identities_emit_inert_diagnostics()
    {
        var html = Render(CommentRenderMode.EndnoteStyle, renderComments: true,
            BuildIdentityCollisionFixture());
        var diagnostics = html.Descendants()
            .Where(element => (string?)element.Attribute("name")
                == "docxodus-comment-topology")
            .Select(element => (string?)element.Attribute("data-comment-topology"))
            .ToList();

        Assert.Contains("duplicate_comment_id", diagnostics);
        Assert.Contains("duplicate_paragraph_id", diagnostics);
        Assert.Contains("duplicate_thread_metadata", diagnostics);

        var first = ById(html, "comment-1");
        Assert.Equal("First", (string?)first.Attribute("data-author"));
        Assert.Equal("open", (string?)first.Attribute("data-comment-status"));
        Assert.Null(first.Attribute("data-comment-parent-id"));
        Assert.Contains("First.", first.Value);
        Assert.DoesNotContain("Duplicate id.", first.Value);

        var second = ById(html, "comment-2");
        Assert.Equal("Second", (string?)second.Attribute("data-author"));
        Assert.Equal("unknown", (string?)second.Attribute("data-comment-status"));
        Assert.Null(second.Attribute("data-comment-parent-id"));
        Assert.Contains("Duplicate paragraph.", second.Value);
    }

    [Fact]
    public void Malformed_reply_topology_renders_as_auditable_roots_without_recursion()
    {
        var malformedCommentsExtended = """
            <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
              <w15:commentEx w15:paraId="11111111" w15:paraIdParent="22222222" w15:done="1"/>
              <w15:commentEx w15:paraId="22222222" w15:paraIdParent="11111111" w15:done="0"/>
              <w15:commentEx w15:paraId="33333333" w15:paraIdParent="99999999"/>
            </w15:commentsEx>
            """;
        var html = Render(CommentRenderMode.EndnoteStyle, renderComments: true,
            BuildFixture(malformedCommentsExtended));

        foreach (var id in new[] { "1", "2" })
        {
            var cyclic = ById(html, $"comment-{id}");
            Assert.Equal("cyclic_parent", (string?)cyclic.Attribute("data-comment-topology"));
            Assert.Null(cyclic.Attribute("data-comment-parent-id"));
            Assert.Matches("^[12]{8}$",
                (string?)cyclic.Attribute("data-comment-parent-para-id") ?? string.Empty);
        }

        var orphan = ById(html, "comment-3");
        Assert.Equal("orphaned_parent", (string?)orphan.Attribute("data-comment-topology"));
        Assert.Equal("99999999", (string?)orphan.Attribute("data-comment-parent-para-id"));
        Assert.Null(orphan.Attribute("data-comment-parent-id"));
        Assert.Contains("Root review body.", html.Value);
        Assert.Contains("Reply review body.", html.Value);
        Assert.Contains("Header review body.", html.Value);
    }

    [Fact]
    public void Hidden_profile_retains_inert_malformed_topology_evidence()
    {
        var malformedCommentsExtended = """
            <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
              <w15:commentEx w15:paraId="11111111" w15:paraIdParent="22222222"/>
              <w15:commentEx w15:paraId="22222222" w15:paraIdParent="11111111"/>
              <w15:commentEx w15:paraId="33333333" w15:paraIdParent="99999999"/>
            </w15:commentsEx>
            """;
        var html = Render(CommentRenderMode.EndnoteStyle, renderComments: false,
            BuildFixture(malformedCommentsExtended));

        var diagnostics = html.Descendants()
            .Where(element => (string?)element.Attribute("name")
                == "docxodus-comment-topology")
            .ToList();
        Assert.Equal(3, diagnostics.Count);
        Assert.Equal(new[] { "1", "2", "3" }, diagnostics
            .Select(element => (string)element.Attribute("data-comment-node-id")!)
            .ToArray());
        Assert.All(diagnostics, element => Assert.Equal("/word/commentsExtended.xml",
            (string?)element.Attribute("data-comment-part-uri")));
        Assert.Equal(2, diagnostics.Count(element =>
            (string?)element.Attribute("data-comment-topology") == "cyclic_parent"));
        Assert.Single(diagnostics, element =>
            (string?)element.Attribute("data-comment-topology") == "orphaned_parent");
        Assert.DoesNotContain("Root review body.", html.Value);
        Assert.DoesNotContain(html.Descendants(), element =>
            element.Attribute("data-comment-status") != null);
    }

    [Theory]
    [InlineData(CommentRenderMode.EndnoteStyle)]
    [InlineData(CommentRenderMode.Inline)]
    [InlineData(CommentRenderMode.Margin)]
    public void Deep_valid_reply_chain_renders_iteratively(CommentRenderMode mode)
    {
        const int depth = 768;
        var html = Render(mode, renderComments: true, BuildDeepReplyFixture(depth));

        Assert.Equal(depth, html.DescendantsAndSelf().Count(element =>
            element.Attribute("data-comment-node-id") != null));
        Assert.Contains($"Deep reply {depth}.", html.ToString(SaveOptions.DisableFormatting));
    }

    private static XElement Render(CommentRenderMode mode, bool renderComments, byte[]? fixture = null)
    {
        using var stream = new MemoryStream();
        fixture ??= BuildFixture();
        stream.Write(fixture, 0, fixture.Length);
        stream.Position = 0;
        using var document = WordprocessingDocument.Open(stream, true);
        return WmlToHtmlConverter.ConvertToHtml(document, new WmlToHtmlConverterSettings
        {
            PageTitle = "Review comment profiles",
            FabricateCssClasses = false,
            RenderComments = renderComments,
            CommentRenderMode = mode,
            IncludeCommentMetadata = true,
            RenderHeadersAndFooters = true,
            RenderFootnotesAndEndnotes = true,
            RenderTrackedChanges = true,
            IncludeRevisionMetadata = true,
        });
    }

    private static XElement ById(XElement root, string id) =>
        Assert.Single(root.DescendantsAndSelf(), e => (string?)e.Attribute("id") == id);

    private static bool HasClass(XElement element, string className) =>
        ((string?)element.Attribute("class") ?? string.Empty)
            .Split(' ', StringSplitOptions.RemoveEmptyEntries)
            .Contains(className, StringComparer.Ordinal);

    private static string StoryParagraph(int commentId, string token, bool includeRange = true) =>
        $"""
        <w:p>
          {(includeRange ? $"<w:commentRangeStart w:id=\"{commentId}\"/>" : string.Empty)}
          <w:r><w:t>{token}</w:t></w:r>
          {(includeRange ? $"<w:commentRangeEnd w:id=\"{commentId}\"/>" : string.Empty)}
          <w:r><w:commentReference w:id="{commentId}"/></w:r>
        </w:p>
        """;

    private static byte[] BuildFixture(string? commentsExtendedXml = null)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(
                   stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            WriteXml(main, $"""
                <w:document xmlns:w="{W}" xmlns:r="{R}">
                  <w:body>
                    {StoryParagraph(1, "BODY TOKEN")}
                    {StoryParagraph(2, "REPLY REFERENCE", includeRange: false)}
                    <w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="30" w:author="Formatter" w:date="2024-01-06T07:08:09Z"><w:rPr><w:b w:val="0"/></w:rPr></w:rPrChange></w:rPr><w:t>FORMAT TOKEN</w:t></w:r></w:p>
                    <w:p>
                      <w:r><w:t xml:space="preserve">Notes </w:t></w:r>
                      <w:r><w:footnoteReference w:id="1"/></w:r>
                      <w:r><w:endnoteReference w:id="1"/></w:r>
                    </w:p>
                    <w:sectPr>
                      <w:headerReference w:type="default" r:id="rIdHeader"/>
                      <w:footerReference w:type="default" r:id="rIdFooter"/>
                    </w:sectPr>
                  </w:body>
                </w:document>
                """);

            var styles = main.AddNewPart<StyleDefinitionsPart>();
            WriteXml(styles, $"""
                <w:styles xmlns:w="{W}">
                  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
                    <w:name w:val="Normal"/>
                  </w:style>
                </w:styles>
                """);

            var header = main.AddNewPart<HeaderPart>("rIdHeader");
            WriteXml(header, $"<w:hdr xmlns:w=\"{W}\">{StoryParagraph(3, "HEADER TOKEN")}</w:hdr>");
            var footer = main.AddNewPart<FooterPart>("rIdFooter");
            WriteXml(footer, $"<w:ftr xmlns:w=\"{W}\">{StoryParagraph(4, "FOOTER TOKEN")}</w:ftr>");

            var footnotes = main.AddNewPart<FootnotesPart>();
            WriteXml(footnotes, $"""
                <w:footnotes xmlns:w="{W}">
                  <w:footnote w:id="1">{StoryParagraph(5, "FOOTNOTE TOKEN")}</w:footnote>
                </w:footnotes>
                """);
            var endnotes = main.AddNewPart<EndnotesPart>();
            WriteXml(endnotes, $"""
                <w:endnotes xmlns:w="{W}">
                  <w:endnote w:id="1">{StoryParagraph(6, "ENDNOTE TOKEN")}</w:endnote>
                </w:endnotes>
                """);

            var comments = main.AddNewPart<WordprocessingCommentsPart>();
            WriteXml(comments, $"""
                <w:comments xmlns:w="{W}"
                  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
                  <w:comment w:id="1" w:author="Alice" w:date="2024-01-02T03:04:05Z">
                    <w:p w14:paraId="11111111"><w:r><w:t>Root review body.</w:t></w:r>
                      <w:del w:id="40" w:author="Comment Editor" w:date="2024-01-08T09:10:11Z"><w:r><w:delText>COMMENT BODY ORIGINAL</w:delText></w:r></w:del>
                      <w:ins w:id="41" w:author="Comment Editor" w:date="2024-01-08T09:10:11Z"><w:r><w:t>COMMENT BODY FINAL</w:t></w:r></w:ins>
                    </w:p>
                  </w:comment>
                  <w:comment w:id="2" w:author="Bob" w:date="2024-01-03T04:05:06Z">
                    <w:p w14:paraId="22222222"><w:r><w:t>Reply review body.</w:t></w:r></w:p>
                  </w:comment>
                  <w:comment w:id="3" w:author="Carol" w:date="2024-01-04T05:06:07Z">
                    <w:p w14:paraId="33333333"><w:r><w:t>Header review body.</w:t></w:r></w:p>
                  </w:comment>
                  <w:comment w:id="4" w:author="Dan" w:date="2024-01-05T06:07:08Z">
                    <w:p w14:paraId="44444444"><w:r><w:t>Footer review body.</w:t></w:r></w:p>
                  </w:comment>
                  <w:comment w:id="5" w:author="Eve" w:date="2024-01-06T07:08:09Z">
                    <w:p w14:paraId="55555555"><w:r><w:t>Footnote review body.</w:t></w:r></w:p>
                  </w:comment>
                  <w:comment w:id="6" w:author="Frank" w:date="2024-01-07T08:09:10Z">
                    <w:p w14:paraId="66666666"><w:r><w:t>Endnote review body.</w:t></w:r></w:p>
                  </w:comment>
                </w:comments>
                """);

            var commentsEx = main.AddNewPart<WordprocessingCommentsExPart>();
            WriteXml(commentsEx, commentsExtendedXml ?? """
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w15:commentEx w15:paraId="11111111" w15:done="1"/>
                  <w15:commentEx w15:paraId="22222222" w15:paraIdParent="11111111" w15:done="0"/>
                  <w15:commentEx w15:paraId="44444444" w15:done="0"/>
                  <w15:commentEx w15:paraId="55555555" w15:done="true"/>
                  <w15:commentEx w15:paraId="66666666" w15:done="unrecognized"/>
                </w15:commentsEx>
                """);
        }
        return stream.ToArray();
    }

    private static byte[] BuildIdentityCollisionFixture()
    {
        using var stream = new MemoryStream();
        var fixture = BuildFixture();
        stream.Write(fixture, 0, fixture.Length);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            WriteXml(document.MainDocumentPart!.WordprocessingCommentsPart!, $"""
                <w:comments xmlns:w="{W}"
                  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
                  <w:comment w:id="1" w:author="First"><w:p w14:paraId="AAAAAAAA"><w:r><w:t>First.</w:t></w:r></w:p></w:comment>
                  <w:comment w:id="1" w:author="Duplicate"><w:p w14:paraId="BBBBBBBB"><w:r><w:t>Duplicate id.</w:t></w:r></w:p></w:comment>
                  <w:comment w:id="2" w:author="Second"><w:p w14:paraId="AAAAAAAA"><w:r><w:t>Duplicate paragraph.</w:t></w:r></w:p></w:comment>
                </w:comments>
                """);
            WriteXml(document.MainDocumentPart.WordprocessingCommentsExPart!, """
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  <w15:commentEx w15:paraId="AAAAAAAA" w15:done="0"/>
                  <w15:commentEx w15:paraId="BBBBBBBB" w15:paraIdParent="AAAAAAAA" w15:done="1"/>
                  <w15:commentEx w15:paraId="AAAAAAAA" w15:done="1"/>
                </w15:commentsEx>
                """);
        }
        return stream.ToArray();
    }

    private static byte[] BuildDeepReplyFixture(int depth)
    {
        var comments = new StringBuilder();
        var commentsExtended = new StringBuilder();
        for (var id = 1; id <= depth; id++)
        {
            var paraId = id.ToString("X8");
            comments.Append($"<w:comment w:id=\"{id}\" w:author=\"Depth {id}\"><w:p w14:paraId=\"{paraId}\"><w:r><w:t>Deep reply {id}.</w:t></w:r></w:p></w:comment>");
            commentsExtended.Append($"<w15:commentEx w15:paraId=\"{paraId}\"");
            if (id > 1)
                commentsExtended.Append($" w15:paraIdParent=\"{(id - 1).ToString("X8")}\"");
            commentsExtended.Append("/>");
        }

        using var stream = new MemoryStream();
        var fixture = BuildFixture();
        stream.Write(fixture, 0, fixture.Length);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            WriteXml(document.MainDocumentPart!.WordprocessingCommentsPart!, $"""
                <w:comments xmlns:w="{W}"
                  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
                  {comments}
                </w:comments>
                """);
            WriteXml(document.MainDocumentPart.WordprocessingCommentsExPart!, $"""
                <w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
                  {commentsExtended}
                </w15:commentsEx>
                """);
        }
        return stream.ToArray();
    }

    private static void WriteXml(OpenXmlPart part, string xml)
    {
        using var writer = new StreamWriter(part.GetStream(FileMode.Create), new UTF8Encoding(false));
        writer.Write(xml);
    }
}
