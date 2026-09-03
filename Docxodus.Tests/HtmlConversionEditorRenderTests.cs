#nullable enable

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Internal;
using Wp = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// The browser editor's comments-aware render profile
/// (<see cref="DocxSessionOps.RenderEditorHtml"/> / <see cref="DocxSessionOps.RenderEditorBlocksHtml"/>
/// / <see cref="DocxSessionOps.RenderEditorBlockHtml"/>): Inline comment markup on the per-block
/// path matches the full render, the comment-less profile stays byte-identical to the pinned
/// <c>RenderBlockHtml</c> output, cross-block ranges highlight in isolation, and the shell tracks
/// comment mutations. Test IDs use the HER range.
/// </summary>
public class HtmlConversionEditorRenderTests
{
    private const string WithComments = "{\"comments\":true}";
    private const string WithoutComments = "{\"comments\":false}";

    /// <summary>Three plain body paragraphs, no comments part.</summary>
    private static byte[] BuildThreeParagraphs()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(new Wp.Body(
                new Wp.Paragraph(new Wp.Run(new Wp.Text("Alpha one"))),
                new Wp.Paragraph(new Wp.Run(new Wp.Text("Beta two"))),
                new Wp.Paragraph(new Wp.Run(new Wp.Text("Gamma three")))));
        }
        return ms.ToArray();
    }

    /// <summary>Three paragraphs where comment 7 opens in the first and closes in the second —
    /// Word's shape for a comment dragged across a paragraph boundary.</summary>
    private static byte[] BuildCrossBlockComment()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(new Wp.Body(
                new Wp.Paragraph(
                    new Wp.CommentRangeStart { Id = "7" },
                    new Wp.Run(new Wp.Text("Alpha one"))),
                new Wp.Paragraph(
                    new Wp.Run(new Wp.Text("Beta two")),
                    new Wp.CommentRangeEnd { Id = "7" },
                    new Wp.Run(new Wp.CommentReference { Id = "7" })),
                new Wp.Paragraph(new Wp.Run(new Wp.Text("Gamma three")))));
            var comments = main.AddNewPart<WordprocessingCommentsPart>();
            comments.Comments = new Wp.Comments(
                new Wp.Comment(new Wp.Paragraph(new Wp.Run(new Wp.Text("Spans two paragraphs"))))
                {
                    Id = "7", Author = "Reviewer", Initials = "R",
                });
        }
        return ms.ToArray();
    }

    private static string[] BodyParagraphAnchors(DocxSession session) =>
        session.ListBlocks().Body.Where(u => u.Kind is "p" or "h").Select(u => u.Id).ToArray();

    private static string Unid(string anchorId) => anchorId.Substring(anchorId.LastIndexOf(':') + 1);

    private static XElement FullRenderBlock(string html, string anchorId)
    {
        var full = XElement.Parse(html);
        return full.Descendants().First(e => (string?)e.Attribute("data-anchor") == Unid(anchorId));
    }

    /// <summary>Attribute order has no semantics; a namespace declaration lives on whichever
    /// ancestor happened to serialize it (the full document's root vs a standalone fragment's own
    /// element); and the block path (StampAnchors off at the options level) never stamps
    /// <c>data-source-anchor-id</c> on BODY blocks — the editor's incremental swap path has always
    /// rendered without it. Everything else, including the comment-story identities the
    /// highlight spans carry, is compared exactly.</summary>
    private static XElement Canonical(XElement e)
    {
        var clone = new XElement(e);
        foreach (var el in clone.DescendantsAndSelf())
        {
            var attrs = el.Attributes()
                .Where(a => !a.IsNamespaceDeclaration)
                .Where(a => a.Name.LocalName != "data-source-anchor-id" || el.Name.LocalName == "span")
                .OrderBy(a => a.Name.ToString(), System.StringComparer.Ordinal)
                .ToList();
            el.RemoveAttributes();
            el.Add(attrs);
        }
        return clone;
    }

    private static int SingleCommentId(DocxSession session) => session.ListComments().Single().Id;

    /// <summary>Canonical DeepEquals with a message that pinpoints the first differing character
    /// of the unformatted serializations — the indented forms two fragments print with can hide a
    /// whitespace-only text node or an entity difference.</summary>
    private static void AssertSameFragment(XElement expected, XElement actual, string what)
    {
        var e = Canonical(expected);
        var a = Canonical(actual);
        if (XNode.DeepEquals(e, a)) return;
        var es = e.ToString(SaveOptions.DisableFormatting);
        var @as = a.ToString(SaveOptions.DisableFormatting);
        int i = 0;
        while (i < es.Length && i < @as.Length && es[i] == @as[i]) i++;
        int from = System.Math.Max(0, i - 60);
        Assert.Fail($"{what} differs at char {i}:\nexpected …{es.Substring(from, System.Math.Min(160, es.Length - from))}\nactual   …{@as.Substring(from, System.Math.Min(160, @as.Length - from))}");
    }

    [Fact]
    public void HER001_EditorBlockRender_ShowsHighlightAndMatchesFullRender()
    {
        int handle = DocxSessionOps.OpenSession(BuildThreeParagraphs(), null);
        try
        {
            var session = SessionRegistry.Get(handle);
            var anchors = BodyParagraphAnchors(session);
            Assert.True(session.AddComment(anchors[1], new CharSpan(0, 4), "Reviewer", "Look here").Success);
            int id = SingleCommentId(session);

            var block = DocxSessionOps.RenderEditorBlockHtml(handle, anchors[1], WithComments);
            Assert.StartsWith("<", block);
            Assert.Contains("comment-highlight", block);
            Assert.Contains($"data-comment-id=\"{id}\"", block);
            Assert.Contains("comment-marker", block);

            var full = DocxSessionOps.RenderEditorHtml(handle, WithComments);
            Assert.StartsWith("<", full);
            AssertSameFragment(FullRenderBlock(full, anchors[1]), XElement.Parse(block), "commented block");

            // The batch export agrees with the single-block one.
            var batch = DocxSessionOps.RenderEditorBlocksHtml(
                handle, $"[\"{anchors[1]}\"]", WithComments);
            using var map = System.Text.Json.JsonDocument.Parse(batch);
            Assert.Equal(block, map.RootElement.GetProperty(anchors[1]).GetString());
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    [Fact]
    public void HER002_CommentsOff_IsByteIdenticalToTheExistingRenders()
    {
        int handle = DocxSessionOps.OpenSession(BuildThreeParagraphs(), null);
        try
        {
            var session = SessionRegistry.Get(handle);
            var anchors = BodyParagraphAnchors(session);
            Assert.True(session.AddComment(anchors[1], new CharSpan(0, 4), "Reviewer", "Look here").Success);

            var plainBlock = DocxSessionOps.RenderBlockHtml(handle, anchors[1], "docx-", false);
            Assert.DoesNotContain("comment-", plainBlock);
            Assert.Equal(plainBlock, DocxSessionOps.RenderEditorBlockHtml(handle, anchors[1], WithoutComments));
            Assert.Equal(plainBlock, DocxSessionOps.RenderEditorBlockHtml(handle, anchors[1], ""));

            var plainBatch = DocxSessionOps.RenderBlocksHtml(handle, $"[\"{anchors[0]}\",\"{anchors[1]}\"]", "docx-", false);
            Assert.Equal(plainBatch, DocxSessionOps.RenderEditorBlocksHtml(
                handle, $"[\"{anchors[0]}\",\"{anchors[1]}\"]", WithoutComments));

            var plainFull = DocxSessionOps.RenderHtml(handle, "docx-", false, false, 1);
            Assert.Equal(plainFull, DocxSessionOps.RenderEditorHtml(handle, WithoutComments));
            Assert.Equal(
                DocxSessionOps.RenderHtml(handle, "docx-", false, false, 1, renderTrackedChanges: true),
                DocxSessionOps.RenderEditorHtml(handle, "{\"renderTrackedChanges\":true}"));
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    [Fact]
    public void HER003_CrossBlockRange_HighlightsTheSecondParagraphInIsolation()
    {
        int handle = DocxSessionOps.OpenSession(BuildCrossBlockComment(), null);
        try
        {
            var session = SessionRegistry.Get(handle);
            var anchors = BodyParagraphAnchors(session);
            Assert.Equal(7, SingleCommentId(session));

            // The second paragraph holds only the range END: without the reconstructed start the
            // isolated shell body would render it unhighlighted.
            var second = DocxSessionOps.RenderEditorBlockHtml(handle, anchors[1], WithComments);
            Assert.Contains("comment-highlight", second);
            Assert.Contains("data-comment-id=\"7\"", second);
            Assert.Contains("Beta two", second);

            var full = DocxSessionOps.RenderEditorHtml(handle, WithComments);
            AssertSameFragment(FullRenderBlock(full, anchors[1]), XElement.Parse(second), "second paragraph");

            // The first paragraph (holds the start) and the third (outside the range) agree too,
            // and the synthetic end keeps the third from inheriting an open range.
            var first = DocxSessionOps.RenderEditorBlockHtml(handle, anchors[0], WithComments);
            Assert.Contains("comment-highlight", first);
            AssertSameFragment(FullRenderBlock(full, anchors[0]), XElement.Parse(first), "first paragraph");
            var third = DocxSessionOps.RenderEditorBlockHtml(handle, anchors[2], WithComments);
            Assert.DoesNotContain("comment-highlight", third);

            // A batch of the second and third alone reconstructs the start for the run and closes
            // it cleanly before the next block.
            var batch = DocxSessionOps.RenderEditorBlocksHtml(
                handle, $"[\"{anchors[2]}\",\"{anchors[1]}\"]", WithComments);
            using var map = System.Text.Json.JsonDocument.Parse(batch);
            Assert.Contains("comment-highlight", map.RootElement.GetProperty(anchors[1]).GetString());
            Assert.DoesNotContain("comment-highlight", map.RootElement.GetProperty(anchors[2]).GetString());
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    [Fact]
    public void HER005_FooterBlockRender_StampsPageFieldMarkers()
    {
        int handle = DocxSessionOps.OpenSession(BuildThreeParagraphs(), null);
        try
        {
            var session = SessionRegistry.Get(handle);
            var body = BodyParagraphAnchors(session)[0];
            var footer = session.SetFooterText(body, HeaderFooterKind.Default, "Page ");
            Assert.True(footer.Success, footer.Error?.Message);
            var footerPara = footer.Created[0].Id;
            Assert.True(session.InsertPageNumberField(footerPara, PageNumberField.CurrentPage).Success);
            Assert.True(session.InsertPageNumberField(footerPara, PageNumberField.TotalPages, NumberFormat.LowerRoman).Success);

            // The editor profile in page view carries the paginated full render's markers …
            var paged = DocxSessionOps.RenderEditorBlockHtml(handle, footerPara, "{\"paginated\":true}");
            Assert.Contains("data-field=\"PAGE\"", paged);
            Assert.Contains("data-field=\"NUMPAGES\"", paged);
            Assert.Contains("data-field-format=\"roman\"", paged);

            // … and so does a running-story block even when the profile says continuous.
            var continuous = DocxSessionOps.RenderEditorBlockHtml(handle, footerPara, "{}");
            Assert.Contains("data-field=\"PAGE\"", continuous);

            // A body block in continuous view does not (it matches the continuous full render),
            // and the pre-existing block render is untouched either way.
            Assert.True(session.InsertPageNumberField(body, PageNumberField.CurrentPage).Success);
            Assert.DoesNotContain("data-field", DocxSessionOps.RenderEditorBlockHtml(handle, body, "{}"));
            Assert.Contains("data-field=\"PAGE\"", DocxSessionOps.RenderEditorBlockHtml(handle, body, "{\"paginated\":true}"));
            Assert.DoesNotContain("data-field", DocxSessionOps.RenderBlockHtml(handle, footerPara, "docx-", false));
            Assert.DoesNotContain("data-field", DocxSessionOps.RenderBlockHtml(handle, body, "docx-", false));

            // "Page X of Y" renders its literal runs with the spaces and marks both fields.
            var pageOf = session.SetFooterText(body, HeaderFooterKind.First, "");
            Assert.True(pageOf.Success, pageOf.Error?.Message);
            var firstPara = pageOf.Created[0].Id;
            Assert.True(session.InsertPageNumberField(firstPara, PageNumberField.PageOfTotal).Success);
            var composite = DocxSessionOps.RenderEditorBlockHtml(handle, firstPara, "{\"paginated\":true}");
            // The converter renders a trailing space as U+00A0 so it survives HTML whitespace collapsing.
            Assert.Matches("Page[ \u00A0]</span>", composite);
            Assert.Contains("data-field=\"PAGE\"", composite);
            Assert.Contains("data-field=\"NUMPAGES\"", composite);
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    [Fact]
    public void HER004_ShellRebuildsWhenCommentsChange()
    {
        int handle = DocxSessionOps.OpenSession(BuildThreeParagraphs(), null);
        try
        {
            var session = SessionRegistry.Get(handle);
            var anchors = BodyParagraphAnchors(session);

            // First render builds the shell over a document with no comments part at all.
            var before = DocxSessionOps.RenderEditorBlockHtml(handle, anchors[1], WithComments);
            Assert.DoesNotContain("comment-highlight", before);
            Assert.NotNull(session.RenderShellDoc);
            long sigBefore = session.RenderShellSignature;

            Assert.True(session.AddComment(anchors[1], new CharSpan(0, 4), "Reviewer", "Now").Success);
            var after = DocxSessionOps.RenderEditorBlockHtml(handle, anchors[1], WithComments);
            Assert.Contains("comment-highlight", after);
            Assert.NotEqual(sigBefore, session.RenderShellSignature);

            // A text edit elsewhere reuses the shell; a comment edit (tooltip text) rebuilds it.
            long sigAfterComment = session.RenderShellSignature;
            Assert.True(session.ReplaceText(anchors[2], "Gamma edited").Success);
            _ = DocxSessionOps.RenderEditorBlockHtml(handle, anchors[2], WithComments);
            Assert.Equal(sigAfterComment, session.RenderShellSignature);

            var cmt = session.ListComments().Single().DefAnchorId;
            Assert.True(session.UpdateComment(cmt, "Changed body").Success);
            var updated = DocxSessionOps.RenderEditorBlockHtml(handle, anchors[1], WithComments);
            Assert.Contains("Changed body", updated);
            Assert.NotEqual(sigAfterComment, session.RenderShellSignature);

            // Undo of the comment restores the comment-less render (snapshot restore bumps too).
            Assert.True(session.Undo());
            Assert.True(session.Undo());
            Assert.True(session.Undo());
            var undone = DocxSessionOps.RenderEditorBlockHtml(handle, anchors[1], WithComments);
            Assert.DoesNotContain("comment-highlight", undone);
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }

    /// <summary>
    /// Browsers parse the converter's output as HTML, where <c>&lt;span /&gt;</c> is an OPEN tag.
    /// A complex field's begin/separate/end runs render as empty spans, so a "Page X of Y" footer
    /// serialized with XHTML self-closing spans parsed with " of " and the NUMPAGES field nested
    /// inside the PAGE field — and the paginator's per-page substitution then wiped them. Empty
    /// non-void elements must serialize as pairs; void elements keep their self-closing form.
    /// </summary>
    [Fact]
    public void HER006_EmptyElementsSerializeAsPairs_SoFieldsStaySiblingsInHtml()
    {
        int handle = DocxSessionOps.OpenSession(BuildThreeParagraphs(), null);
        try
        {
            var session = SessionRegistry.Get(handle);
            var body = BodyParagraphAnchors(session)[0];
            var footer = session.SetFooterText(body, HeaderFooterKind.Default, "");
            Assert.True(footer.Success, footer.Error?.Message);
            var footerPara = footer.Created[0].Id;
            Assert.True(session.InsertPageNumberField(footerPara, PageNumberField.PageOfTotal).Success);

            var block = DocxSessionOps.RenderEditorBlockHtml(handle, footerPara, "{\"paginated\":true}");
            Assert.DoesNotMatch("<span( [^>]*)?/>", block);
            Assert.Matches("<span( [^>]*)?></span>", block);
            // Sibling order survives an HTML parse: the PAGE field closes before " of " begins.
            var page = block.IndexOf("data-field=\"PAGE\"", System.StringComparison.Ordinal);
            // The converter may render the literal's spaces as U+00A0 so they survive collapsing.
            var ofMatch = System.Text.RegularExpressions.Regex.Match(block, "[ \u00A0]of[ \u00A0]</span>");
            var of = ofMatch.Success ? ofMatch.Index : -1;
            var numPages = block.IndexOf("data-field=\"NUMPAGES\"", System.StringComparison.Ordinal);
            Assert.True(page >= 0 && of > page && numPages > of, block);
            Assert.DoesNotMatch("<(span|p|div|a|sup|sub|td|th|tr|table|ins|del)( [^>]*)?/>", block);

            // The whole-document render (what the paginator clones from) is normalized too, and
            // void elements are left alone.
            var full = DocxSessionOps.RenderEditorHtml(handle, "{\"paginated\":true}");
            Assert.DoesNotMatch("<(span|p|div|a|sup|sub|td|th|tr|table|ins|del)( [^>]*)?/>", full);
            Assert.Contains("<meta charset=\"UTF-8\" />", full);
        }
        finally
        {
            DocxSessionOps.CloseSession(handle);
        }
    }
}
