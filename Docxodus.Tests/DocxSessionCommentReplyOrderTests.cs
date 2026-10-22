using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Xunit;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace Docxodus.Tests;

/// <summary>
/// A comment thread's replies keep the order they were added in, both as body-side
/// <c>w:commentReference</c> positions and in the rendered thread (which sorts by reference position).
/// </summary>
public class DocxSessionCommentReplyOrderTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static readonly string[] Replies = { "First reply.", "Second reply.", "Third reply." };

    /// <summary>
    /// A root comment plus three replies added one at a time, each to the root; with
    /// <paramref name="nestedReply"/>, a final reply to the first reply (still part of the root's thread).
    /// </summary>
    private static byte[] BuildThread(bool nestedReply = false)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Wp.Document(new Wp.Body(
                new Wp.Paragraph(new Wp.Run(new Wp.Text("The seller delivers the goods.")))));
        }

        using var session = new DocxSession(ms.ToArray());
        var paragraph = session.Project().AnchorIndex.Values.Select(v => v.Anchor).First(a => a.Kind == "p").Id;
        var root = session.AddComment(paragraph, null, "Ann", "Root comment.");
        var cmt = root.Created.First(a => a.Kind == "cmt").Id;
        var first = "";
        foreach (var text in Replies)
        {
            var reply = session.AddCommentReply(cmt, "Ben", text);
            Assert.True(reply.Success);
            if (first == "") first = reply.Created.First(a => a.Kind == "cmt").Id;
        }
        if (nestedReply)
            Assert.True(session.AddCommentReply(first, "Cy", "Nested reply.").Success);
        return session.Save();
    }

    private static string[] ReferenceOrder(byte[] docx)
    {
        using var ms = ExpandableStream(docx);
        using var doc = WordprocessingDocument.Open(ms, false);
        var main = doc.MainDocumentPart!;
        var textById = main.WordprocessingCommentsPart!.GetXDocument().Root!.Elements(W + "comment")
            .ToDictionary(c => (string)c.Attribute(W + "id")!, c => c.Value);
        return main.GetXDocument().Descendants(W + "commentReference")
            .Select(r => textById[(string)r.Attribute(W + "id")!]).ToArray();
    }

    private static MemoryStream ExpandableStream(byte[] bytes)
    {
        var ms = new MemoryStream();
        ms.Write(bytes, 0, bytes.Length);
        return ms;
    }

    [Fact]
    public void AddCommentReply_KeepsReferencesInTheOrderRepliesWereAdded()
    {
        var order = ReferenceOrder(BuildThread());

        Assert.Equal(new[] { "Root comment." }.Concat(Replies), order);
    }

    [Fact]
    public void AddCommentReply_ToAReply_GoesAfterTheWholeThread()
    {
        Assert.Equal(new[] { "Root comment." }.Concat(Replies).Append("Nested reply."),
            ReferenceOrder(BuildThread(nestedReply: true)));
    }

    [Fact]
    public void AddCommentReply_RendersTheThreadInTheOrderRepliesWereAdded()
    {
        using var ms = ExpandableStream(BuildThread());
        using var doc = WordprocessingDocument.Open(ms, true);
        var html = WmlToHtmlConverter.ConvertToHtml(doc,
            new WmlToHtmlConverterSettings { RenderComments = true }).ToString();

        var order = new[] { "Root comment." }.Concat(Replies)
            .Select(t => Regex.Match(html, Regex.Escape(t)).Index).ToList();

        Assert.All(new[] { "Root comment." }.Concat(Replies), t => Assert.Contains(t, html));
        Assert.Equal(order.OrderBy(i => i), order);
    }
}
