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

namespace Docxodus.Tests.Ir.Diff;

/// <summary>
/// A comment definition belongs to comments.xml, whose relationship id namespace is independent from
/// document.xml and every other story. Right-added definitions must therefore import their own image and
/// hyperlink relationships; copying only the XML leaves r:embed/r:id dangling or pointing at a colliding
/// left-side target.
/// </summary>
[Trait("Category", "Markup")]
public class IrMarkupRendererCommentRelationshipTests
{
    private const string Wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string Rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string ImageRelationshipId = "rIdCommentImage";
    private const string HyperlinkRelationshipId = "rIdCommentLink";
    private const string LeftTarget = "https://example.com/left-comment-target";
    private const string RightTarget = "https://example.com/right-comment-target";
    private static readonly byte[] LeftImage = { 0x89, 0x50, 0x4E, 0x47, 0x11 };
    private static readonly byte[] RightImage = { 0x89, 0x50, 0x4E, 0x47, 0x22 };

    [Fact]
    public void Compare_RightAddedComment_ImportsCommentPartMediaAndHyperlinkRelationships()
    {
        var left = BuildDocument(hasRightAddedComment: false);
        var right = BuildDocument(hasRightAddedComment: true);

        var rendered = DocxDiff.Compare(left, right);

        AssertCommentRelationships(rendered, expectRelationshipIdsRemapped: true);
        AssertCommentRelationships(RevisionProcessor.AcceptRevisions(rendered), expectRelationshipIdsRemapped: true);
    }

    [Fact]
    public void Compare_RightAddedComment_CreatesCommentPartAndImportsItsRelationships()
    {
        var left = BuildDocument(hasRightAddedComment: false, includeCommentsPart: false);
        var right = BuildDocument(hasRightAddedComment: true);

        AssertCommentRelationships(DocxDiff.Compare(left, right), expectRelationshipIdsRemapped: false);
    }

    [Fact]
    public void Consolidate_RightAddedComment_ImportsReviewerCommentPartRelationships()
    {
        var baseDocument = BuildDocument(hasRightAddedComment: false);
        // Composite rendering adopts the base source for an Equal block. Give the reviewer a real text
        // revision so its comment marker is carried by a reviewer-sourced ModifyBlock; marker-only comment
        // additions are a separate composite-diff behavior from relationship import.
        var reviewer = BuildDocument(hasRightAddedComment: true, hasBodyRevision: true);

        var rendered = DocxDiff.Consolidate(baseDocument,
            new[] { new DocxDiffReviewer { Document = reviewer, Author = "Reviewer" } });

        AssertCommentRelationships(rendered, expectRelationshipIdsRemapped: true);
    }

    private static void AssertCommentRelationships(WmlDocument document, bool expectRelationshipIdsRemapped)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var wordDocument = WordprocessingDocument.Open(stream, false);
        var comments = wordDocument.MainDocumentPart!.WordprocessingCommentsPart!;
        var root = LoadXml(comments).Root!;
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace r = Rns;
        XNamespace w = Wns;

        var imageRelationshipId = (string?)root.Descendants(a + "blip").Single().Attribute(r + "embed");
        Assert.NotNull(imageRelationshipId);
        if (expectRelationshipIdsRemapped)
            Assert.NotEqual(ImageRelationshipId, imageRelationshipId);
        Assert.Equal(RightImage, ReadAllBytes(comments.GetPartById(imageRelationshipId!)));

        var hyperlink = root.Descendants(w + "hyperlink").Single();
        var hyperlinkRelationshipId = (string?)hyperlink.Attribute(r + "id");
        Assert.NotNull(hyperlinkRelationshipId);
        if (expectRelationshipIdsRemapped)
            Assert.NotEqual(HyperlinkRelationshipId, hyperlinkRelationshipId);
        Assert.Equal(RightTarget,
            comments.HyperlinkRelationships.Single(rel => rel.Id == hyperlinkRelationshipId).Uri.ToString());
    }

    private static WmlDocument BuildDocument(bool hasRightAddedComment, bool hasBodyRevision = false,
        bool includeCommentsPart = true)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            if (includeCommentsPart)
            {
                var comments = main.AddNewPart<WordprocessingCommentsPart>();
                var image = comments.AddNewPart<ImagePart>("image/png", ImageRelationshipId);
                using (var imageStream = image.GetStream(FileMode.Create, FileAccess.Write))
                {
                    var bytes = hasRightAddedComment ? RightImage : LeftImage;
                    imageStream.Write(bytes, 0, bytes.Length);
                }
                comments.AddHyperlinkRelationship(
                    new Uri(hasRightAddedComment ? RightTarget : LeftTarget, UriKind.Absolute), true,
                    HyperlinkRelationshipId);

                var commentXml = hasRightAddedComment
                    ? RightCommentXml()
                    : "<w:comment w:id=\"0\" w:author=\"Left\" w:date=\"2026-01-01T00:00:00Z\">" +
                      "<w:p><w:r><w:t>left-only definition</w:t></w:r></w:p></w:comment>";
                WriteXml(comments,
                    $"<w:comments xmlns:w=\"{Wns}\" xmlns:r=\"{Rns}\" " +
                    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
                    "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\">" +
                    commentXml + "</w:comments>");
            }

            var bodyText = hasBodyRevision ? "commented revised body" : "commented body";
            var bodyXml = hasRightAddedComment
                ? $"<w:p><w:commentRangeStart w:id=\"1\"/><w:r><w:t>{bodyText}</w:t></w:r>" +
                  "<w:commentRangeEnd w:id=\"1\"/><w:r><w:commentReference w:id=\"1\"/></w:r></w:p>"
                : "<w:p><w:r><w:t>commented body</w:t></w:r></w:p>";
            WriteXml(main, $"<w:document xmlns:w=\"{Wns}\"><w:body>{bodyXml}</w:body></w:document>");
        }
        return new WmlDocument("comment-relationships.docx", stream.ToArray());
    }

    private static string RightCommentXml() =>
        "<w:comment w:id=\"1\" w:author=\"Right\" w:date=\"2026-01-01T00:00:00Z\">" +
        "<w:p><w:r><w:drawing>" +
        "<wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">" +
        "<wp:extent cx=\"95250\" cy=\"95250\"/><wp:docPr id=\"1\" name=\"Comment image\"/>" +
        "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        "<pic:nvPicPr><pic:cNvPr id=\"1\" name=\"Comment image\"/><pic:cNvPicPr/></pic:nvPicPr>" +
        $"<pic:blipFill><a:blip r:embed=\"{ImageRelationshipId}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
        "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"95250\" cy=\"95250\"/></a:xfrm>" +
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>" +
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>" +
        $"<w:hyperlink r:id=\"{HyperlinkRelationshipId}\"><w:r><w:t>right comment hyperlink</w:t></w:r></w:hyperlink>" +
        "</w:p></w:comment>";

    private static XDocument LoadXml(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        return XDocument.Load(stream);
    }

    private static byte[] ReadAllBytes(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }

    private static void WriteXml(OpenXmlPart part, string xml)
    {
        using var stream = part.GetStream(FileMode.Create, FileAccess.Write);
        using var writer = new StreamWriter(stream);
        writer.Write(xml);
    }
}
