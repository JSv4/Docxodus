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
/// Relationship-closure regressions for right-sourced note content. Unlike body content, a footnote or
/// endnote owns its own relationships, so copying its XML without copying its owning part's relationships
/// leaves r:embed/r:id references dangling or pointed at a left-side collision.
/// </summary>
[Trait("Category", "Markup")]
public class IrMarkupRendererNoteRelationshipTests
{
    private const string Wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string Rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string ImageRelationshipId = "rIdNoteImage";
    private const string HyperlinkRelationshipId = "rIdNoteLink";
    private const string LeftTarget = "https://example.com/left-note-target";
    private const string RightTarget = "https://example.com/right-note-target";
    private static readonly byte[] OnePixelPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    [Fact]
    public void Render_inserted_footnote_image_imports_media_into_footnotes_part()
    {
        var left = BuildFootnoteImageDocument(includeFootnote: false);
        var right = BuildFootnoteImageDocument(includeFootnote: true);

        var rendered = DocxDiff.Compare(left, right);

        AssertFootnoteImageRelationship(rendered);
        AssertFootnoteImageRelationship(RevisionProcessor.AcceptRevisions(rendered));
    }

    [Fact]
    public void Render_inserted_endnote_hyperlink_remaps_part_scoped_rId_collision()
    {
        // Both note parts deliberately own rIdNoteLink, but point it at different targets. The LEFT part
        // is retained as the output baseline; the inserted RIGHT endnote must therefore receive a fresh
        // relationship id whose target is right-note-target, rather than resolving the old id to the left URI.
        var left = BuildEndnoteHyperlinkDocument(includeRightNote: false);
        var right = BuildEndnoteHyperlinkDocument(includeRightNote: true);

        var rendered = DocxDiff.Compare(left, right);
        var accepted = RevisionProcessor.AcceptRevisions(rendered);

        using var stream = new MemoryStream(accepted.DocumentByteArray);
        using var document = WordprocessingDocument.Open(stream, false);
        var endnotes = document.MainDocumentPart!.EndnotesPart!;
        var root = LoadXml(endnotes).Root!;
        XNamespace w = Wns;
        XNamespace r = Rns;
        var hyperlink = root.Descendants(w + "hyperlink").Single();
        var relationshipId = (string?)hyperlink.Attribute(r + "id");

        Assert.NotNull(relationshipId);
        Assert.NotEqual(HyperlinkRelationshipId, relationshipId);
        var relationship = endnotes.HyperlinkRelationships.Single(rel => rel.Id == relationshipId);
        Assert.Equal(RightTarget, relationship.Uri.ToString());
    }

    private static void AssertFootnoteImageRelationship(WmlDocument document)
    {
        using var stream = new MemoryStream(document.DocumentByteArray);
        using var wordDocument = WordprocessingDocument.Open(stream, false);
        var footnotes = wordDocument.MainDocumentPart!.FootnotesPart!;
        var root = LoadXml(footnotes).Root!;
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace r = Rns;
        var relationshipId = (string?)root.Descendants(a + "blip").Single().Attribute(r + "embed");

        Assert.NotNull(relationshipId);
        var imagePart = footnotes.GetPartById(relationshipId!);
        Assert.StartsWith("image/", imagePart.ContentType);
        Assert.Equal(OnePixelPng, ReadAllBytes(imagePart));
    }

    private static WmlDocument BuildFootnoteImageDocument(bool includeFootnote)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = AddRequiredMainParts(document);
            if (includeFootnote)
            {
                var footnotes = main.AddNewPart<FootnotesPart>();
                WriteXml(footnotes,
                    $"<w:footnotes xmlns:w=\"{Wns}\" xmlns:r=\"{Rns}\">" +
                    ReservedFootnoteXml() +
                    $"<w:footnote w:id=\"1\">{ImageParagraphXml(ImageRelationshipId)}</w:footnote>" +
                    "</w:footnotes>");
                var image = footnotes.AddNewPart<ImagePart>("image/png", ImageRelationshipId);
                using var imageStream = image.GetStream(FileMode.Create, FileAccess.Write);
                imageStream.Write(OnePixelPng, 0, OnePixelPng.Length);
            }

            var noteReference = includeFootnote
                ? "<w:r><w:footnoteReference w:id=\"1\"/></w:r>"
                : string.Empty;
            WriteXml(main,
                $"<w:document xmlns:w=\"{Wns}\"><w:body>" +
                $"<w:p><w:r><w:t>Body</w:t></w:r>{noteReference}</w:p>" +
                "</w:body></w:document>");
        }
        return new WmlDocument("note-image.docx", stream.ToArray());
    }

    private static WmlDocument BuildEndnoteHyperlinkDocument(bool includeRightNote)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = AddRequiredMainParts(document);
            var endnotes = main.AddNewPart<EndnotesPart>();
            endnotes.AddHyperlinkRelationship(
                new Uri(includeRightNote ? RightTarget : LeftTarget, UriKind.Absolute), true,
                HyperlinkRelationshipId);
            var insertedEndnote = includeRightNote
                ? $"<w:endnote w:id=\"1\"><w:p><w:hyperlink r:id=\"{HyperlinkRelationshipId}\">" +
                  "<w:r><w:t>Inserted right note hyperlink</w:t></w:r>" +
                  "</w:hyperlink></w:p></w:endnote>"
                : string.Empty;
            WriteXml(endnotes,
                $"<w:endnotes xmlns:w=\"{Wns}\" xmlns:r=\"{Rns}\">" +
                ReservedEndnoteXml() + insertedEndnote + "</w:endnotes>");

            var noteReference = includeRightNote
                ? "<w:r><w:endnoteReference w:id=\"1\"/></w:r>"
                : string.Empty;
            WriteXml(main,
                $"<w:document xmlns:w=\"{Wns}\"><w:body>" +
                $"<w:p><w:r><w:t>Body</w:t></w:r>{noteReference}</w:p>" +
                "</w:body></w:document>");
        }
        return new WmlDocument("note-hyperlink.docx", stream.ToArray());
    }

    private static MainDocumentPart AddRequiredMainParts(WordprocessingDocument document)
    {
        var main = document.AddMainDocumentPart();
        main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
        main.AddNewPart<DocumentSettingsPart>().Settings = new Settings();
        return main;
    }

    private static string ReservedFootnoteXml() =>
        "<w:footnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>" +
        "<w:footnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>";

    private static string ReservedEndnoteXml() =>
        "<w:endnote w:type=\"separator\" w:id=\"-1\"><w:p><w:r><w:separator/></w:r></w:p></w:endnote>" +
        "<w:endnote w:type=\"continuationSeparator\" w:id=\"0\"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>";

    private static string ImageParagraphXml(string relId) =>
        "<w:p><w:r><w:drawing>" +
        "<wp:inline xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
        "distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">" +
        "<wp:extent cx=\"95250\" cy=\"95250\"/><wp:docPr id=\"1\" name=\"Note image\"/>" +
        "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
        "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        "<pic:nvPicPr><pic:cNvPr id=\"1\" name=\"Note image\"/><pic:cNvPicPr/></pic:nvPicPr>" +
        $"<pic:blipFill><a:blip r:embed=\"{relId}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
        "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"95250\" cy=\"95250\"/></a:xfrm>" +
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>" +
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

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
