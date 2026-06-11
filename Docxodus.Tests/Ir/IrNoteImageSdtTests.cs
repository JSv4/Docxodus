#nullable enable

using System.Linq;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

/// <summary>
/// M1.2 Task 3 tests: note references (<c>w:footnoteReference</c>/<c>w:endnoteReference</c> →
/// <see cref="IrNoteRef"/>), inline images (<c>w:drawing</c> with an embedded <c>a:blip</c> →
/// <see cref="IrInlineImage"/>), and N12 SDT/smartTag unwrapping. Covers the content-hash semantics
/// from spec §6.1: note refs hash by kind sentinel only (no id), images by sentinel + image-bytes
/// hash, and SDT/smartTag unwrap is content-transparent.
/// </summary>
public class IrNoteImageSdtTests
{
    private static IrDocument Read(string bodyXml) =>
        IrReader.Read(IrTestDocuments.FromBodyXml(bodyXml));

    private static IrParagraph Para(string bodyXml) =>
        Read(bodyXml).Body.Blocks.OfType<IrParagraph>().Single();

    // A drawing element wrapping an inline picture whose a:blip references the given embed rel id.
    private static string Drawing(string embedId, long cx = 100, long cy = 200,
        string? name = null, string? descr = null)
    {
        var docPrAttrs = $"id=\"1\" name=\"{name ?? "Picture 1"}\"" +
                         (descr is null ? "" : $" descr=\"{descr}\"");
        return
            "<w:drawing>" +
              "<wp:inline>" +
                $"<wp:extent cx=\"{cx}\" cy=\"{cy}\"/>" +
                $"<wp:docPr {docPrAttrs}/>" +
                "<a:graphic><a:graphicData>" +
                  "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                    $"<pic:blipFill><a:blip r:embed=\"{embedId}\"/></pic:blipFill>" +
                  "</pic:pic>" +
                "</a:graphicData></a:graphic>" +
              "</wp:inline>" +
            "</w:drawing>";
    }

    // --- note refs --------------------------------------------------------

    [Fact]
    public void Read_FootnoteRef_BecomesNoteRef()
    {
        var p = Para(
            "<w:p><w:r><w:footnoteReference w:id=\"3\"/></w:r></w:p>");

        var noteRef = Assert.Single(p.Inlines.OfType<IrNoteRef>());
        Assert.Equal(IrNoteKind.Footnote, noteRef.Kind);
        Assert.Equal("3", noteRef.NoteId);
    }

    [Fact]
    public void Read_EndnoteRef_BecomesNoteRef()
    {
        var p = Para(
            "<w:p><w:r><w:endnoteReference w:id=\"7\"/></w:r></w:p>");

        var noteRef = Assert.Single(p.Inlines.OfType<IrNoteRef>());
        Assert.Equal(IrNoteKind.Endnote, noteRef.Kind);
        Assert.Equal("7", noteRef.NoteId);
    }

    [Fact]
    public void Read_NoteRef_IdDoesNotAffectContentHash()
    {
        // Spec §6.1: note refs hash by kind sentinel ONLY — renumbering must not flip body hashes.
        var p2 = Para("<w:p><w:r><w:t>x</w:t><w:footnoteReference w:id=\"2\"/></w:r></w:p>");
        var p7 = Para("<w:p><w:r><w:t>x</w:t><w:footnoteReference w:id=\"7\"/></w:r></w:p>");

        Assert.Equal(p2.ContentHash, p7.ContentHash);
    }

    [Fact]
    public void Read_FootnoteVsEndnoteRef_DifferentContentHash()
    {
        // Distinct kind sentinels (0x05 vs 0x06) must keep footnote and endnote refs distinguishable.
        var fn = Para("<w:p><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>");
        var en = Para("<w:p><w:r><w:endnoteReference w:id=\"1\"/></w:r></w:p>");

        Assert.NotEqual(fn.ContentHash, en.ContentHash);
    }

    [Fact]
    public void Read_NoteRefMissingId_ToleratesEmptyString()
    {
        // separator/continuationSeparator ref variants carry no w:id — must not crash, id => "".
        var p = Para("<w:p><w:r><w:footnoteReference/></w:r></w:p>");

        var noteRef = Assert.Single(p.Inlines.OfType<IrNoteRef>());
        Assert.Equal("", noteRef.NoteId);
    }

    // --- inline images ----------------------------------------------------

    [Fact]
    public void Read_InlineImage_Promoted()
    {
        var doc = IrTestDocuments.FromBodyXmlWithImageParts(
            $"<w:p><w:r>{Drawing("rId99", cx: 12345, cy: 67890, name: "Logo", descr: "A logo")}</w:r></w:p>",
            ("rId99", IrTestDocuments.TinyPng));

        var p = IrReader.Read(doc).Body.Blocks.OfType<IrParagraph>().Single();
        var image = Assert.Single(p.Inlines.OfType<IrInlineImage>());

        Assert.Equal(12345, image.WidthEmu);
        Assert.Equal(67890, image.HeightEmu);
        Assert.Equal("A logo", image.AltText);
        Assert.Equal(IrHash.Compute(IrTestDocuments.TinyPng), image.ImageBytesHash);
    }

    [Fact]
    public void Read_InlineImage_AltTextFallsBackToName()
    {
        var doc = IrTestDocuments.FromBodyXmlWithImageParts(
            $"<w:p><w:r>{Drawing("rId1", name: "OnlyName")}</w:r></w:p>",
            ("rId1", IrTestDocuments.TinyPng));

        var image = IrReader.Read(doc).Body.Blocks.OfType<IrParagraph>().Single()
            .Inlines.OfType<IrInlineImage>().Single();

        Assert.Equal("OnlyName", image.AltText);
    }

    [Fact]
    public void Read_SameImageDifferentRelId_SameBytesHash()
    {
        // Image identity is the part bytes, not the relationship id (spec §12 q4).
        var docA = IrTestDocuments.FromBodyXmlWithImageParts(
            $"<w:p><w:r>{Drawing("rIdA")}</w:r></w:p>",
            ("rIdA", IrTestDocuments.TinyPng));
        var docB = IrTestDocuments.FromBodyXmlWithImageParts(
            $"<w:p><w:r>{Drawing("rIdZZZ")}</w:r></w:p>",
            ("rIdZZZ", IrTestDocuments.TinyPng));

        var imgA = IrReader.Read(docA).Body.Blocks.OfType<IrParagraph>().Single()
            .Inlines.OfType<IrInlineImage>().Single();
        var imgB = IrReader.Read(docB).Body.Blocks.OfType<IrParagraph>().Single()
            .Inlines.OfType<IrInlineImage>().Single();

        Assert.Equal(imgA.ImageBytesHash, imgB.ImageBytesHash);
    }

    [Fact]
    public void Read_ImageMissingRel_StaysOpaque()
    {
        // r:embed references a rel id with no backing image part: tolerate to opaque, never throw.
        var doc = IrTestDocuments.FromBodyXmlWithImageParts(
            $"<w:p><w:r>{Drawing("rIdMissing")}</w:r></w:p>");

        var p = IrReader.Read(doc).Body.Blocks.OfType<IrParagraph>().Single();
        Assert.Empty(p.Inlines.OfType<IrInlineImage>());
        var opaque = Assert.Single(p.Inlines.OfType<IrOpaqueInline>());
        Assert.Equal("drawing", opaque.ElementName.LocalName);
    }

    [Fact]
    public void Read_VmlPict_StaysOpaque()
    {
        // w:pict (VML) has no a:blip@embed — stays opaque.
        var p = Para("<w:p><w:r><w:pict><v:rect xmlns:v=\"urn:schemas-microsoft-com:vml\"/></w:pict></w:r></w:p>");

        Assert.Empty(p.Inlines.OfType<IrInlineImage>());
        var opaque = Assert.Single(p.Inlines.OfType<IrOpaqueInline>());
        Assert.Equal("pict", opaque.ElementName.LocalName);
    }

    [Fact]
    public void Read_InlineImage_ContentHashIncludesImageBytes()
    {
        // Two different image byte streams in otherwise-identical paragraphs → different ContentHash.
        var doc1 = IrTestDocuments.FromBodyXmlWithImageParts(
            $"<w:p><w:r>{Drawing("rId1")}</w:r></w:p>",
            ("rId1", new byte[] { 0x89, 0x50, 0x4E, 0x47, 1, 2, 3 }));
        var doc2 = IrTestDocuments.FromBodyXmlWithImageParts(
            $"<w:p><w:r>{Drawing("rId1")}</w:r></w:p>",
            ("rId1", new byte[] { 0x89, 0x50, 0x4E, 0x47, 9, 9, 9 }));

        var p1 = IrReader.Read(doc1).Body.Blocks.OfType<IrParagraph>().Single();
        var p2 = IrReader.Read(doc2).Body.Blocks.OfType<IrParagraph>().Single();

        Assert.NotEqual(p1.ContentHash, p2.ContentHash);
    }

    // --- N12: SDT / smartTag unwrap --------------------------------------

    [Fact]
    public void Read_BlockSdt_Unwrapped()
    {
        var doc = Read(
            "<w:sdt><w:sdtPr/><w:sdtContent>" +
              "<w:p><w:r><w:t>first</w:t></w:r></w:p>" +
              "<w:p><w:r><w:t>second</w:t></w:r></w:p>" +
            "</w:sdtContent></w:sdt>");

        var paras = doc.Body.Blocks.OfType<IrParagraph>().ToList();
        Assert.Equal(2, paras.Count);
        Assert.Empty(doc.Body.Blocks.OfType<IrOpaqueBlock>());

        // Both inner paragraphs are anchored and findable in the AnchorIndex.
        foreach (var p in paras)
            Assert.Same(p, doc.FindByAnchor(p.Anchor));
    }

    [Fact]
    public void Read_BlockSdt_WithTable_Unwrapped()
    {
        var doc = Read(
            "<w:sdt><w:sdtContent>" +
              "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
            "</w:sdtContent></w:sdt>");

        Assert.Single(doc.Body.Blocks.OfType<IrTable>());
        Assert.Empty(doc.Body.Blocks.OfType<IrOpaqueBlock>());
    }

    [Fact]
    public void Read_InlineSdt_Spliced()
    {
        // An inline w:sdt wrapping a run must read content-equal to the same run without the wrapper.
        var wrapped = Para(
            "<w:p><w:r><w:t>a</w:t></w:r>" +
            "<w:sdt><w:sdtContent><w:r><w:t>b</w:t></w:r></w:sdtContent></w:sdt></w:p>");
        var plain = Para("<w:p><w:r><w:t>a</w:t></w:r><w:r><w:t>b</w:t></w:r></w:p>");

        Assert.Equal(plain.ContentHash, wrapped.ContentHash);
        Assert.Equal("ab", string.Concat(wrapped.Inlines.OfType<IrTextRun>().Select(r => r.Text)));
    }

    [Fact]
    public void Read_SmartTag_Spliced()
    {
        var wrapped = Para(
            "<w:p><w:smartTag w:element=\"place\"><w:r><w:t>NYC</w:t></w:r></w:smartTag></w:p>");
        var plain = Para("<w:p><w:r><w:t>NYC</w:t></w:r></w:p>");

        Assert.Equal(plain.ContentHash, wrapped.ContentHash);
    }

    [Fact]
    public void Read_NestedSmartTag_Spliced()
    {
        var wrapped = Para(
            "<w:p><w:smartTag w:element=\"a\">" +
              "<w:smartTag w:element=\"b\"><w:r><w:t>deep</w:t></w:r></w:smartTag>" +
            "</w:smartTag></w:p>");
        var plain = Para("<w:p><w:r><w:t>deep</w:t></w:r></w:p>");

        Assert.Equal(plain.ContentHash, wrapped.ContentHash);
    }
}
