#nullable enable

using System.Linq;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

/// <summary>
/// M1.2 Task 3 tests: note references (<c>w:footnoteReference</c>/<c>w:endnoteReference</c> →
/// <see cref="IrNoteRef"/>), inline images (<c>w:drawing</c> with an embedded <c>a:blip</c> →
/// <see cref="IrInlineImage"/>), and SDT/smartTag handling. Covers the content-hash semantics from
/// spec §6.1: note refs hash by kind sentinel only (no id), images by sentinel + image-bytes hash,
/// inline SDTs/smartTags splice transparently, and block SDTs retain their complete envelope.
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

    // --- M2.4b Workstream A: relationship-id-stable opaque hashing -------
    //
    // The headline guard. An OPAQUE element that carries a relationship id (here a VML w:pict /
    // v:imagedata, which stays opaque — no a:blip to promote) hashes by what the relationship POINTS AT,
    // not by the (freely renumbering) rel id. So: same rel id + different part bytes ⇒ DIFFERENT opaque
    // hash (content sensitivity preserved — the M1.2 guarantee); different rel id + same part bytes ⇒
    // SAME opaque hash (the new rel-numbering invariance — content identity over rel numbering).

    // A VML picture whose v:imagedata references an image part by rel id. No a:blip, so the reader keeps
    // the whole w:pict as an IrOpaqueInline — exercising the opaque canonical-hash rel-token path.
    private static string VmlPict(string relId) =>
        "<w:pict>" +
          "<v:shape xmlns:v=\"urn:schemas-microsoft-com:vml\">" +
            $"<v:imagedata r:id=\"{relId}\"/>" +
          "</v:shape>" +
        "</w:pict>";

    private static IrOpaqueInline OpaquePictInline(string relId, params (string, byte[])[] parts)
    {
        var doc = IrTestDocuments.FromBodyXmlWithImageParts(
            $"<w:p><w:r>{VmlPict(relId)}</w:r></w:p>", parts);
        var p = IrReader.Read(doc).Body.Blocks.OfType<IrParagraph>().Single();
        return p.Inlines.OfType<IrOpaqueInline>().Single();
    }

    [Fact]
    public void Read_OpaqueRel_DifferentRelId_SamePartBytes_SameOpaqueHash()
    {
        // Different relationship ids pointing at byte-identical parts ⇒ identical opaque hash. This is the
        // WC-1940 / SmartArt-rel-renumber closure expressed at the unit level.
        var a = OpaquePictInline("rIdA", ("rIdA", IrTestDocuments.TinyPng));
        var b = OpaquePictInline("rIdZZZ", ("rIdZZZ", IrTestDocuments.TinyPng));

        Assert.Equal(a.CanonicalHash, b.CanonicalHash);
    }

    [Fact]
    public void Read_OpaqueRel_SameRelId_DifferentPartBytes_DifferentOpaqueHash()
    {
        // Same relationship id but different part bytes ⇒ different opaque hash. Content sensitivity (the
        // M1.2 guarantee) must survive the rel-token canonicalization — a real content change still hashes
        // differently.
        var a = OpaquePictInline("rId1", ("rId1", new byte[] { 0x89, 0x50, 1, 2, 3 }));
        var b = OpaquePictInline("rId1", ("rId1", new byte[] { 0x89, 0x50, 9, 9, 9 }));

        Assert.NotEqual(a.CanonicalHash, b.CanonicalHash);
    }

    [Fact]
    public void Read_OpaqueRel_DanglingRel_ToleratesToStableToken()
    {
        // A rel id with no backing part (dangling) must not throw and must hash stably (totality). Two
        // dangling references to the same missing id hash equal.
        var a = OpaquePictInline("rIdGone");
        var b = OpaquePictInline("rIdGone");

        Assert.Equal(a.CanonicalHash, b.CanonicalHash);
    }

    // --- SDT / smartTag ---------------------------------------------------

    [Fact]
    public void Read_BlockSdt_PreservedAsEnvelope()
    {
        var doc = Read(
            "<w:sdt><w:sdtPr/><w:sdtContent>" +
              "<w:p><w:r><w:t>first</w:t></w:r></w:p>" +
              "<w:p><w:r><w:t>second</w:t></w:r></w:p>" +
            "</w:sdtContent></w:sdt>");

        var sdt = Assert.IsType<IrSdtBlock>(Assert.Single(doc.Body.Blocks));
        Assert.Equal(IrAnchorKind.Sdt, sdt.Anchor.Kind);
        Assert.Same(sdt, doc.FindByAnchor(sdt.Anchor));
        Assert.Equal("sdt", sdt.Source.Element?.Name.LocalName);

        var paras = sdt.Blocks.OfType<IrParagraph>().ToList();
        Assert.Equal(2, paras.Count);
        Assert.Empty(doc.Body.Blocks.OfType<IrOpaqueBlock>());

        // Both inner paragraphs are anchored and findable in the AnchorIndex.
        foreach (var p in paras)
            Assert.Same(p, doc.FindByAnchor(p.Anchor));
    }

    [Fact]
    public void Read_BlockSdt_WithTable_PreservesEnvelope()
    {
        var doc = Read(
            "<w:sdt><w:sdtContent>" +
              "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
            "</w:sdtContent></w:sdt>");

        var sdt = Assert.IsType<IrSdtBlock>(Assert.Single(doc.Body.Blocks));
        var table = Assert.IsType<IrTable>(Assert.Single(sdt.Blocks));
        Assert.Same(table, doc.FindByAnchor(table.Anchor));
        Assert.Empty(doc.Body.Blocks.OfType<IrOpaqueBlock>());
    }

    [Fact]
    public void Read_BlockSdt_EnvelopeMetadataParticipatesInIdentity()
    {
        var left = Assert.IsType<IrSdtBlock>(Assert.Single(Read(
            "<w:sdt><w:sdtPr><w:tag w:val=\"left\"/></w:sdtPr>" +
            "<w:sdtContent><w:p><w:r><w:t>same</w:t></w:r></w:p></w:sdtContent></w:sdt>")
            .Body.Blocks));
        var right = Assert.IsType<IrSdtBlock>(Assert.Single(Read(
            "<w:sdt><w:sdtPr><w:tag w:val=\"right\"/></w:sdtPr>" +
            "<w:sdtContent><w:p><w:r><w:t>same</w:t></w:r></w:p></w:sdtContent></w:sdt>")
            .Body.Blocks));

        // The direct child structure/text is identical, but wrapper metadata is part of the outer
        // canonical envelope and therefore must force an atomic SDT identity change.
        Assert.Equal(left.Blocks.Single().ContentHash, right.Blocks.Single().ContentHash);
        Assert.NotEqual(left.EnvelopeDigest, right.EnvelopeDigest);
        Assert.NotEqual(left.ContentHash, right.ContentHash);
    }

    [Fact]
    public void Read_BlockSdtWithoutContent_FallsBackToOpaque()
    {
        var doc = Read("<w:sdt><w:sdtPr><w:tag w:val=\"broken\"/></w:sdtPr></w:sdt>");

        var opaque = Assert.IsType<IrOpaqueBlock>(Assert.Single(doc.Body.Blocks));
        Assert.Equal("sdt", opaque.ElementName.LocalName);
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

    [Fact]
    public void Read_NestedBlockSdt_PreservesNestedEnvelopes()
    {
        // sdt-in-sdt, 2 deep: both wrappers stay addressable and the inner paragraph remains reachable.
        var doc = Read(
            "<w:sdt><w:sdtContent>" +
              "<w:sdt><w:sdtContent>" +
                "<w:p><w:r><w:t>inner</w:t></w:r></w:p>" +
              "</w:sdtContent></w:sdt>" +
            "</w:sdtContent></w:sdt>");

        var outer = Assert.IsType<IrSdtBlock>(Assert.Single(doc.Body.Blocks));
        var inner = Assert.IsType<IrSdtBlock>(Assert.Single(outer.Blocks));
        var para = Assert.IsType<IrParagraph>(Assert.Single(inner.Blocks));
        Assert.Empty(doc.Body.Blocks.OfType<IrOpaqueBlock>());
        Assert.Equal("inner", string.Concat(para.Inlines.OfType<IrTextRun>().Select(r => r.Text)));
        Assert.Same(outer, doc.FindByAnchor(outer.Anchor));
        Assert.Same(inner, doc.FindByAnchor(inner.Anchor));
        Assert.Same(para, doc.FindByAnchor(para.Anchor));
    }

    [Fact]
    public void Read_PathologicallyDeepSdt_FallsBackToOpaque()
    {
        // 70-deep nested block SDTs exceed the MaxSdtDepth (64) recursion cap: the reader must not
        // throw and must preserve the subtree as an opaque block rather than unwrapping forever.
        var sb = new System.Text.StringBuilder();
        const int depth = 70;
        for (int i = 0; i < depth; i++)
            sb.Append("<w:sdt><w:sdtContent>");
        sb.Append("<w:p><w:r><w:t>buried</w:t></w:r></w:p>");
        for (int i = 0; i < depth; i++)
            sb.Append("</w:sdtContent></w:sdt>");

        var doc = Read(sb.ToString());

        // No throw, and the cap produced an opaque fallback BELOW the preserved outer envelopes.
        Assert.Contains(FlattenBlocks(doc.Body.Blocks), b => b is IrOpaqueBlock);
    }

    [Fact]
    public void Read_BlockSdtInCell_PreservesEnvelopeAndCellText()
    {
        var doc = Read(
            "<w:tbl><w:tr><w:tc>" +
            "<w:sdt><w:sdtPr><w:tag w:val=\"cell-control\"/></w:sdtPr>" +
            "<w:sdtContent><w:p><w:r><w:t>inside cell</w:t></w:r></w:p></w:sdtContent></w:sdt>" +
            "</w:tc></w:tr></w:tbl>");

        var cell = Assert.Single(Assert.Single(doc.Body.Blocks.OfType<IrTable>()).Rows).Cells.Single();
        var sdt = Assert.IsType<IrSdtBlock>(Assert.Single(cell.Blocks));
        var para = Assert.IsType<IrParagraph>(Assert.Single(sdt.Blocks));
        Assert.Equal("inside cell", CellText(cell));
        Assert.Same(sdt, doc.FindByAnchor(sdt.Anchor));
        Assert.Same(para, doc.FindByAnchor(para.Anchor));
    }

    // --- row-level SDT cell unwrap (M1.4-T3 content-loss fix) -------------------------------------

    [Fact]
    public void Read_RowLevelSdt_WrappedCell_PresentInIrAndHashed()
    {
        // A w:sdt wrapping a w:tc as a row child: the oracle's table walk (Elements(tr).Elements(tc))
        // is blind to it, but the IR must NOT lose the cell — it appears in the row, flagged
        // FromRowSdt, and participates in the row/cell ContentHash.
        var doc = Read(
            "<w:tbl><w:tr>" +
            "<w:tc><w:p><w:r><w:t>direct</w:t></w:r></w:p></w:tc>" +
            "<w:sdt><w:sdtContent>" +
            "<w:tc><w:p><w:r><w:t>wrapped</w:t></w:r></w:p></w:tc>" +
            "</w:sdtContent></w:sdt>" +
            "</w:tr></w:tbl>");

        var table = Assert.Single(doc.Body.Blocks.OfType<IrTable>());
        var row = Assert.Single(table.Rows);
        Assert.Equal(2, row.Cells.Count);

        var direct = row.Cells[0];
        var wrapped = row.Cells[1];
        Assert.False(direct.FromRowSdt);
        Assert.True(wrapped.FromRowSdt);
        Assert.Equal("wrapped", CellText(wrapped));

        // The SDT-delivered cell is hashed: dropping it would change the row ContentHash, so a row
        // with only the direct cell must NOT hash equal to this two-cell row.
        var droppedDoc = Read(
            "<w:tbl><w:tr>" +
            "<w:tc><w:p><w:r><w:t>direct</w:t></w:r></w:p></w:tc>" +
            "</w:tr></w:tbl>");
        var droppedRow = Assert.Single(droppedDoc.Body.Blocks.OfType<IrTable>().Single().Rows);
        Assert.NotEqual(droppedRow.ContentHash, row.ContentHash);
    }

    [Fact]
    public void Read_RowLevelSdt_NestedWrapper_Unwrapped()
    {
        // sdt-in-sdt wrapping a w:tc at row level: still surfaced as a FromRowSdt cell.
        var doc = Read(
            "<w:tbl><w:tr>" +
            "<w:sdt><w:sdtContent><w:sdt><w:sdtContent>" +
            "<w:tc><w:p><w:r><w:t>deep</w:t></w:r></w:p></w:tc>" +
            "</w:sdtContent></w:sdt></w:sdtContent></w:sdt>" +
            "</w:tr></w:tbl>");

        var row = Assert.Single(doc.Body.Blocks.OfType<IrTable>().Single().Rows);
        var cell = Assert.Single(row.Cells);
        Assert.True(cell.FromRowSdt);
        Assert.Equal("deep", CellText(cell));
    }

    private static string CellText(IrCell cell) => string.Concat(cell.Blocks.Select(BlockText));

    private static string BlockText(IrBlock block) => block switch
    {
        IrParagraph p => string.Concat(p.Inlines.OfType<IrTextRun>().Select(r => r.Text)),
        IrSdtBlock sdt => string.Concat(sdt.Blocks.Select(BlockText)),
        IrTable table => string.Concat(table.Rows.SelectMany(r => r.Cells).Select(CellText)),
        _ => string.Empty,
    };

    private static IEnumerable<IrBlock> FlattenBlocks(IrNodeList<IrBlock> blocks)
    {
        foreach (var block in blocks)
        {
            yield return block;
            if (block is IrSdtBlock sdt)
                foreach (var child in FlattenBlocks(sdt.Blocks))
                    yield return child;
            else if (block is IrTable table)
                foreach (var cell in table.Rows.SelectMany(r => r.Cells))
                    foreach (var child in FlattenBlocks(cell.Blocks))
                        yield return child;
        }
    }
}
