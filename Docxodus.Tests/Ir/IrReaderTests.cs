#nullable enable

using System;
using System.Linq;
using Docxodus;
using Docxodus.Ir;
using Xunit;

namespace Docxodus.Tests.Ir;

public class IrReaderTests
{
    private static string Text(IrParagraph p) =>
        string.Concat(p.Inlines.OfType<IrTextRun>().Select(r => r.Text));

    [Fact]
    public void Read_SimpleParagraphs_ProducesParagraphBlocks()
    {
        var doc = IrTestDocuments.Create("Hello world", "Second line");
        var ir = IrReader.Read(doc);

        var paras = ir.Body.Blocks.OfType<IrParagraph>().ToList();
        Assert.Equal(2, paras.Count);
        Assert.Equal("Hello world", Text(paras[0]));
        Assert.Equal("Second line", Text(paras[1]));

        foreach (var p in paras)
        {
            Assert.Equal(IrAnchorKind.P, p.Anchor.Kind);
            Assert.Equal("body", p.Anchor.Scope);
            Assert.Equal(32, p.Anchor.Unid.Length);
            Assert.Matches("^[0-9a-f]{32}$", p.Anchor.Unid);
        }
    }

    [Fact]
    public void Read_DoesNotMutateInput()
    {
        var doc = IrTestDocuments.Create("Alpha", "Beta");
        var before = (byte[])doc.DocumentByteArray.Clone();

        IrReader.Read(doc);

        Assert.Equal(before, doc.DocumentByteArray);
    }

    [Fact]
    public void Read_Twice_IdenticalAnchorsAndHashes()
    {
        var doc = IrTestDocuments.Create("Same bytes", "Twice over");
        var bytes = (byte[])doc.DocumentByteArray.Clone();

        var ir1 = IrReader.Read(new WmlDocument("a.docx", (byte[])bytes.Clone()));
        var ir2 = IrReader.Read(new WmlDocument("a.docx", (byte[])bytes.Clone()));

        var b1 = ir1.Body.Blocks.ToList();
        var b2 = ir2.Body.Blocks.ToList();
        Assert.Equal(b1.Count, b2.Count);
        for (int i = 0; i < b1.Count; i++)
        {
            Assert.Equal(b1[i].Anchor.ToString(), b2[i].Anchor.ToString());
            Assert.Equal(b1[i].ContentHash.ToHex(), b2[i].ContentHash.ToHex());
            Assert.Equal(b1[i].FormatFingerprint.ToHex(), b2[i].FormatFingerprint.ToHex());
        }

        Assert.Equal(ir1.Body, ir2.Body);
    }

    [Fact]
    public void Read_BoldRun_MapsRunFormat()
    {
        var doc = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r></w:p>");
        var ir = IrReader.Read(doc);

        var run = ir.Body.Blocks.OfType<IrParagraph>().Single()
            .Inlines.OfType<IrTextRun>().Single();
        Assert.True(run.Format.Bold);
        Assert.Equal("bold", run.Text);
    }

    /// <summary>
    /// M2.4b Workstream D — a bookmark marker (bookmarkStart/bookmarkEnd) that appears as a DIRECT body child
    /// (legal OOXML, a sibling of w:p) is DROPPED by the block-level N3 rule, exactly as the inline walker drops
    /// a bookmark inside a paragraph. This mirrors WmlComparer's PreProcessMarkup (RemoveBookmarks=true) so a
    /// body-level bookmark never becomes a spurious content block (the WC022 stray bookmarkEnd / WC-BodyBookmarks
    /// section bookmarks). The real paragraphs around it are unaffected.
    /// </summary>
    [Fact]
    public void Read_BodyLevelBookmarkMarkers_AreDropped()
    {
        var doc = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>before</w:t></w:r></w:p>" +
            "<w:bookmarkStart w:id=\"7\" w:name=\"sec\"/>" +
            "<w:bookmarkEnd w:id=\"7\"/>" +
            "<w:p><w:r><w:t>after</w:t></w:r></w:p>");
        var ir = IrReader.Read(doc);

        // Only the two real paragraphs survive; the body-level bookmark markers are gone (no opaque block).
        var paras = ir.Body.Blocks.OfType<IrParagraph>().ToList();
        Assert.Equal(2, paras.Count);
        Assert.DoesNotContain(ir.Body.Blocks, b => b is IrOpaqueBlock);
        Assert.Equal("before", string.Concat(paras[0].Inlines.OfType<IrTextRun>().Select(r => r.Text)));
        Assert.Equal("after", string.Concat(paras[1].Inlines.OfType<IrTextRun>().Select(r => r.Text)));
    }

    [Fact]
    public void Read_AdjacentEqualRuns_Coalesce()
    {
        var doc = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t xml:space=\"preserve\">Hello </w:t></w:r>" +
            "<w:r><w:t>world</w:t></w:r></w:p>");
        var ir = IrReader.Read(doc);

        var runs = ir.Body.Blocks.OfType<IrParagraph>().Single()
            .Inlines.OfType<IrTextRun>().ToList();
        Assert.Single(runs);
        Assert.Equal("Hello world", runs[0].Text);
    }

    [Fact]
    public void Read_TabAndBreak_BecomeTypedInlines()
    {
        var doc = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t>" +
            "<w:br w:type=\"page\"/></w:r></w:p>");
        var ir = IrReader.Read(doc);

        var inlines = ir.Body.Blocks.OfType<IrParagraph>().Single().Inlines.ToList();
        Assert.Contains(inlines, i => i is IrTab);
        var brk = Assert.IsType<IrBreak>(inlines.Single(i => i is IrBreak));
        Assert.Equal(IrBreakKind.Page, brk.Kind);
    }

    [Fact]
    public void Read_Table_StructureAndAnchors()
    {
        var doc = IrTestDocuments.FromBodyXml(
            "<w:tbl>" +
            "<w:tblPr/><w:tblGrid><w:gridCol w:w=\"100\"/><w:gridCol w:w=\"100\"/></w:tblGrid>" +
            "<w:tr><w:tc><w:p><w:r><w:t>R0C0</w:t></w:r></w:p></w:tc>" +
            "<w:tc><w:p><w:r><w:t>R0C1</w:t></w:r></w:p></w:tc></w:tr>" +
            "<w:tr><w:tc><w:p><w:r><w:t>R1C0</w:t></w:r></w:p></w:tc>" +
            "<w:tc><w:p><w:r><w:t>R1C1</w:t></w:r></w:p></w:tc></w:tr>" +
            "</w:tbl>");
        var ir = IrReader.Read(doc);

        var table = Assert.IsType<IrTable>(ir.Body.Blocks.Single());
        Assert.Equal(IrAnchorKind.Tbl, table.Anchor.Kind);
        Assert.Equal(2, table.Rows.Count);
        foreach (var row in table.Rows)
        {
            Assert.Equal(IrAnchorKind.Tr, row.Anchor.Kind);
            Assert.Equal(2, row.Cells.Count);
            foreach (var cell in row.Cells)
            {
                Assert.Equal(IrAnchorKind.Tc, cell.Anchor.Kind);
                var para = Assert.IsType<IrParagraph>(cell.Blocks.Single());
                Assert.NotNull(ir.FindByAnchor(para.Anchor));
            }
        }
    }

    [Fact]
    public void Read_NestedTable_Recurses()
    {
        var doc = IrTestDocuments.FromBodyXml(
            "<w:tbl><w:tr><w:tc>" +
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>" +
            "</w:tc></w:tr></w:tbl>");
        var ir = IrReader.Read(doc);

        var outer = Assert.IsType<IrTable>(ir.Body.Blocks.Single());
        var cell = outer.Rows.Single().Cells.Single();
        var inner = Assert.IsType<IrTable>(cell.Blocks.Single(b => b is IrTable));
        Assert.NotNull(ir.FindByAnchor(inner.Anchor));
    }

    [Fact]
    public void Read_UnknownElement_BecomesOpaque()
    {
        // w:customXmlInsRangeStart is an unmodeled block (unlike w:sdt, which N12 now unwraps in
        // M1.2 Task 3); w:ptab (an absolute-position tab) is an unmodeled inline. Read with
        // FailIfPresent so the default Accept round-trip (which normalizes such synthetic unknown
        // blocks away before the IR ever sees them) is skipped — this test is about the reader's
        // opaque mapping, not RevisionProcessor's body normalization.
        var doc = IrTestDocuments.FromBodyXml(
            "<w:customXmlInsRangeStart w:id=\"1\"/>" +
            "<w:p><w:r><w:ptab w:relativeTo=\"margin\" w:alignment=\"left\" w:leader=\"none\"/></w:r></w:p>");
        var ir = IrReader.Read(doc, new IrReaderOptions { RevisionView = RevisionView.FailIfPresent });

        Assert.Contains(ir.Body.Blocks, b => b is IrOpaqueBlock);
        var para = ir.Body.Blocks.OfType<IrParagraph>().Single();
        Assert.Contains(para.Inlines, i => i is IrOpaqueInline);
    }

    [Fact]
    public void Read_ContentHash_IgnoresFormatting()
    {
        var plain = IrReader.Read(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>hello</w:t></w:r></w:p>"));
        var bold = IrReader.Read(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>hello</w:t></w:r></w:p>"));

        var p1 = plain.Body.Blocks.OfType<IrParagraph>().Single();
        var p2 = bold.Body.Blocks.OfType<IrParagraph>().Single();

        Assert.Equal(p1.ContentHash.ToHex(), p2.ContentHash.ToHex());
        Assert.NotEqual(p1.FormatFingerprint.ToHex(), p2.FormatFingerprint.ToHex());
    }

    [Fact]
    public void Read_RevisionView_AcceptVsReject()
    {
        const string body =
            "<w:p><w:r><w:t xml:space=\"preserve\">kept </w:t></w:r>" +
            "<w:ins w:id=\"1\" w:author=\"a\"><w:r><w:t>inserted</w:t></w:r></w:ins></w:p>";

        var accepted = IrReader.Read(IrTestDocuments.FromBodyXml(body),
            new IrReaderOptions { RevisionView = RevisionView.Accept });
        var rejected = IrReader.Read(IrTestDocuments.FromBodyXml(body),
            new IrReaderOptions { RevisionView = RevisionView.Reject });

        var acceptedText = Text(accepted.Body.Blocks.OfType<IrParagraph>().Single());
        var rejectedText = Text(rejected.Body.Blocks.OfType<IrParagraph>().Single());
        Assert.Contains("inserted", acceptedText);
        Assert.DoesNotContain("inserted", rejectedText);

        Assert.Throws<DocxodusException>(() =>
            IrReader.Read(IrTestDocuments.FromBodyXml(body),
                new IrReaderOptions { RevisionView = RevisionView.FailIfPresent }));
    }

    [Fact]
    public void Read_TrailingSectPr_BecomesSectionBreak()
    {
        var doc = IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>body</w:t></w:r></w:p>" +
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>");
        var ir = IrReader.Read(doc);

        var sec = Assert.IsType<IrSectionBreak>(ir.Body.Blocks.Last());
        Assert.Equal(IrAnchorKind.Sec, sec.Anchor.Kind);
        Assert.Equal(12240, sec.Format.PageWidthTwips);
    }

    [Fact]
    public void Read_StyleInheritedListItem_ClassifiedAsLi()
    {
        // The paragraph carries NO inline w:numPr; it is a list item only because its pStyle
        // ("MyListPara") is basedOn "ListBase", whose pPr carries w:numPr. KindFor → IsListItem
        // must walk the styles part (reachable via the part annotation IrReader stashes) to see it.
        const string styles =
            "<w:style w:type=\"paragraph\" w:styleId=\"ListBase\">" +
            "<w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr></w:style>" +
            "<w:style w:type=\"paragraph\" w:styleId=\"MyListPara\">" +
            "<w:basedOn w:val=\"ListBase\"/></w:style>";
        const string body =
            "<w:p><w:pPr><w:pStyle w:val=\"MyListPara\"/></w:pPr>" +
            "<w:r><w:t>item</w:t></w:r></w:p>";

        var doc = IrTestDocuments.FromBodyAndStylesXml(body, styles);
        var ir = IrReader.Read(doc);

        var para = ir.Body.Blocks.OfType<IrParagraph>().Single();
        Assert.Equal(IrAnchorKind.Li, para.Anchor.Kind);

        // Determinism: reading the same bytes again yields identical anchors (and kind).
        var ir2 = IrReader.Read(doc);
        var para2 = ir2.Body.Blocks.OfType<IrParagraph>().Single();
        Assert.Equal(para.Anchor.ToString(), para2.Anchor.ToString());
        Assert.Equal(IrAnchorKind.Li, para2.Anchor.Kind);
    }

    [Fact]
    public void Read_UnmodeledFormatting_FlipsFingerprintOnly()
    {
        // rPr case: w:rFonts w:hAnsi is unmodeled (only w:ascii is modeled), so text is identical
        // (ContentHash equal) but the unmodeled digest — hence FormatFingerprint — differs.
        var rPlain = IrReader.Read(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>same</w:t></w:r></w:p>"));
        var rUnmodeled = IrReader.Read(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:rPr><w:rFonts w:hAnsi=\"Arial\"/></w:rPr><w:t>same</w:t></w:r></w:p>"));

        var rp1 = rPlain.Body.Blocks.OfType<IrParagraph>().Single();
        var rp2 = rUnmodeled.Body.Blocks.OfType<IrParagraph>().Single();
        Assert.Equal(rp1.ContentHash.ToHex(), rp2.ContentHash.ToHex());
        Assert.NotEqual(rp1.FormatFingerprint.ToHex(), rp2.FormatFingerprint.ToHex());

        // pPr case: w:kinsoku is an unmodeled paragraph property. Same shape: content equal,
        // fingerprint differs.
        var pPlain = IrReader.Read(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>same</w:t></w:r></w:p>"));
        var pUnmodeled = IrReader.Read(IrTestDocuments.FromBodyXml(
            "<w:p><w:pPr><w:kinsoku/></w:pPr><w:r><w:t>same</w:t></w:r></w:p>"));

        var pp1 = pPlain.Body.Blocks.OfType<IrParagraph>().Single();
        var pp2 = pUnmodeled.Body.Blocks.OfType<IrParagraph>().Single();
        Assert.Equal(pp1.ContentHash.ToHex(), pp2.ContentHash.ToHex());
        Assert.NotEqual(pp1.FormatFingerprint.ToHex(), pp2.FormatFingerprint.ToHex());
    }

    [Fact]
    public void Read_ProofErr_DoesNotAffectHashes()
    {
        var without = IrReader.Read(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>spell</w:t></w:r></w:p>"));
        var with = IrReader.Read(IrTestDocuments.FromBodyXml(
            "<w:p><w:proofErr w:type=\"spellStart\"/><w:r><w:t>spell</w:t></w:r>" +
            "<w:proofErr w:type=\"spellEnd\"/></w:p>"));

        var p1 = without.Body.Blocks.OfType<IrParagraph>().Single();
        var p2 = with.Body.Blocks.OfType<IrParagraph>().Single();
        Assert.Equal(p1.ContentHash.ToHex(), p2.ContentHash.ToHex());
        Assert.Equal(p1.FormatFingerprint.ToHex(), p2.FormatFingerprint.ToHex());
    }

    // --- textbox bodies (M1.5) -------------------------------------------------

    private static IrTextbox SingleTextbox(IrParagraph p) =>
        p.Inlines.OfType<IrTextbox>().Single();

    [Fact]
    public void Read_TextboxInDrawing_InnerParagraphModeledAndAnchored()
    {
        // A w:drawing carrying a wps:txbx → w:txbxContent body with one inner paragraph. No blip, so
        // the drawing is NOT promoted to an image — its textbox body is modeled instead.
        var doc = IrTestDocuments.FromBodyXmlWithDrawingNamespaces(
            "<w:p><w:r><w:drawing><wp:inline><a:graphic><a:graphicData>" +
            "<wps:wsp><wps:txbx><w:txbxContent>" +
            "<w:p><w:r><w:t>Inside</w:t></w:r></w:p>" +
            "</w:txbxContent></wps:txbx></wps:wsp>" +
            "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>");
        var ir = IrReader.Read(doc);

        var outer = ir.Body.Blocks.OfType<IrParagraph>().Single();
        var textbox = SingleTextbox(outer);
        var inner = Assert.IsType<IrParagraph>(textbox.Blocks.Single());

        Assert.Equal("Inside", string.Concat(inner.Inlines.OfType<IrTextRun>().Select(r => r.Text)));
        Assert.Equal(IrAnchorKind.P, inner.Anchor.Kind);
        Assert.Equal("body", inner.Anchor.Scope);
        Assert.Equal(32, inner.Anchor.Unid.Length);

        // The inner paragraph is registered in the document AnchorIndex (oracle DescendantsAndSelf parity).
        Assert.Same(inner, ir.FindByAnchor(inner.Anchor));
    }

    [Fact]
    public void Read_TextboxInPict_VmlVariantModeled()
    {
        // The VML w:pict/v:textbox variant must be modeled identically (the reader keys on the
        // w:txbxContent descendant, regardless of the DrawingML-vs-VML wrapper).
        var doc = IrTestDocuments.FromBodyXmlWithDrawingNamespaces(
            "<w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent>" +
            "<w:p><w:r><w:t>VmlInside</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>");
        var ir = IrReader.Read(doc);

        var outer = ir.Body.Blocks.OfType<IrParagraph>().Single();
        var inner = Assert.IsType<IrParagraph>(SingleTextbox(outer).Blocks.Single());
        Assert.Equal("VmlInside", string.Concat(inner.Inlines.OfType<IrTextRun>().Select(r => r.Text)));
        Assert.Same(inner, ir.FindByAnchor(inner.Anchor));
    }

    [Fact]
    public void Read_AlternateContent_BothChoiceAndFallbackTextboxesModeled()
    {
        // Word emits the same logical textbox twice: a DrawingML mc:Choice (wps:txbx) and a VML
        // mc:Fallback (v:textbox). The reader models BOTH (mirroring the oracle's both-copies walk),
        // so two IrTextbox nodes appear in document order.
        var doc = IrTestDocuments.FromBodyXmlWithDrawingNamespaces(
            "<w:p><w:r><mc:AlternateContent>" +
            "<mc:Choice Requires=\"wps\"><w:drawing><wp:inline><a:graphic><a:graphicData>" +
            "<wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>In</w:t></w:r></w:p>" +
            "</w:txbxContent></wps:txbx></wps:wsp>" +
            "</a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice>" +
            "<mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent>" +
            "<w:p><w:r><w:t>In</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>" +
            "</mc:AlternateContent></w:r><w:r><w:t>Out</w:t></w:r></w:p>");
        var ir = IrReader.Read(doc);

        var outer = ir.Body.Blocks.OfType<IrParagraph>().Single();
        var textboxes = outer.Inlines.OfType<IrTextbox>().ToList();
        Assert.Equal(2, textboxes.Count);
        foreach (var tb in textboxes)
        {
            var inner = Assert.IsType<IrParagraph>(tb.Blocks.Single());
            Assert.Equal("In", string.Concat(inner.Inlines.OfType<IrTextRun>().Select(r => r.Text)));
        }
        // Both inner paragraphs have DISTINCT anchors (different XML subtrees → different Unids) and
        // both are indexed — no collision.
        var innerAnchors = textboxes
            .Select(tb => ((IrParagraph)tb.Blocks.Single()).Anchor.ToString()).ToList();
        Assert.Equal(2, innerAnchors.Distinct().Count());
        foreach (var a in innerAnchors)
            Assert.True(ir.AnchorIndex.ContainsKey(a));
    }

    [Fact]
    public void Read_TextboxWithTable_InnerTableModeledAndIndexed()
    {
        var doc = IrTestDocuments.FromBodyXmlWithDrawingNamespaces(
            "<w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent>" +
            "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w=\"100\"/></w:tblGrid>" +
            "<w:tr><w:tc><w:p><w:r><w:t>CellText</w:t></w:r></w:p></w:tc></w:tr>" +
            "</w:tbl>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>");
        var ir = IrReader.Read(doc);

        var outer = ir.Body.Blocks.OfType<IrParagraph>().Single();
        var innerTable = Assert.IsType<IrTable>(SingleTextbox(outer).Blocks.Single());
        var cellPara = Assert.IsType<IrParagraph>(innerTable.Rows.Single().Cells.Single().Blocks.Single());
        Assert.Equal("CellText", string.Concat(cellPara.Inlines.OfType<IrTextRun>().Select(r => r.Text)));
        // The inner table, its cell paragraph, and the table's own anchor are all indexed.
        Assert.Same(innerTable, ir.FindByAnchor(innerTable.Anchor));
        Assert.Same(cellPara, ir.FindByAnchor(cellPara.Anchor));
    }

    [Fact]
    public void Read_DrawingWithBlipAndTextbox_YieldsBothImageAndTextbox()
    {
        // A w:drawing carrying BOTH a resolvable a:blip image AND a wps:txbx textbox: image promotion
        // and textbox modeling are INDEPENDENT (oracle parity — both the blip and the txbxContent text
        // are seen). The inline list carries an IrInlineImage AND an IrTextbox.
        var doc = IrTestDocuments.FromBodyXmlWithImageParts(
            "<w:p><w:r><w:drawing><wp:inline>" +
            "<wp:extent cx=\"100\" cy=\"100\"/>" +
            "<a:graphic><a:graphicData>" +
            "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
            "<pic:blipFill><a:blip r:embed=\"rIdImg\"/></pic:blipFill></pic:pic>" +
            "<wps:wsp xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">" +
            "<wps:txbx><w:txbxContent><w:p><w:r><w:t>InBox</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp>" +
            "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>",
            ("rIdImg", IrTestDocuments.TinyPng));
        var ir = IrReader.Read(doc);

        var outer = ir.Body.Blocks.OfType<IrParagraph>().Single();
        Assert.Single(outer.Inlines.OfType<IrInlineImage>());
        var inner = Assert.IsType<IrParagraph>(SingleTextbox(outer).Blocks.Single());
        Assert.Equal("InBox", string.Concat(inner.Inlines.OfType<IrTextRun>().Select(r => r.Text)));
    }

    [Fact]
    public void Read_NestedTextbox_RespectsDepthCapWithoutThrowing()
    {
        // Build a textbox nested far beyond the textbox depth cap (16): a chain of w:pict/v:textbox/
        // w:txbxContent wrapping a w:p with text at the bottom. The reader must not throw (totality);
        // beyond the cap the deepest body is preserved opaquely rather than recursing further.
        const int depth = 40;
        var inner = "<w:p><w:r><w:t>Bottom</w:t></w:r></w:p>";
        var sb = new System.Text.StringBuilder();
        for (int i = 0; i < depth; i++)
            sb.Append("<w:pict><v:shape><v:textbox><w:txbxContent>");
        sb.Append(inner);
        for (int i = 0; i < depth; i++)
            sb.Append("</w:txbxContent></v:textbox></v:shape></w:pict>");
        var body = $"<w:p><w:r>{sb}</w:r></w:p>";

        var ex = Record.Exception(() =>
            IrReader.Read(IrTestDocuments.FromBodyXmlWithDrawingNamespaces(body)));
        Assert.Null(ex);

        // The outermost paragraph still carries exactly one textbox, and the nesting bottoms out in an
        // opaque inline (no unbounded recursion / stack overflow).
        var ir = IrReader.Read(IrTestDocuments.FromBodyXmlWithDrawingNamespaces(body));
        var outer = ir.Body.Blocks.OfType<IrParagraph>().Single();
        Assert.Single(outer.Inlines.OfType<IrTextbox>());
    }

    [Fact]
    public void Read_TextboxTextChange_FlipsContainingParagraphContentHash()
    {
        var make = new System.Func<string, IrParagraph>(text =>
            IrReader.Read(IrTestDocuments.FromBodyXmlWithDrawingNamespaces(
                "<w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent>" +
                $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>" +
                "</w:txbxContent></v:textbox></v:shape></w:pict></w:r>" +
                "<w:r><w:t>Outside</w:t></w:r></w:p>"))
                .Body.Blocks.OfType<IrParagraph>().Single());

        var a = make("Alpha");
        var b = make("Beta");
        // Textbox-only text edit flips the CONTAINING paragraph's ContentHash (blind spot closed).
        Assert.NotEqual(a.ContentHash.ToHex(), b.ContentHash.ToHex());
    }

    [Fact]
    public void Read_TextboxText_DistinctFromInlineTextOfSameContent()
    {
        // Sentinel framing keeps textbox text distinct from identical inline (non-textbox) text:
        // "X" in a textbox must not content-hash-equal "X" as a plain run.
        var textbox = IrReader.Read(IrTestDocuments.FromBodyXmlWithDrawingNamespaces(
            "<w:p><w:r><w:pict><v:shape><v:textbox><w:txbxContent>" +
            "<w:p><w:r><w:t>X</w:t></w:r></w:p>" +
            "</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>"))
            .Body.Blocks.OfType<IrParagraph>().First();
        var plain = IrReader.Read(IrTestDocuments.FromBodyXml(
            "<w:p><w:r><w:t>X</w:t></w:r></w:p>"))
            .Body.Blocks.OfType<IrParagraph>().Single();

        Assert.NotEqual(textbox.ContentHash.ToHex(), plain.ContentHash.ToHex());
    }
}
