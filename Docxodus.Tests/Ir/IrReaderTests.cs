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
        var doc = IrTestDocuments.FromBodyXml(
            "<w:sdt><w:sdtContent><w:p><w:r><w:t>x</w:t></w:r></w:p></w:sdtContent></w:sdt>" +
            "<w:p><w:hyperlink r:id=\"rId9\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
            "<w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>");
        var ir = IrReader.Read(doc);

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
}
