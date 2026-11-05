using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Docxodus.Ir;
using Docxodus.Tests.Ir;
using Xunit;
using static Docxodus.Tests.DocxBackendReconciliationTests;

namespace Docxodus.Tests;

public class DocxDiffAtomicSourceTests
{
    private static XElement Body(byte[] bytes)
    {
        using var package = WordprocessingDocument.Open(new MemoryStream(bytes), false);
        return package.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ShapeDefinitionAtParagraphBoundary_RemainsAvailableToItsShape(bool trailing)
    {
        const string definition = "<w:r><w:pict><v:shapetype id=\"_x0000_t1\"/></w:pict></w:r>";
        WmlDocument Document(string text) => IrTestDocuments.FromBodyXmlWithDrawingNamespaces(
            "<w:p>" + (trailing ? "" : definition) + "<w:r><w:tab/><w:t>" + text + "</w:t></w:r>" +
            (trailing ? definition : "") + "</w:p><w:p><w:r><w:pict><v:shape id=\"shape1\" type=\"#_x0000_t1\"/></w:pict></w:r></w:p>");
        var original = Document("OLD");
        var result = DocxCompare.Compare(original, Document("NEW")).DocumentByteArray;
        XNamespace v = "urn:schemas-microsoft-com:vml";
        foreach (var bytes in new[] { result, DocxDiffOps.AcceptRevisions(result), DocxDiffOps.RejectRevisions(result) })
        {
            var body = Body(bytes);
            Assert.Contains(body.Descendants(v + "shapetype"), e => (string?)e.Attribute("id") == "_x0000_t1");
            Assert.Contains(body.Descendants(v + "shape"), e => (string?)e.Attribute("type") == "#_x0000_t1");
        }
        NoNewValidationErrors(original.DocumentByteArray, result);
    }

    [Theory]
    [InlineData("field-code-text", false)]
    [InlineData("captured-link", false)]
    [InlineData("invalid-symbol", false)]
    [InlineData("field-code-text", true)]
    [InlineData("captured-link", true)]
    [InlineData("invalid-symbol", true)]
    public void ReaderNormalization_PreservesSourceContentInBothViews(string kind, bool consolidate)
    {
        string prefix = kind switch
        {
            "field-code-text" => "<w:r><w:fldChar w:fldCharType=\"begin\"/><w:t>ghost</w:t>" +
                "<w:instrText> PAGE </w:instrText><w:fldChar w:fldCharType=\"separate\"/><w:t>1</w:t>" +
                "<w:fldChar w:fldCharType=\"end\"/><w:tab/></w:r>",
            "captured-link" => "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
                "<w:hyperlink w:anchor=\"target\"><w:r><w:tab/></w:r></w:hyperlink>",
            _ => "<w:r><w:sym w:font=\"Symbol\" w:char=\"D800\"/><w:tab/></w:r>",
        };
        WmlDocument Document(string text) => IrTestDocuments.FromBodyXmlWithDrawingNamespaces("<w:p>" + prefix +
            (kind == "captured-link" ? "<w:r><w:t>" + text + "</w:t></w:r>" :
                "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>" + text + "</m:t></m:r></m:oMath><w:r><w:t> end</w:t></w:r>") + "</w:p>");
        var original = Document("OLD");
        var revised = Document("NEW");
        var result = (consolidate ? DocxDiff.Consolidate(original,
            new[] { new DocxDiffReviewer { Author = "Reviewer", Document = revised } }) :
            DocxCompare.Compare(original, revised)).DocumentByteArray;
        string[] Content(byte[] bytes) => Body(bytes).Descendants()
            .Where(e => e.Name == W.t || e.Name == M.t || e.Name == W.tab || e.Name == W.sym ||
                e.Name == W.fldChar || e.Name == W.instrText)
            .Select(e => e.Name.LocalName + ":" + (e.Name == W.fldChar ? (string?)e.Attribute(W.fldCharType) :
                e.Name == W.sym ? (string?)e.Attribute(W._char) : e.Value)).ToArray();
        Assert.Equal(Content(revised.DocumentByteArray), Content(DocxDiffOps.AcceptRevisions(result)));
        Assert.Equal(Content(original.DocumentByteArray), Content(DocxDiffOps.RejectRevisions(result)));
        NoNewValidationErrors(original.DocumentByteArray, result);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void SharedTextboxCarrier_CannotStraddleAnOppositeSideEdit(bool reverse, bool consolidate)
    {
        string Box(string text) => "<v:shape id=\"" + text + "\"><v:textbox><w:txbxContent><w:p><w:r><w:t>" +
            text + "</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape>";
        string Pict(string content) => "<w:r><w:pict>" + content + "</w:pict></w:r>";
        WmlDocument Doc(bool joined) => IrTestDocuments.FromBodyXmlWithDrawingNamespaces(
            "<w:p><w:r><w:t>before </w:t></w:r>" + (joined ? Pict(Box("first") + Box("second")) :
                Pict(Box("first")) + "<w:r><w:tab/></w:r>" + Pict(Box("second"))) +
            "<w:r><w:t> after</w:t></w:r></w:p>");
        var original = Doc(reverse);
        var revised = Doc(!reverse);
        var result = (consolidate ? DocxDiff.Consolidate(original,
            new[] { new DocxDiffReviewer { Author = "Reviewer", Document = revised } }) :
            DocxCompare.Compare(original, revised)).DocumentByteArray;
        string[] Order(byte[] bytes) => Body(bytes).Descendants().Where(e => e.Name == W.tab || e.Name == W.txbxContent)
            .Select(e => e.Name == W.tab ? "TAB" : e.Value).ToArray();
        Assert.Equal(Order(revised.DocumentByteArray), Order(DocxDiffOps.AcceptRevisions(result)));
        Assert.Equal(Order(original.DocumentByteArray), Order(DocxDiffOps.RejectRevisions(result)));
        NoNewValidationErrors(original.DocumentByteArray, result);
        var retained = IrReader.Read(Doc(true), new IrReaderOptions { RetainSources = true });
        var detached = IrReader.Read(Doc(true), new IrReaderOptions { RetainSources = false });
        var left = retained.AnchorIndex.Values.OfType<IrParagraph>().First();
        var right = detached.AnchorIndex.Values.OfType<IrParagraph>().First();
        Assert.Equal(left, right);
        Assert.Equal(left.Source.SharedCarriers, right.Source.SharedCarriers);
        Assert.Single(left.Source.SharedCarriers);
    }

    private static WmlDocument TextboxWithFill(bool image, string tail = "tail")
    {
        string fill = image
            ? "<a:blipFill><a:blip r:embed=\"rImg\"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>"
            : "<a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill>";
        string drawing = "<w:drawing xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">" +
            "<wp:inline><wp:extent cx=\"1828800\" cy=\"457200\"/><wp:docPr id=\"1\" name=\"shape\"/>" +
            "<a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\">" +
            "<wps:wsp><wps:cNvSpPr txBox=\"1\"/><wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/>" +
            "<a:ext cx=\"1828800\" cy=\"457200\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>" + fill +
            "</wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>unchanged textbox</w:t></w:r></w:p>" +
            "</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>";
        var png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a8bkAAAAASUVORK5CYII=");
        return IrTestDocuments.FromBodyXmlWithImageParts(
            "<w:p><w:r><w:t>before </w:t>" + drawing + "<w:t> after</w:t></w:r></w:p><w:p><w:r><w:t>" + tail + "</w:t></w:r></w:p>",
            ("rImg", png));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void ImageAndTextboxSharingOneCarrier_ResolveAsOneElement(bool removeImage, bool consolidate)
    {
        var original = TextboxWithFill(removeImage);
        var revised = TextboxWithFill(!removeImage);
        var result = (consolidate ? DocxDiff.Consolidate(original, new[]
        {
            new DocxDiffReviewer { Author = "Fill", Document = revised },
            new DocxDiffReviewer { Author = "Tail", Document = TextboxWithFill(removeImage, "updated tail") },
        }) : DocxCompare.Compare(original, revised)).DocumentByteArray;
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        foreach (bool accept in new[] { false, true })
        {
            var body = Body(accept ? DocxDiffOps.AcceptRevisions(result) : DocxDiffOps.RejectRevisions(result));
            var drawing = Assert.Single(body.Descendants(W.drawing));
            Assert.Equal(accept ? !removeImage : removeImage, drawing.Descendants(a + "blip").Any());
            Assert.Equal("unchanged textbox", Assert.Single(drawing.Descendants(W.txbxContent)).Value);
            Assert.Equal(accept && consolidate ? "updated tail" : "tail", body.Elements(W.p).Last().Value);
        }
        NoNewValidationErrors(original.DocumentByteArray, result);
    }
}
