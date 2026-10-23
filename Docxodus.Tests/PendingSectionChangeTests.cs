using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus;
using Docxodus.Verification;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// Issue #1314: a pending section-property change (<c>w:sectPrChange</c>) keeps the section's
/// earlier properties as a nested <c>w:sectPr</c>. That copy is change history, not a section, so
/// no surface may count it as one or read its properties as the section's own. Each test pairs a
/// document whose sections carry pending changes with the same document without them.
/// </summary>
public class PendingSectionChangeTests
{
    private const string Ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static string Change(int id) =>
        $"<w:sectPrChange w:id=\"{id}\" w:author=\"Alice\" w:date=\"2025-01-01T00:00:00Z\"><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/>"
        + "<w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr></w:sectPrChange>";

    /// <summary>A landscape-Letter section then an A4 section; with <paramref name="pending"/>,
    /// each carries a pending change whose earlier properties were portrait Letter.</summary>
    private static byte[] TwoSections(bool pending) => Package(
        "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"15840\" w:h=\"12240\" w:orient=\"landscape\"/>"
        + (pending ? Change(5) : "") + "</w:sectPr></w:pPr><w:r><w:t>One.</w:t></w:r></w:p>"
        + "<w:p><w:r><w:t>Two.</w:t></w:r></w:p>"
        + "<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/>" + (pending ? Change(6) : "") + "</w:sectPr>");

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PackageManifest_CountsSectionsNotPendingChanges(bool pending)
    {
        Assert.Equal(2, PackageManifestGenerator.Generate(TwoSections(pending)).Facts.SectionCount);
    }

    [Fact]
    public void OpenContractsExport_PagesTakeTheSectionsCurrentSize()
    {
        (double, double)[] Pages(bool pending) => OpenContractExporter.Export(new WmlDocument("two.docx", TwoSections(pending)))
            .PawlsFileContent.Select(p => (System.Math.Round(p.Page.Width, 1), System.Math.Round(p.Page.Height, 1))).ToArray();

        // The page the two paragraphs land on is the A4 section's (11906 x 16838 twips), never the
        // portrait Letter (612 x 792pt) the pending changes stored.
        Assert.Contains((595.3, 841.9), Pages(pending: false));
        Assert.Equal(Pages(pending: false), Pages(pending: true));
    }

    [Fact]
    public void HtmlConverter_LaysOutTheLastSectionWithItsCurrentPage()
    {
        // A picture anchored 264pt from the page's left edge, 72pt wide. On the current A4 page
        // (body 451pt wide from 72pt) it sits right of centre and floats right; measured against the
        // earlier Letter page it would not. The render must not depend on the pending change.
        string Float(bool pending)
        {
            var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("float.docx", FloatingPicture(pending)),
                new WmlToHtmlConverterSettings
                {
                    RenderTrackedChanges = true,
                    FabricateCssClasses = false,
                    ImageHandler = _ => new XElement(XName.Get("img", "http://www.w3.org/1999/xhtml"), new XAttribute("src", "x")),
                });
            var img = html.Descendants().First(e => e.Name.LocalName == "img");
            return System.Text.RegularExpressions.Regex.Match(
                (string?)img.Parent!.Attribute("style") + " " + (string?)img.Attribute("style"), @"float:\s*\w+").Value;
        }

        Assert.NotEmpty(Float(pending: false));
        Assert.Equal(Float(pending: false), Float(pending: true));
    }

    private static byte[] FloatingPicture(bool pending)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var image = main.AddImagePart("image/png", "picture");
            using (var png = new MemoryStream(System.Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAC0lEQVR4nGNgQAYAAA4AAamRc7EAAAAASUVORK5CYII=")))
                image.FeedData(png);
            main.PutXDocument(XDocument.Parse($"""
                <w:document xmlns:w="{Ns}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <w:body><w:p><w:r><w:t>Lead.</w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:drawing>
                    <wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
                      <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>3352800</wp:posOffset></wp:positionH>
                      <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
                      <wp:extent cx="914400" cy="457200"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="p"/>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="black.png"/><pic:cNvPicPr/></pic:nvPicPr>
                          <pic:blipFill><a:blip r:embed="picture"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                          <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                        </pic:pic></a:graphicData></a:graphic>
                    </wp:anchor></w:drawing></w:r><w:r><w:t>Text.</w:t></w:r></w:p>
                    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>{(pending ? Change(6) : "")}</w:sectPr></w:body>
                </w:document>
                """));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings xmlns:w=\"{Ns}\"/>"));
        }
        return stream.ToArray();
    }

    private static byte[] Package(string bodyXml)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.PutXDocument(XDocument.Parse($"<w:document xmlns:w=\"{Ns}\"><w:body>{bodyXml}</w:body></w:document>"));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings xmlns:w=\"{Ns}\"/>"));
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(XDocument.Parse(
                $"<w:styles xmlns:w=\"{Ns}\"><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style></w:styles>"));
        }
        return stream.ToArray();
    }
}
