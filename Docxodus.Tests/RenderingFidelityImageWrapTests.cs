using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class RenderingFidelityImageWrapTests
{
    private static XElement Render(string wrap, long offset = 0, bool behind = false, int paginationMode = 1)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var image = main.AddImagePart("image/png", "picture");
            using (var bytes = new MemoryStream(Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAC0lEQVR4nGNgQAYAAA4AAamRc7EAAAAASUVORK5CYII=")))
                image.FeedData(bytes);
            main.PutXDocument(XDocument.Parse($"""
                <w:document xmlns:w="{W.w}" xmlns:r="{R.r}" xmlns:wp="{WP.wp}" xmlns:a="{A.a}" xmlns:pic="{Pic.pic}">
                  <w:body><w:p><w:r><w:drawing>
                    <wp:anchor distT="0" distB="0" distL="68580" distR="68580" simplePos="0" relativeHeight="0" behindDoc="{(behind ? 1 : 0)}" locked="0" layoutInCell="1" allowOverlap="1">
                      <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>{offset}</wp:posOffset></wp:positionH>
                      <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
                      <wp:extent cx="914400" cy="457200"/>{wrap}<wp:docPr id="1" name="Authored triangle"/>
                      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="black.png"/><pic:cNvPicPr/></pic:nvPicPr>
                          <pic:blipFill><a:blip r:embed="picture"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                          <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
                        </pic:pic>
                      </a:graphicData></a:graphic>
                    </wp:anchor>
                  </w:drawing></w:r><w:r><w:t>Authored wrapping test text.</w:t></w:r></w:p></w:body>
                </w:document>
                """));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings xmlns:w=\"{W.w}\"/>"));
        }
        return XElement.Parse(HtmlConversionOps.ConvertToHtml(stream.ToArray(),
            new HtmlConversionOptions { FabricateCssClasses = false, PaginationMode = paginationMode }));
    }

    [Theory]
    [InlineData("wrapNone", true)]
    [InlineData("wrapNone", false)]
    public void NonFlowingPicturesUseTheSharedDrawingAnchor(string wrap, bool behind)
    {
        var html = Render($"<wp:{wrap}/>", offset: 254000, behind: behind);
        var anchor = html.Descendants().Single(e => (string?)e.Attribute("data-docx-drawing-anchor") == "true");
        Assert.Contains("position: absolute", (string?)anchor.Attribute("style"));
        Assert.Equal(behind ? "true" : "false", (string?)anchor.Attribute("data-docx-anchor-behind"));
        Assert.Equal("20", (string?)anchor.Attribute("data-docx-anchor-h-offset"));
        Assert.Single(anchor.Descendants().Where(e => e.Name.LocalName == "img"));
        Assert.Contains("Authored wrapping test text.", html.Value);
    }

    [Fact]
    public void ContinuousPicturesRetainTheirInlineFallback()
    {
        var html = Render("<wp:wrapNone/>", behind: true, paginationMode: 0);
        Assert.DoesNotContain("data-docx-drawing-anchor", html.ToString());
        Assert.Single(html.Descendants().Where(e => e.Name.LocalName == "img"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExplicitWrappingStillExcludesTextWhenThePictureIsBehindIt(bool behind)
    {
        var html = Render("<wp:wrapSquare wrapText=\"right\"/>", behind: behind);
        var image = Assert.Single(html.Descendants().Where(e => e.Name.LocalName == "img"));
        Assert.Contains("float: left", (string?)image.Parent!.Attribute("style"));
        Assert.DoesNotContain("data-docx-drawing-anchor", html.ToString());
    }

    [Theory]
    [InlineData("wrapTight", "right")]
    [InlineData("wrapTight", "bothSides")]
    [InlineData("wrapThrough", "right")]
    [InlineData("wrapThrough", "bothSides")]
    public void ContourUsesDrawingCoordinatesAndWrapElementClearance(string kind, string wrapText)
    {
        var html = Render($"""
            <wp:{kind} wrapText="{wrapText}" distR="114300"><wp:wrapPolygon edited="1">
              <wp:start x="0" y="0"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="21600"/><wp:lineTo x="0" y="0"/>
            </wp:wrapPolygon></wp:{kind}>
            """);
        var style = html.Descendants().Single(e => e.Name.LocalName == "img").Parent!.Attribute("style")!.Value;
        Assert.Contains("shape-outside: polygon(0% 0%, 100% 0%, 0% 100%, 0% 0%) border-box", style);
        Assert.Contains("width: 72pt", style);
        Assert.Contains("height: 36pt", style);
        Assert.Contains("shape-margin: 9pt", style);
        Assert.Contains("margin-left: 0pt", style);
        Assert.Contains("margin-right: 9pt", style);
        Assert.Contains("--docx-wrap-distance-left: 5.4pt", style);
        Assert.Contains("--docx-wrap-distance-right: 9pt", style);
        Assert.Equal(wrapText, (string?)html.Descendants().Single(e => e.Name.LocalName == "img").Parent!.Attribute("data-docx-wrap-text"));
        Assert.Equal("true", (string?)html.Descendants().Single(e => e.Name.LocalName == "img").Parent!.Attribute("data-docx-edge-picture"));
        Assert.Contains("Authored wrapping test text.", html.Value);
    }

    [Fact]
    public void SquareWrapKeepsThePictureAtItsColumnAnchor()
    {
        var html = Render("<wp:wrapSquare wrapText=\"bothSides\"/>");
        var style = html.Descendants().Single(e => e.Name.LocalName == "img").Parent!.Attribute("style")!.Value;
        Assert.Contains("margin-left: 0pt", style);
        Assert.Contains("margin-right: 5.4pt", style);
        Assert.DoesNotContain("shape-outside", style);
        Assert.DoesNotContain("data-docx-wrap-text", html.ToString());
        Assert.Equal("true", (string?)html.Descendants().Single(e => e.Name.LocalName == "img").Parent!.Attribute("data-docx-edge-picture"));
    }

    [Fact]
    public void IncompleteContourKeepsTheRectangleFallback()
    {
        var html = Render("""
            <wp:wrapTight wrapText="right"><wp:wrapPolygon>
              <wp:start x="0" y="0"/><wp:lineTo x="21600"/><wp:lineTo x="0" y="21600"/>
            </wp:wrapPolygon></wp:wrapTight>
            """);
        Assert.DoesNotContain("shape-outside", html.ToString());
        Assert.DoesNotContain("data-docx-wrap-text", html.ToString());
        Assert.Contains("Authored wrapping test text.", html.Value);
    }

    [Fact]
    public void OffsetPictureKeepsItsExistingFloatFallback()
    {
        var html = Render("<wp:wrapSquare wrapText=\"bothSides\"/>", 127000);
        Assert.DoesNotContain("data-docx-edge-picture", html.ToString());
        Assert.Contains("Authored wrapping test text.", html.Value);
    }
}
