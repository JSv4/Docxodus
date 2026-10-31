using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class HtmlGroupedVmlTextBoxTests
{
    private const string Placement = "position:absolute;z-index:1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page;";

    private static string Box(string id, string text, string style) => $"""
        <v:shape id="{id}" style="{style}" stroked="f"><v:textbox inset="0,0,0,0">
          <w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent>
        </v:textbox></v:shape>
        """;

    private static string Drawing(bool grouped, bool nested = false)
    {
        var first = Box("alder", "ALDER BOX NORTH", grouped
            ? "position:absolute;left:100;top:200;width:240;height:56"
            : Placement + "margin-left:72pt;margin-top:100pt;width:120pt;height:28pt");
        var second = Box("birch", "BIRCH BOX SOUTH", grouped
            ? nested ? "position:absolute;left:10;top:20;width:240;height:56"
                : "position:absolute;left:400;top:300;width:240;height:56"
            : Placement + "margin-left:222pt;margin-top:150pt;width:120pt;height:28pt");
        if (nested)
            second = $"<v:group style=\"position:absolute;left:400;top:300;width:240;height:100\" coordorigin=\"10,20\" coordsize=\"240,100\">{second}</v:group>";
        return grouped
            ? $"<w:pict><v:group id=\"grove\" style=\"{Placement}margin-left:72pt;margin-top:100pt;width:300pt;height:100pt\" coordorigin=\"100,200\" coordsize=\"600,200\">{first}{second}</v:group></w:pict>"
            : $"<w:pict>{first}</w:pict><w:pict>{second}</w:pict>";
    }

    private static byte[] Document(string drawing, bool header)
    {
        var namespaces = $"xmlns:w=\"{W.w}\" xmlns:v=\"{VML.vml}\" xmlns:r=\"{R.r}\" xmlns:mc=\"{MC.mc}\"";
        var carrier = $"<w:p><w:r>{drawing}</w:r></w:p>";
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(XDocument.Parse($"<w:styles xmlns:w=\"{W.w}\"/>"));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings xmlns:w=\"{W.w}\"/>"));
            OpenXmlPart owner = main;
            if (header)
            {
                var part = main.AddNewPart<HeaderPart>("header");
                part.PutXDocument(XDocument.Parse($"<w:hdr {namespaces}>{carrier}</w:hdr>"));
                owner = part;
            }
            if (drawing.Contains("=\"image\"", StringComparison.Ordinal))
            {
                var image = owner.AddNewPart<ImagePart>("image/png", "image");
                using var pixels = new MemoryStream(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a0HcAAAAASUVORK5CYII="));
                image.FeedData(pixels);
            }
            main.PutXDocument(XDocument.Parse($"""
                <w:document {namespaces}><w:body>
                  <w:p><w:r><w:t>Grove register control</w:t></w:r></w:p>{(header ? "" : carrier)}
                  <w:sectPr>{(header ? "<w:headerReference w:type=\"default\" r:id=\"header\"/>" : "")}
                    <w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:left="1440" w:bottom="1440" w:right="1440" w:header="360" w:footer="360"/>
                  </w:sectPr>
                </w:body></w:document>
                """));
            Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365).Validate(document));
        }
        return stream.ToArray();
    }

    private static XElement Render(string drawing, bool header = false) => XElement.Parse(
        HtmlConversionOps.ConvertToHtml(Document(drawing, header), new HtmlConversionOptions
        {
            FabricateCssClasses = false, PaginationMode = 1, RenderHeadersAndFooters = true,
            RenderUnsupportedContentPlaceholders = true,
        }));

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void BodyAndHeaderKeepBothIndependentStories(bool header, bool grouped)
    {
        var boxes = Render(Drawing(grouped), header).Descendants()
            .Where(e => (string?)e.Attribute("data-docx-textbox") == "true").ToArray();
        Assert.Equal(new[] { "ALDER BOX NORTH", "BIRCH BOX SOUTH" }, boxes.Select(box => box.Value.Trim()));
        Assert.All(boxes, box => Assert.Contains("width: 120pt", (string?)box.Attribute("style")));
        Assert.All(boxes, box => Assert.Contains("height: 28pt", (string?)box.Attribute("style")));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void GroupCoordinatesPreserveSiblingOffsetsAndNestedScale(bool nested)
    {
        var html = Render(Drawing(true, nested));
        var boxes = html.Descendants().Where(e => (string?)e.Attribute("data-docx-textbox") == "true").ToArray();
        Assert.Equal(2, boxes.Length);
        Assert.All(boxes, box => Assert.Contains("width: 120pt", (string?)box.Attribute("style")));
        Assert.Contains("left: 0pt", (string?)boxes[0].Attribute("style"));
        Assert.Contains("top: 0pt", (string?)boxes[0].Attribute("style"));
        var secondFrame = nested ? boxes[1].Parent! : boxes[1];
        Assert.Contains("left: 150pt", (string?)secondFrame.Attribute("style"));
        Assert.Contains("top: 50pt", (string?)secondFrame.Attribute("style"));
        var anchor = Assert.Single(html.Descendants(), e => e.Attribute("data-docx-drawing-anchor") != null);
        Assert.Equal("72", (string?)anchor.Attribute("data-docx-anchor-h-offset"));
        Assert.Equal("100", (string?)anchor.Attribute("data-docx-anchor-v-offset"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OmittedGroupCoordinateSizeUsesTheThousandUnitDefault(bool explicitDefault)
    {
        var drawing = Drawing(true).Replace("coordorigin=\"100,200\"", "coordorigin=\"0,0\"")
            .Replace("coordsize=\"600,200\"", explicitDefault ? "coordsize=\"1000,1000\"" : "")
            .Replace("left:100;top:200;width:240;height:56", "left:0;top:0;width:400;height:280")
            .Replace("left:400;top:300;width:240;height:56", "left:500;top:500;width:400;height:280");
        var boxes = Render(drawing).Descendants().Where(e => e.Attribute("data-docx-textbox") != null).ToArray();
        Assert.Equal(2, boxes.Length);
        Assert.All(boxes, box => Assert.Contains("width: 120pt", (string?)box.Attribute("style")));
        Assert.All(boxes, box => Assert.Contains("height: 28pt", (string?)box.Attribute("style")));
        Assert.Contains("left: 150pt", (string?)boxes[1].Attribute("style"));
        Assert.Contains("top: 50pt", (string?)boxes[1].Attribute("style"));
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(false, true)]
    public void CompatibilitySelectionKeepsOneGroupAndNoFallbackDuplicates(bool supported, bool insidePict)
    {
        var drawing = $"<mc:AlternateContent xmlns:future=\"urn:unsupported\"><mc:Choice Requires=\"{(supported ? "v" : "future")}\">{Drawing(true)}</mc:Choice><mc:Fallback>{Drawing(true).Replace("ALDER BOX NORTH", "FALLBACK NORTH").Replace("BIRCH BOX SOUTH", "FALLBACK SOUTH")}</mc:Fallback></mc:AlternateContent>";
        if (insidePict) drawing = "<w:pict>" + drawing.Replace("<w:pict>", "").Replace("</w:pict>", "") + "</w:pict>";
        var boxes = Render(drawing).Descendants().Where(e => e.Attribute("data-docx-textbox") != null).ToArray();
        Assert.Equal(supported ? new[] { "ALDER BOX NORTH", "BIRCH BOX SOUTH" } : new[] { "FALLBACK NORTH", "FALLBACK SOUTH" },
            boxes.Select(box => box.Value.Trim()));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AGroupedImageDoesNotReplaceEitherTextStory(bool header)
    {
        var drawing = Drawing(true).Replace("</v:group>", "<v:shape style=\"position:absolute;left:100;top:340;width:40;height:40\"><v:imagedata r:id=\"image\"/></v:shape></v:group>");
        var html = Render(drawing, header);
        Assert.Contains("ALDER BOX NORTH", html.Value);
        Assert.Contains("BIRCH BOX SOUTH", html.Value);
        var image = Assert.Single(html.Descendants(), e => e.Name.LocalName == "img");
        Assert.StartsWith("data:image/png;base64,", (string?)image.Attribute("src"));
        Assert.Contains("width: 20pt", (string?)image.Attribute("style"));
        Assert.Contains("left: 0pt", (string?)image.Attribute("style"));
        Assert.Contains("top: 70pt", (string?)image.Attribute("style"));
    }

    [Fact]
    public void SiblingShapesInOnePictRetainIndependentPageAnchors()
    {
        var html = Render(Drawing(false).Replace("</w:pict><w:pict>", ""));
        var boxes = html.Descendants().Where(e => e.Attribute("data-docx-textbox") != null).ToArray();
        Assert.Equal(2, boxes.Length);
        Assert.Equal(new[] { "72", "222" }, boxes.Select(e => (string?)e.Attribute("data-docx-anchor-h-offset")));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DifferentVmlShapeTypesKeepTheirOwnTextboxAndFrame(bool grouped)
    {
        var drawing = Drawing(grouped).Replace("</w:pict><w:pict>", "");
        var start = drawing.IndexOf("<v:shape id=\"birch\"", StringComparison.Ordinal);
        drawing = drawing[..start] + drawing[start..].Replace("v:shape", "v:roundrect");
        var boxes = Render(drawing).Descendants().Where(e => e.Attribute("data-docx-textbox") != null).ToArray();
        Assert.Equal(new[] { "ALDER BOX NORTH", "BIRCH BOX SOUTH" }, boxes.Select(b => b.Value.Trim()));
        Assert.All(boxes, b => Assert.Contains("width: 120pt", (string?)b.Attribute("style")));
        Assert.All(boxes, b => Assert.Contains("height: 28pt", (string?)b.Attribute("style")));
    }

    [Fact]
    public void PointSuffixedChildrenUseTheSameGroupCoordinateMapping()
    {
        var drawing = Drawing(true).Replace("left:100;top:200;width:240;height:56", "left:100pt;top:200pt;width:240pt;height:56pt")
            .Replace("left:400;top:300;width:240;height:56", "left:400pt;top:300pt;width:240pt;height:56pt");
        var boxes = Render(drawing).Descendants().Where(e => e.Attribute("data-docx-textbox") != null).ToArray();
        Assert.Equal(2, boxes.Length);
        Assert.All(boxes, b => Assert.Contains("width: 120pt", (string?)b.Attribute("style")));
        Assert.Contains("left: 0pt", (string?)boxes[0].Attribute("style"));
        Assert.Contains("left: 150pt", (string?)boxes[1].Attribute("style"));
        Assert.Contains("top: 50pt", (string?)boxes[1].Attribute("style"));
    }

    [Theory]
    [InlineData("rotation:90", false)]
    [InlineData("rotation:90", true)]
    [InlineData("flip:x", false)]
    [InlineData("flip:y", true)]
    [InlineData("center-x:300", false)]
    public void UnsupportedGroupTransformsReportLossWithoutTrustingChildClipping(string transform, bool nested)
    {
        var drawing = Drawing(true, nested);
        var marker = nested ? "left:400;top:300;width:240;height:100" : "margin-left:72pt;margin-top:100pt";
        var html = Render(drawing.Replace(marker, marker + ";" + transform));
        var partial = Assert.Single(html.Descendants(), e => (string?)e.Attribute("data-docx-shape") == "partial");
        Assert.Equal("Other", (string?)partial.Attribute("data-content-type"));
        Assert.Contains("VML group", (string?)partial.Attribute("title"));
        Assert.Contains("BIRCH BOX SOUTH", partial.Value);
        Assert.All(partial.Descendants(), e =>
        {
            Assert.Null(e.Attribute("data-docx-textbox-source-clip"));
            Assert.Null(e.Attribute("data-docx-textbox-source-position"));
        });
        if (nested)
            Assert.Contains(html.Descendants(), e => e.Attribute("data-docx-textbox-source-clip") != null && e.Value.Contains("ALDER"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TheHostRetainsAnIndependentInventoryOfBothTextboxParagraphs(bool header)
    {
        var html = Render(Drawing(true), header);
        var host = Assert.Single(html.Descendants(), e => e.Attribute("data-docx-required-source-anchors") != null);
        var required = ((string)host.Attribute("data-docx-required-source-anchors")!).Split(' ');
        Assert.Equal(2, required.Length);
        var paragraphs = html.Descendants().Where(e => e.Attribute("data-source-anchor-id") != null &&
            e.Ancestors().Any(a => a.Attribute("data-docx-textbox") != null));
        Assert.Equal(paragraphs.Select(p => (string?)p.Attribute("data-source-anchor-id")), required);
    }

    [Fact]
    public void EmptyImageMarkupDoesNotSuppressTheTextboxSourceInventory()
    {
        var html = Render(Drawing(true).Replace("<v:textbox", "<v:imagedata/><v:textbox"));
        var host = Assert.Single(html.Descendants(), e => e.Attribute("data-docx-required-source-anchors") != null);
        Assert.Equal(2, ((string)host.Attribute("data-docx-required-source-anchors")!).Split(' ').Length);
        Assert.Contains("BIRCH BOX SOUTH", host.Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImageChoiceInventoryFollowsCallbackFallbackAndResetsOnEveryConversion(bool drawingMl)
    {
        var drawing = drawingMl ? DrawingImageAndText() : "<w:pict>" +
            Box("alternative", "Alternative text story", "width:120pt;height:60pt")
                .Replace("<v:textbox", "<v:imagedata r:id=\"image\"/><v:textbox") + "</w:pict>";
        using var stream = new MemoryStream();
        stream.Write(Document(drawing, false));
        using var document = WordprocessingDocument.Open(stream, true);
        var settings = new WmlToHtmlConverterSettings
        {
            RenderPagination = PaginationMode.Paginated, StampCanonicalSourceAnchors = true,
            ImageHandler = _ => new XElement("img", new XAttribute("src", "chosen")),
        };
        XElement RenderWith(bool image)
        {
            settings.ImageHandler = image ? _ => new XElement("img", new XAttribute("src", "chosen")) : _ => null;
            return WmlToHtmlConverter.ConvertToHtml(document, settings);
        }
        foreach (var image in new[] { true, false, true })
        {
            var html = RenderWith(image);
            Assert.Empty(settings.ImageRepresentations);
            Assert.Equal(image ? 0 : 1, html.Descendants().Count(e => e.Attribute("data-docx-required-source-anchors") != null));
            Assert.Equal(!image, html.Value.Contains("Alternative text story", StringComparison.Ordinal));
            Assert.Equal(image ? 1 : 0, html.Descendants().Count(e => e.Name.LocalName == "img"));
        }
    }

    private static string DrawingImageAndText() => $"""
        <w:drawing xmlns:wp="{WP.wp}" xmlns:a="{A.a}" xmlns:pic="{Pic.pic}" xmlns:wps="{WPS.wps}">
          <wp:inline><wp:extent cx="1524000" cy="762000"/><wp:docPr id="8" name="Image and textbox"/>
            <a:graphic><a:graphicData uri="{Pic.pic}"><pic:pic>
              <pic:nvPicPr><pic:cNvPr id="8" name="image"/><pic:cNvPicPr/></pic:nvPicPr>
              <pic:blipFill><a:blip r:embed="image"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
              <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1524000" cy="762000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
            </pic:pic><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1524000" cy="762000"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent>
              <w:p><w:r><w:t>Alternative text story</w:t></w:r></w:p></w:txbxContent></wps:txbx>
              <wps:bodyPr><a:noAutofit/></wps:bodyPr></wps:wsp>
            </a:graphicData></a:graphic>
          </wp:inline>
        </w:drawing>
        """;
}
