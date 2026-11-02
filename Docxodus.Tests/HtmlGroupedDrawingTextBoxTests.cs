using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

public class HtmlGroupedDrawingTextBoxTests
{
    private static readonly string Namespaces = $"xmlns:w=\"{W.w}\" xmlns:r=\"{R.r}\" xmlns:wp=\"{WP.wp}\" " +
        $"xmlns:a=\"{A.a}\" xmlns:wps=\"{WPS.wps}\" xmlns:wpg=\"{WPG.wpg}\" xmlns:pic=\"{Pic.pic}\"";
    private const string SignalPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a0HcAAAAASUVORK5CYII=";
    private const string BodyControlPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGO4I2cEAAMGAS2aHkTAAAAAAElFTkSuQmCC";
    private static readonly string[] Labels = { "CEDAR WINDOW", "MOSS WINDOW", "SKY WINDOW" };

    private static string Transform(int x, int y, int width, int height, string children = "") =>
        $"<a:xfrm><a:off x=\"{x * 12700}\" y=\"{y * 12700}\"/><a:ext cx=\"{width * 12700}\" cy=\"{height * 12700}\"/>{children}</a:xfrm>";

    private static string Box(string text, int x = 0, int y = 0, int width = 100, int height = 30) => $"""
        <wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>{Transform(x, y, width, height)}
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>
        </wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>
        <wps:bodyPr rot="0" lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></wps:bodyPr></wps:wsp>
        """;

    private static string Group(string members, bool nested = false, int x = 0, int y = 0,
        int width = 360, int height = 80, int childX = 0, int childY = 0, int childWidth = 360, int childHeight = 80) =>
        $"<wpg:{(nested ? "grpSp" : "wgp")}><wpg:cNvGrpSpPr/><wpg:grpSpPr>" +
        Transform(x, y, width, height, $"<a:chOff x=\"{childX * 12700}\" y=\"{childY * 12700}\"/><a:chExt cx=\"{childWidth * 12700}\" cy=\"{childHeight * 12700}\"/>") +
        $"</wpg:grpSpPr>{members}</wpg:{(nested ? "grpSp" : "wgp")}>";

    private static string Drawing(string members, bool grouped = true, bool inserted = false)
    {
        var run = $"""
            <w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0"
              relativeHeight="0" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
              <wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>
              <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
              <wp:extent cx="{(grouped ? 360 : 100) * 12700}" cy="{(grouped ? 80 : 30) * 12700}"/>
              <wp:wrapNone/><wp:docPr id="1" name="Window group"/>
              <a:graphic><a:graphicData uri="{(grouped ? WPG.wpg : WPS.wps)}">{members}</a:graphicData></a:graphic>
            </wp:anchor></w:drawing></w:r>
            """;
        return inserted ? $"<w:ins w:id=\"11\" w:author=\"Fixture Author\" w:date=\"2026-02-03T10:00:00Z\">{run}</w:ins>" : run;
    }

    private static byte[] Document(string drawing, bool header)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<StyleDefinitionsPart>().PutXDocument(XDocument.Parse($"<w:styles xmlns:w=\"{W.w}\"/>"));
            main.AddNewPart<DocumentSettingsPart>().PutXDocument(XDocument.Parse($"<w:settings xmlns:w=\"{W.w}\"/>"));
            OpenXmlPart owner = main;
            var paragraph = $"<w:p>{drawing}</w:p>";
            if (header)
            {
                var part = main.AddNewPart<HeaderPart>("header");
                part.PutXDocument(XDocument.Parse($"<w:hdr {Namespaces}>{paragraph}</w:hdr>"));
                owner = part;
            }
            if (drawing.Contains("r:embed=\"image\"", StringComparison.Ordinal))
            {
                var image = owner.AddNewPart<ImagePart>("image/png", "image");
                using var pixels = new MemoryStream(Convert.FromBase64String(SignalPng));
                image.FeedData(pixels);
                if (header)
                {
                    var bodyImage = main.AddNewPart<ImagePart>("image/png", "image");
                    using var bodyPixels = new MemoryStream(Convert.FromBase64String(BodyControlPng));
                    bodyImage.FeedData(bodyPixels);
                }
            }
            main.PutXDocument(XDocument.Parse($"""
                <w:document {Namespaces}><w:body><w:p><w:r><w:t>Harbor maintenance ledger.</w:t></w:r></w:p>
                  {(header ? "" : paragraph)}<w:sectPr>
                    {(header ? "<w:headerReference w:type=\"default\" r:id=\"header\"/>" : "")}
                    <w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="360" w:footer="360"/>
                  </w:sectPr></w:body></w:document>
                """));
            Assert.Empty(new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365).Validate(document));
        }
        return stream.ToArray();
    }

    private static XElement Render(string drawing, bool header) => XElement.Parse(
        HtmlConversionOps.ConvertToHtml(Document(drawing, header), new HtmlConversionOptions
        {
            FabricateCssClasses = false, PaginationMode = 1, RenderHeadersAndFooters = true,
            RenderTrackedChanges = true, RevisionPresentation = 1, RenderUnsupportedContentPlaceholders = true,
        }));

    private static XElement[] Boxes(XElement html) => html.Descendants()
        .Where(element => (string?)element.Attribute("data-docx-textbox") == "true").ToArray();

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void EveryIndependentStorySurvivesCleanAndInsertedGroups(bool header, bool inserted)
    {
        var html = Render(Drawing(Group(string.Concat(Labels.Select((text, index) => Box(text, index * 120)))), inserted: inserted), header);
        var boxes = Boxes(html);
        Assert.Equal(Labels, boxes.Select(box => box.Value.Trim()));
        Assert.All(boxes, box => Assert.Contains("width: 100pt", (string?)box.Attribute("style")));
        Assert.All(boxes, box => Assert.Contains("height: 30pt", (string?)box.Attribute("style")));
        Assert.Single(html.Descendants(), element => element.Attribute("data-docx-drawing-anchor") != null);
        var host = Assert.Single(html.Descendants(), element => element.Attribute("data-docx-required-source-anchors") != null);
        var required = ((string)host.Attribute("data-docx-required-source-anchors")!).Split(' ');
        Assert.Equal(3, required.Length);
        Assert.Equal(required, boxes.SelectMany(box => box.Descendants().Attributes("data-source-anchor-id")).Select(a => a.Value));
        if (inserted) Assert.All(boxes, box => Assert.Contains(box.Ancestors(), e => e.Name.LocalName == "ins"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ChildOffsetsAndNestedScalePreserveSeparateFrames(bool nested)
    {
        var second = nested
            ? Group(Box(Labels[1], 10, 20, 200, 60), true, 120, 20, 100, 30, 10, 20, 200, 60)
            : Box(Labels[1], 120, 20);
        var boxes = Boxes(Render(Drawing(Group(Box(Labels[0]) + second + Box(Labels[2], 240))), true));
        Assert.Equal(3, boxes.Length);
        Assert.All(boxes, box => Assert.Contains("width: 100pt", (string?)box.Attribute("style")));
        Assert.Contains("left: 0pt", (string?)boxes[0].Parent!.Attribute("style"));
        Assert.Contains("left: 240pt", (string?)boxes[2].Parent!.Attribute("style"));
        var secondFrame = nested ? boxes[1].Parent!.Parent! : boxes[1].Parent!;
        Assert.Contains("left: 120pt", (string?)secondFrame.Attribute("style"));
        Assert.Contains("top: 20pt", (string?)secondFrame.Attribute("style"));
        if (nested) Assert.Contains("left: 0pt", (string?)boxes[1].Parent!.Attribute("style"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AGroupImageKeepsItsOwningPartAndAllTextStories(bool header)
    {
        var picture = $"""
            <pic:pic><pic:nvPicPr><pic:cNvPr id="2" name="Signal pixel"/><pic:cNvPicPr/></pic:nvPicPr>
              <pic:blipFill><a:blip r:embed="image"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
              <pic:spPr>{Transform(240, 40, 20, 20)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>
            """;
        var html = Render(Drawing(Group(Box(Labels[0]) + picture + Box(Labels[1], 120))), header);
        Assert.Equal(Labels.Take(2), Boxes(html).Select(box => box.Value.Trim()));
        var image = Assert.Single(html.Descendants(), element => element.Name.LocalName == "img");
        Assert.Equal("data:image/png;base64," + SignalPng, (string?)image.Attribute("src"));
        Assert.Contains("width: 20pt", (string?)image.Attribute("style"));
        Assert.Contains("left: 240pt", (string?)image.Attribute("style"));
        Assert.Contains("top: 40pt", (string?)image.Attribute("style"));
        var host = Assert.Single(html.Descendants(), element => element.Attribute("data-docx-required-source-anchors") != null);
        Assert.Equal(2, ((string)host.Attribute("data-docx-required-source-anchors")!).Split(' ').Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SingleTextboxControlRetainsItsSizeAndText(bool header)
    {
        var box = Assert.Single(Boxes(Render(Drawing(Box(Labels[0]), grouped: false), header)));
        Assert.Equal(Labels[0], box.Value.Trim());
        Assert.Contains("width: 100pt", (string?)box.Attribute("style"));
        Assert.Contains("height: 30pt", (string?)box.Attribute("style"));
    }

    [Fact]
    public void UnsupportedGeometryRetainsEveryStoryAndItsDiagnostic()
    {
        var html = Render(Drawing(Group(Box(Labels[0]) + Box(Labels[1], 120).Replace("prst=\"rect\"", "prst=\"cloud\""))), true);
        Assert.Equal(Labels.Take(2), Boxes(html).Select(box => box.Value.Trim()));
        Assert.Contains(html.Descendants(), e => (string?)e.Attribute("data-content-type") == "Other");
        Assert.All(Boxes(html), box => Assert.Null(box.Attribute("data-docx-textbox-source-clip")));
    }

    [Fact]
    public void LinkedContinuationMembersRetainTheirStoryAndFrames()
    {
        var first = Box(Labels[0]).Replace("<wps:txbx>", "<wps:txbx id=\"7\">");
        var last = XElement.Parse($"<root {Namespaces}>{Box(Labels[1], 120)}</root>").Elements().Single();
        last.Element(WPS.txbx)!.ReplaceWith(new XElement(WPS.wps + "linkedTxbx", new XAttribute("id", 7), new XAttribute("seq", 1)));
        last.Element(WPS.bodyPr)!.SetAttributeValue("lIns", 25400);
        var boxes = Boxes(Render(Drawing(Group(first + last)), true));
        Assert.Equal(2, boxes.Length);
        Assert.Equal("0", (string?)boxes[0].Attribute("data-docx-textbox-seq"));
        Assert.Equal("1", (string?)boxes[1].Attribute("data-docx-textbox-seq"));
        Assert.Equal((string?)boxes[0].Attribute("data-docx-textbox-story"), (string?)boxes[1].Attribute("data-docx-textbox-story"));
        Assert.All(boxes, box => Assert.Equal("true", (string?)box.Attribute("data-docx-textbox-repeat")));
        Assert.Contains("padding-left: 2pt", (string?)boxes[1].Attribute("style"));
    }
}
