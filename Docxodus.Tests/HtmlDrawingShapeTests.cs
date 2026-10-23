// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// DrawingML shapes and groups that are not text boxes render as inline SVG of their extent, and an
/// inline drawing the converter cannot draw still occupies its extent (issue #1270). Before, such a
/// drawing produced no HTML while the empty-paragraph pass still counted it as content, so its
/// paragraph had no line box and PDF export failed PageMap verification.
/// </summary>
public class HtmlDrawingShapeTests
{
    private const string Namespaces =
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" " +
        "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" " +
        "xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\" " +
        "xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" " +
        "xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" " +
        "xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\"";

    private const string ShapeUri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
    private const string GroupUri = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
    private const string Teal = "<a:solidFill><a:srgbClr val=\"19647E\"/></a:solidFill>";

    /// <summary>A 2 × 1 inch shape, as the issue's reproduction generator writes it.</summary>
    private static string Shape(string geometry = "rect", string fill = Teal, string xfrm = "",
        string extra = "", long cx = 1828800, long cy = 914400) =>
        $"<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm{xfrm}><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm>" +
        $"<a:prstGeom prst=\"{geometry}\"><a:avLst/></a:prstGeom>{fill}{extra}</wps:spPr><wps:bodyPr/></wps:wsp>";

    /// <summary>The issue's group: the same rectangle under an identity child coordinate space.</summary>
    private static string Group(string children, string childOffset = "0", long childCx = 1828800,
        long childCy = 914400) =>
        "<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1828800\" cy=\"914400\"/>" +
        $"<a:chOff x=\"{childOffset}\" y=\"{childOffset}\"/><a:chExt cx=\"{childCx}\" cy=\"{childCy}\"/></a:xfrm></wpg:grpSpPr>" +
        children + "</wpg:wgp>";

    private static string InlineDrawing(string uri, string data) =>
        "<w:r><w:drawing><wp:inline><wp:extent cx=\"1828800\" cy=\"914400\"/><wp:docPr id=\"7\" name=\"Teal inspection panel\"/>" +
        $"<a:graphic><a:graphicData uri=\"{uri}\">{data}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";

    private static string FloatingDrawing(string uri, string data) =>
        "<w:r><w:drawing><wp:anchor distT=\"0\" distR=\"114300\" distB=\"0\" distL=\"114300\" simplePos=\"0\" " +
        "relativeHeight=\"10\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\"><wp:simplePos x=\"0\" y=\"0\"/>" +
        "<wp:positionH relativeFrom=\"page\"><wp:posOffset>4572000</wp:posOffset></wp:positionH>" +
        "<wp:positionV relativeFrom=\"page\"><wp:posOffset>4572000</wp:posOffset></wp:positionV>" +
        "<wp:extent cx=\"1828800\" cy=\"914400\"/><wp:wrapNone/><wp:docPr id=\"7\" name=\"Teal inspection panel\"/>" +
        $"<a:graphic><a:graphicData uri=\"{uri}\">{data}</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>";

    private static string Deleted(string runs) =>
        $"<w:del w:id=\"31\" w:author=\"Review Example\" w:date=\"2026-09-12T10:00:00Z\">{runs}</w:del>";

    private static byte[] Docx(string middleParagraphContent)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            using (var writer = new StreamWriter(main.GetStream(FileMode.Create, FileAccess.Write)))
                writer.Write(
                    $"<w:document {Namespaces}><w:body>" +
                    "<w:p><w:r><w:t>Optical panel</w:t></w:r></w:p>" +
                    $"<w:p>{middleParagraphContent}</w:p>" +
                    "<w:p><w:r><w:t>End of inspection</w:t></w:r></w:p><w:sectPr/></w:body></w:document>");
            main.AddNewPart<StyleDefinitionsPart>().Styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();
            main.AddNewPart<DocumentSettingsPart>().Settings = new DocumentFormat.OpenXml.Wordprocessing.Settings();
            doc.Save();
        }
        return stream.ToArray();
    }

    private static XElement MiddleParagraph(byte[] docx, PaginationMode mode = PaginationMode.None)
    {
        var html = HtmlConversionOps.ConvertToHtml(docx, new HtmlConversionOptions
        {
            FabricateCssClasses = false,
            PaginationMode = (int)mode,
            RenderTrackedChanges = true,
        });
        return XElement.Parse(html).Descendants().Where(e => e.Name.LocalName == "p")
            .Single(p => !p.Value.Contains("Optical panel") && !p.Value.Contains("End of inspection"));
    }

    private static XElement OnlySvg(XElement paragraph) =>
        paragraph.Descendants().Single(e => e.Name.LocalName == "svg");

    private static XElement[] Drawn(XElement svg, string localName) =>
        svg.Descendants().Where(e => e.Name.LocalName == localName).ToArray();

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void IssueShapeOrGroup_RendersTheTealRectangleAtItsExtent(bool grouped, bool deleted)
    {
        var drawing = grouped ? InlineDrawing(GroupUri, Group(Shape())) : InlineDrawing(ShapeUri, Shape());
        var paragraph = MiddleParagraph(Docx(deleted ? Deleted(drawing) : drawing));

        var svg = OnlySvg(paragraph);
        var box = (string?)svg.Parent!.Attribute("style");
        Assert.Contains("width: 144pt", box);
        Assert.Contains("height: 72pt", box);
        Assert.Equal(("192", "96"), ((string?)svg.Attribute("width"), (string?)svg.Attribute("height")));
        var rect = Assert.Single(Drawn(svg, "rect"));
        Assert.Equal("#19647E", (string?)rect.Attribute("fill"));
        Assert.Equal(deleted, svg.Ancestors().Any(a => a.Name.LocalName == "del"));
    }

    [Fact]
    public void GroupChildren_AreMappedThroughTheChildCoordinateSpace()
    {
        // A child space half the group's size, offset by 1000: the child's frame doubles on render.
        var child = "<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x=\"1000\" y=\"1000\"/><a:ext cx=\"457200\" cy=\"457200\"/></a:xfrm>" +
                    $"<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>{Teal}</wps:spPr><wps:bodyPr/></wps:wsp>";
        var svg = OnlySvg(MiddleParagraph(Docx(InlineDrawing(GroupUri, Group(child, "1000", 914400, 457200)))));

        // The group maps child (1000, 1000)–(915400, 458400) onto its frame (0, 0)–(1828800, 914400).
        var group = Assert.Single(Drawn(svg, "g"), g => g.Elements().Any(e => e.Name.LocalName == "rect"));
        Assert.Equal("matrix(2 0 0 2 -2000 -2000)", (string?)group.Attribute("transform"));
        var rect = Assert.Single(Drawn(group, "rect"));
        Assert.Equal("1000", (string?)rect.Attribute("x"));
        Assert.Equal("457200", (string?)rect.Attribute("width"));
    }

    [Fact]
    public void CanvasChildren_ArePlacedInTheDrawingsOwnCoordinates()
    {
        var child = Shape(cx: 457200, cy: 457200).Replace("<a:off x=\"0\" y=\"0\"/>", "<a:off x=\"914400\" y=\"0\"/>");
        var svg = OnlySvg(MiddleParagraph(Docx(InlineDrawing(
            "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas", $"<wpc:wpc><wpc:bg/><wpc:whole/>{child}</wpc:wpc>"))));

        var rect = Assert.Single(Drawn(svg, "rect"));
        Assert.Equal("914400", (string?)rect.Attribute("x"));
        Assert.Equal("457200", (string?)rect.Attribute("width"));
    }

    [Fact]
    public void Ellipse_RendersAsAnEllipseFillingItsFrame()
    {
        var ellipse = Assert.Single(Drawn(OnlySvg(MiddleParagraph(Docx(InlineDrawing(ShapeUri, Shape("ellipse"))))), "ellipse"));
        Assert.Equal("914400", (string?)ellipse.Attribute("rx"));
        Assert.Equal("457200", (string?)ellipse.Attribute("ry"));
        Assert.Equal("#19647E", (string?)ellipse.Attribute("fill"));
    }

    [Fact]
    public void RoundedRectangle_TakesItsCornerRadiusFromTheShorterSide()
    {
        var rect = Assert.Single(Drawn(OnlySvg(MiddleParagraph(Docx(InlineDrawing(ShapeUri, Shape("roundRect"))))), "rect"));
        // Default adj 16667 of the 914400 EMU short side.
        Assert.Equal("152403.048", (string?)rect.Attribute("rx"));
    }

    [Fact]
    public void FlippedLine_RunsCornerToCornerWithItsOutline()
    {
        var outline = "<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill></a:ln>";
        var line = Assert.Single(Drawn(OnlySvg(MiddleParagraph(Docx(InlineDrawing(ShapeUri,
            Shape("line", fill: "", xfrm: " flipH=\"1\"", extra: outline))))), "line"));
        Assert.Equal("none", (string?)line.Attribute("fill"));
        Assert.Equal("#FF0000", (string?)line.Attribute("stroke"));
        Assert.Equal("25400", (string?)line.Attribute("stroke-width"));
        Assert.Equal("matrix(-1 0 0 1 1828800 0)", (string?)line.Attribute("transform"));
    }

    [Fact]
    public void HorizontalLine_WithAZeroHeightExtent_IsStillDrawn()
    {
        var outline = "<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln>";
        var drawing = InlineDrawing(ShapeUri, Shape("line", fill: "", extra: outline, cy: 0)).Replace("cy=\"914400\"", "cy=\"0\"");
        var svg = OnlySvg(MiddleParagraph(Docx(drawing)));

        Assert.Contains("height: 0pt", (string?)svg.Parent!.Attribute("style"));
        Assert.Null(svg.Attribute("viewBox"));
        // A degenerate viewBox would disable rendering; the export needs a positive size to lay it out.
        Assert.Equal("1", (string?)svg.Attribute("height"));
        var line = Assert.Single(Drawn(svg, "line"));
        Assert.Equal("0", (string?)line.Attribute("y2"));
        Assert.Equal("1828800", (string?)line.Attribute("x2"));
    }

    [Fact]
    public void StyleReferences_ResolveThemeFillAndOutline()
    {
        // Word's default shape style: accent1 fill, and an outline of accent1 at 50% shade.
        var styled = Shape(fill: "").Replace("<wps:bodyPr/>",
            "<wps:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"50000\"/></a:schemeClr></a:lnRef>" +
            "<a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef><a:effectRef idx=\"0\"><a:scrgbClr r=\"0\" g=\"0\" b=\"0\"/></a:effectRef>" +
            "<a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef></wps:style><wps:bodyPr/>");
        var bytes = Docx(InlineDrawing(ShapeUri, styled));
        using (var stream = new MemoryStream())
        {
            stream.Write(bytes);
            using (var doc = WordprocessingDocument.Open(stream, true))
            {
                var theme = doc.MainDocumentPart!.AddNewPart<ThemePart>();
                using var writer = new StreamWriter(theme.GetStream(FileMode.Create, FileAccess.Write));
                writer.Write("<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"T\"><a:themeElements>" +
                             "<a:clrScheme name=\"T\"><a:accent1><a:srgbClr val=\"4472C4\"/></a:accent1></a:clrScheme>" +
                             "<a:fmtScheme name=\"T\"><a:lnStyleLst><a:ln w=\"6350\"/><a:ln w=\"12700\"/><a:ln w=\"19050\"/></a:lnStyleLst></a:fmtScheme>" +
                             "</a:themeElements></a:theme>");
            }
            bytes = stream.ToArray();
        }

        var rect = Assert.Single(Drawn(OnlySvg(MiddleParagraph(bytes)), "rect"));
        Assert.Equal("#4472C4", (string?)rect.Attribute("fill"));
        Assert.Equal("#2F528F", (string?)rect.Attribute("stroke"));
        Assert.Equal("12700", (string?)rect.Attribute("stroke-width"));
    }

    [Fact]
    public void UndrawnGeometry_StillReservesItsExtent()
    {
        var svg = OnlySvg(MiddleParagraph(Docx(InlineDrawing(ShapeUri, Shape("cloudCallout")))));
        Assert.Contains("width: 144pt", (string?)svg.Parent!.Attribute("style"));
        Assert.All(svg.Descendants(), e => Assert.Equal("g", e.Name.LocalName));
    }

    [Fact]
    public void InlineDrawingWithUnrecognizedContent_ReservesAnEmptyBoxOfItsExtent()
    {
        var svg = OnlySvg(MiddleParagraph(Docx(InlineDrawing("urn:example:unknown", "<unknown xmlns=\"urn:example:unknown\"/>"))));
        Assert.Equal("unrendered", (string?)svg.Attribute("data-docx-shape"));
        Assert.Contains("height: 72pt", (string?)svg.Parent!.Attribute("style"));
    }

    /// <summary>An inline picture whose image relationship does not exist.</summary>
    private static readonly string MissingPicture = InlineDrawing("http://schemas.openxmlformats.org/drawingml/2006/picture",
        "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"1\" name=\"p\"/><pic:cNvPicPr/></pic:nvPicPr>" +
        "<pic:blipFill><a:blip r:embed=\"rIdMissing\"/></pic:blipFill><pic:spPr/></pic:pic>");

    [Fact]
    public void PictureWhoseImageIsMissing_ReservesAnEmptyBoxOfItsExtent()
    {
        var svg = OnlySvg(MiddleParagraph(Docx(MissingPicture)));
        Assert.Equal("unrendered", (string?)svg.Attribute("data-docx-shape"));
        Assert.Contains("width: 144pt", (string?)svg.Parent!.Attribute("style"));
    }

    [Fact]
    public void PictureWithNoImageHandler_StaysOmitted()
    {
        // The caller passing no image handler has asked for pictures to be left out.
        var html = WmlToHtmlConverter.ConvertToHtml(new WmlDocument("p.docx", Docx(MissingPicture)),
            new WmlToHtmlConverterSettings { FabricateCssClasses = false });
        Assert.DoesNotContain(html.Descendants(), e => e.Name.LocalName == "svg");
    }

    [Theory]
    [InlineData(PaginationMode.Paginated)]
    [InlineData(PaginationMode.None)]
    public void FloatingShape_StaysOutOfTheTextFlowAndItsParagraphKeepsAPlaceholderLine(PaginationMode mode)
    {
        // A floating drawing is not paragraph content (issue #880), so nothing may take its line.
        var paragraph = MiddleParagraph(Docx(FloatingDrawing(ShapeUri, Shape())), mode);

        Assert.DoesNotContain(paragraph.Descendants(), e => e.Name.LocalName == "svg");
        Assert.NotEmpty(paragraph.Value);
        Assert.True(string.IsNullOrWhiteSpace(paragraph.Value));
    }
}
